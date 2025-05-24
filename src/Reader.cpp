/**
 * @file Reader.cpp
 * @brief 高性能Excel读取器实现 - 使用minizip-ng和expat替代xlsxio
 */

#include "TinaXlsx/Reader.hpp"

// 首先包含标准库头文件
#include <memory>
#include <vector>
#include <string>
#include <algorithm>
#include <charconv>
#include <filesystem>
#include <fstream>
#include <sstream>
#include <map>
#include <locale>
#include <codecvt>

// 然后包含minizip-ng头文件，按正确顺序
#include <mz.h>
#include <mz_strm.h>
#include <mz_strm_mem.h>
#include <mz_zip.h>
#include <mz_zip_rw.h>

// 最后包含expat
#include <expat.h>

namespace TinaXlsx {

// Excel XML解析状态枚举
enum class ParseState {
    None,
    Workbook,
    SheetData,
    Row,
    Cell,
    Value,
    InlineString,
    SharedString
};

// 单元格数据类型
enum class ExcelCellType {
    String,
    Number,
    Boolean,
    Date,
    SharedString,
    InlineString
};

// 工作表信息
struct SheetInfo {
    std::string name;
    std::string relationId;
    std::string filePath;
    RowIndex sheetId;
};

// XML解析上下文
struct XmlParseContext {
    ParseState state = ParseState::None;
    ParseState previousState = ParseState::None;
    std::string currentValue;
    ExcelCellType currentCellType = ExcelCellType::String;
    CellPosition currentPosition{0, 0};
    RowIndex currentRow = 0;
    ColumnIndex currentColumn = 0;
    std::vector<std::string>* sharedStrings = nullptr;
    RowData* currentRowData = nullptr;
    bool skipEmptyRows = true;
    ColumnIndex maxColumns = 0;
    
    void reset() {
        state = ParseState::None;
        previousState = ParseState::None;
        currentValue.clear();
        currentCellType = ExcelCellType::String;
        currentPosition = {0, 0};
        currentRow = 0;
        currentColumn = 0;
    }
};

struct Reader::Impl {
    // minizip-ng 相关
    void* zipHandle = nullptr;
    void* zipStream = nullptr;
    std::vector<uint8_t> fileBuffer; // 用于从文件或内存加载数据
    
    // Excel 文件结构
    std::string workbookPath;
    std::vector<SheetInfo> sheets;
    std::vector<std::string> sharedStrings;
    std::map<std::string, std::string> relationships; // relationId -> target path
    
    // 当前工作表
    SheetInfo* currentSheet = nullptr;
    XmlParseContext parseContext;
    RowIndex currentRowIndex = 0;
    bool atEnd = false;
    std::string filePath;
    
    // 缓存的数据
    std::optional<std::pair<RowIndex, ColumnIndex>> cachedDimensions;
    
    // 流式解析状态
    XML_Parser xmlParser = nullptr;
    TableData cachedTableData; // 缓存的表格数据
    bool tableDataCached = false; // 标记是否已缓存数据
    
    ~Impl() {
        cleanup();
    }
    
    void cleanup() {
        if (xmlParser) {
            XML_ParserFree(xmlParser);
            xmlParser = nullptr;
        }
        
        if (zipHandle) {
            mz_zip_reader_close(zipHandle);
            mz_zip_reader_delete(&zipHandle);
            zipHandle = nullptr;
        }
        
        if (zipStream) {
            mz_stream_close(zipStream);
            mz_stream_mem_delete(&zipStream);
            zipStream = nullptr;
        }
        
        fileBuffer.clear();
        sheets.clear();
        sharedStrings.clear();
        relationships.clear();
        currentSheet = nullptr;
        atEnd = false;
        currentRowIndex = 0;
        cachedDimensions.reset();
        cachedTableData.clear();
        tableDataCached = false;
    }
    
    bool openFile(const std::string& path) {
        cleanup();
        filePath = path;
        
        // 检查文件是否存在
        if (!std::filesystem::exists(path)) {
            throw FileException("File not found: " + path);
        }
        
        // 读取文件到内存
        std::ifstream file(path, std::ios::binary | std::ios::ate);
        if (!file.is_open()) {
            throw FileException("Cannot open file: " + path);
        }
        
        auto fileSize = file.tellg();
        file.seekg(0, std::ios::beg);
        
        fileBuffer.resize(static_cast<size_t>(fileSize));
        if (!file.read(reinterpret_cast<char*>(fileBuffer.data()), fileSize)) {
            throw FileException("Failed to read file: " + path);
        }
        
        return openFromMemory(fileBuffer.data(), fileBuffer.size());
    }
    
    bool openFromMemory(const void* data, size_t dataSize) {
        cleanup();
        
        // 创建内存流
        zipStream = mz_stream_mem_create();
        mz_stream_mem_set_buffer(zipStream, const_cast<void*>(data), static_cast<int32_t>(dataSize));
        
        if (mz_stream_open(zipStream, nullptr, MZ_OPEN_MODE_READ) != MZ_OK) {
            throw FileException("Failed to open memory stream");
        }
        
        // 创建ZIP读取器
        zipHandle = mz_zip_reader_create();
        
        if (mz_zip_reader_open(zipHandle, zipStream) != MZ_OK) {
            throw FileException("Failed to open ZIP archive");
        }
        
        // 解析Excel文件结构
        parseWorkbookStructure();
        
        return true;
    }
    
    void parseWorkbookStructure() {
        // 1. 解析内容类型
        parseContentTypes();
        
        // 2. 解析工作簿关系
        parseWorkbookRelations();
        
        // 3. 解析工作簿文件，获取工作表信息
        parseWorkbook();
        
        // 4. 加载共享字符串
        loadSharedStrings();
    }
    
    void parseContentTypes() {
        std::string content = readZipEntry("[Content_Types].xml");
        if (content.empty()) {
            throw FileException("Invalid Excel file: missing [Content_Types].xml");
        }
        
        // 简单解析，查找工作簿的主文件
        size_t pos = content.find("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
        if (pos == std::string::npos) {
            // 尝试宏启用的工作簿
            pos = content.find("application/vnd.ms-excel.sheet.macroEnabled.main+xml");
        }
        
        if (pos != std::string::npos) {
            // 向前查找PartName属性
            size_t partNameStart = content.rfind("PartName=\"", pos);
            if (partNameStart != std::string::npos) {
                partNameStart += 10; // strlen("PartName=\"")
                size_t partNameEnd = content.find("\"", partNameStart);
                if (partNameEnd != std::string::npos) {
                    workbookPath = content.substr(partNameStart, partNameEnd - partNameStart);
                    // 移除开头的斜杠
                    if (!workbookPath.empty() && workbookPath[0] == '/') {
                        workbookPath = workbookPath.substr(1);
                    }
                }
            }
        }
        
        if (workbookPath.empty()) {
            workbookPath = "xl/workbook.xml"; // 默认路径
        }
    }
    
    void parseWorkbookRelations() {
        std::string relsPath = "xl/_rels/workbook.xml.rels";
        std::string content = readZipEntry(relsPath);
        if (content.empty()) {
            return; // 关系文件可能不存在
        }
        
        // 使用expat解析关系文件
        XML_Parser parser = XML_ParserCreate(nullptr);
        XML_SetUserData(parser, this);
        XML_SetElementHandler(parser, 
            [](void* userData, const XML_Char* name, const XML_Char** atts) {
                auto* impl = static_cast<Impl*>(userData);
                impl->parseRelationshipElement(name, atts);
            }, nullptr);
        
        if (XML_Parse(parser, content.c_str(), static_cast<int>(content.length()), XML_TRUE) == XML_STATUS_ERROR) {
            XML_ParserFree(parser);
            throw FileException("Failed to parse workbook relations");
        }
        
        XML_ParserFree(parser);
    }
    
    void parseRelationshipElement(const XML_Char* name, const XML_Char** atts) {
        if (strcmp(name, "Relationship") == 0) {
            std::string id, target, type;
            
            for (int i = 0; atts[i]; i += 2) {
                if (strcmp(atts[i], "Id") == 0) {
                    id = atts[i + 1];
                } else if (strcmp(atts[i], "Target") == 0) {
                    target = atts[i + 1];
                } else if (strcmp(atts[i], "Type") == 0) {
                    type = atts[i + 1];
                }
            }
            
            if (!id.empty() && !target.empty()) {
                // 如果是相对路径，添加xl/前缀
                if (!target.empty() && target[0] != '/') {
                    target = "xl/" + target;
                }
                relationships[id] = target;
            }
        }
    }
    
    void parseWorkbook() {
        std::string content = readZipEntry(workbookPath);
        if (content.empty()) {
            throw FileException("Cannot read workbook.xml");
        }
        
        // 解析工作表信息
        XML_Parser parser = XML_ParserCreate(nullptr);
        XML_SetUserData(parser, this);
        XML_SetElementHandler(parser,
            [](void* userData, const XML_Char* name, const XML_Char** atts) {
                auto* impl = static_cast<Impl*>(userData);
                impl->parseWorkbookElement(name, atts);
            }, nullptr);
        
        if (XML_Parse(parser, content.c_str(), static_cast<int>(content.length()), XML_TRUE) == XML_STATUS_ERROR) {
            XML_ParserFree(parser);
            throw FileException("Failed to parse workbook.xml");
        }
        
        XML_ParserFree(parser);
    }
    
    void parseWorkbookElement(const XML_Char* name, const XML_Char** atts) {
        if (strcmp(name, "sheet") == 0) {
            SheetInfo sheet;
            
            for (int i = 0; atts[i]; i += 2) {
                if (strcmp(atts[i], "name") == 0) {
                    sheet.name = atts[i + 1];
                } else if (strcmp(atts[i], "r:id") == 0) {
                    sheet.relationId = atts[i + 1];
                } else if (strcmp(atts[i], "sheetId") == 0) {
                    sheet.sheetId = std::strtoul(atts[i + 1], nullptr, 10);
                }
            }
            
            // 查找对应的文件路径
            auto it = relationships.find(sheet.relationId);
            if (it != relationships.end()) {
                sheet.filePath = it->second;
                sheets.push_back(sheet);
            }
        }
    }
    
    void loadSharedStrings() {
        // 查找共享字符串文件
        std::string sharedStringsPath;
        for (const auto& [id, path] : relationships) {
            if (path.find("sharedStrings.xml") != std::string::npos) {
                sharedStringsPath = path;
                break;
            }
        }
        
        if (sharedStringsPath.empty()) {
            return; // 没有共享字符串
        }
        
        std::string content = readZipEntry(sharedStringsPath);
        if (content.empty()) {
            return;
        }
        
        // 解析共享字符串
        XML_Parser parser = XML_ParserCreate(nullptr);
        
        // 设置用户数据为this指针
        XML_SetUserData(parser, this);
        
        XML_SetElementHandler(parser,
            [](void* userData, const XML_Char* name, const XML_Char** atts) {
                auto* impl = static_cast<Impl*>(userData);
                if (strcmp(name, "si") == 0) {
                    impl->parseContext.state = ParseState::SharedString;
                } else if (strcmp(name, "t") == 0 && impl->parseContext.state == ParseState::SharedString) {
                    impl->parseContext.state = ParseState::Value;
                }
            },
            [](void* userData, const XML_Char* name) {
                auto* impl = static_cast<Impl*>(userData);
                if (strcmp(name, "si") == 0) {
                    impl->parseContext.state = ParseState::None;
                } else if (strcmp(name, "t") == 0) {
                    impl->parseContext.state = ParseState::SharedString;
                }
            });
        
        XML_SetCharacterDataHandler(parser,
            [](void* userData, const XML_Char* s, int len) {
                auto* impl = static_cast<Impl*>(userData);
                if (impl->parseContext.state == ParseState::Value) {
                    impl->sharedStrings.emplace_back(s, len);
                    impl->parseContext.state = ParseState::SharedString;
                }
            });
        
        if (XML_Parse(parser, content.c_str(), static_cast<int>(content.length()), XML_TRUE) == XML_STATUS_ERROR) {
            XML_ParserFree(parser);
            throw FileException("Failed to parse shared strings");
        }
        
        XML_ParserFree(parser);
    }
    
    std::string readZipEntry(const std::string& entryName) {
        if (!zipHandle) {
            return {};
        }
        
        // 查找文件
        if (mz_zip_reader_locate_entry(zipHandle, entryName.c_str(), 0) != MZ_OK) {
            return {};
        }
        
        // 获取文件信息
        mz_zip_file* file_info = nullptr;
        if (mz_zip_reader_entry_get_info(zipHandle, &file_info) != MZ_OK) {
            return {};
        }
        
        // 打开文件
        if (mz_zip_reader_entry_open(zipHandle) != MZ_OK) {
            return {};
        }
        
        // 读取文件内容
        std::string content;
        content.resize(file_info->uncompressed_size);
        
        int32_t bytes_read = mz_zip_reader_entry_read(zipHandle, 
            content.data(), static_cast<int32_t>(content.size()));
        
        mz_zip_reader_entry_close(zipHandle);
        
        if (bytes_read != static_cast<int32_t>(file_info->uncompressed_size)) {
            return {};
        }
        
        return content;
    }
    
    bool openSheet(const std::string& sheetName) {
        for (auto& sheet : sheets) {
            if (sheet.name == sheetName) {
                currentSheet = &sheet;
                resetSheet();
                return true;
            }
        }
        return false;
    }
    
    bool openSheet(SheetIndex sheetIndex) {
        if (sheetIndex < sheets.size()) {
            currentSheet = &sheets[sheetIndex];
            resetSheet();
            return true;
        }
        return false;
    }
    
    // 解析单元格引用，如"A1" -> (0, 0), "B2" -> (1, 1)
    static CellPosition parseCellReference(const std::string& cellRef) {
        if (cellRef.empty()) {
            return {0, 0};
        }
        
        size_t i = 0;
        ColumnIndex col = 0;
        
        // 解析列部分（字母）
        while (i < cellRef.length() && std::isalpha(cellRef[i])) {
            col = col * 26 + (std::toupper(cellRef[i]) - 'A' + 1);
            i++;
        }
        col--; // 转换为0基于的索引
        
        // 解析行部分（数字）
        RowIndex row = 0;
        if (i < cellRef.length()) {
            row = std::strtoul(cellRef.c_str() + i, nullptr, 10);
            if (row > 0) row--; // 转换为0基于的索引
        }
        
        return {row, col};
    }
    
    // 解析工作表数据
    TableData parseSheetData(const std::string& sheetPath) {
        std::string content = readZipEntry(sheetPath);
        if (content.empty()) {
            return {};
        }
        
        TableData result;
        std::map<RowIndex, RowData> rows; // 使用map确保行顺序
        
        // 解析工作表XML - 简化版本，只提取基本数据
        // 查找所有的行数据
        size_t pos = 0;
        while ((pos = content.find("<row", pos)) != std::string::npos) {
            size_t rowEnd = content.find("</row>", pos);
            if (rowEnd == std::string::npos) {
                break;
            }
            
            // 提取行号
            size_t rPos = content.find("r=\"", pos);
            RowIndex rowIndex = 0;
            if (rPos != std::string::npos && rPos < rowEnd) {
                rPos += 3; // skip r="
                size_t rEnd = content.find("\"", rPos);
                if (rEnd != std::string::npos) {
                    rowIndex = std::strtoul(content.substr(rPos, rEnd - rPos).c_str(), nullptr, 10) - 1;
                }
            }
            
            // 解析这一行的单元格
            RowData rowData;
            size_t cellPos = pos;
            while ((cellPos = content.find("<c ", cellPos)) != std::string::npos && cellPos < rowEnd) {
                size_t cellEnd = content.find("</c>", cellPos);
                if (cellEnd == std::string::npos || cellEnd > rowEnd) {
                    cellEnd = content.find("/>", cellPos);
                    if (cellEnd == std::string::npos || cellEnd > rowEnd) {
                        break;
                    }
                }
                
                // 提取单元格引用
                size_t refPos = content.find("r=\"", cellPos);
                std::string cellRef;
                if (refPos != std::string::npos && refPos < cellEnd) {
                    refPos += 3; // skip r="
                    size_t refEnd = content.find("\"", refPos);
                    if (refEnd != std::string::npos) {
                        cellRef = content.substr(refPos, refEnd - refPos);
                    }
                }
                
                // 提取单元格值
                size_t valuePos = content.find("<v>", cellPos);
                CellValue cellValue;
                if (valuePos != std::string::npos && valuePos < cellEnd) {
                    valuePos += 3; // skip <v>
                    size_t valueEnd = content.find("</v>", valuePos);
                    if (valueEnd != std::string::npos) {
                        std::string valueStr = content.substr(valuePos, valueEnd - valuePos);
                        
                        // 检查是否是共享字符串
                        size_t typePos = content.find("t=\"s\"", cellPos);
                        if (typePos != std::string::npos && typePos < cellEnd) {
                            // 共享字符串索引
                            try {
                                size_t index = std::stoul(valueStr);
                                if (index < sharedStrings.size()) {
                                    cellValue = sharedStrings[index];
                                } else {
                                    cellValue = valueStr;
                                }
                            } catch (...) {
                                cellValue = valueStr;
                            }
                        } else {
                            // 普通值
                            cellValue = Reader::stringToCellValue(valueStr);
                        }
                    }
                }
                
                // 解析单元格位置并设置值
                if (!cellRef.empty()) {
                    auto pos = parseCellReference(cellRef);
                    
                    // 确保行有足够的列
                    while (rowData.size() <= pos.column) {
                        rowData.push_back(std::monostate{});
                    }
                    
                    rowData[pos.column] = cellValue;
                }
                
                cellPos = cellEnd + 1;
            }
            
            rows[rowIndex] = rowData;
            pos = rowEnd + 1;
        }
        
        // 将map转换为vector，确保行顺序
        if (!rows.empty()) {
            RowIndex maxRow = rows.rbegin()->first;
            result.resize(maxRow + 1);
            
            for (const auto& [rowIndex, rowData] : rows) {
                result[rowIndex] = rowData;
            }
        }
        
        return result;
    }
    
    // 确保表格数据被缓存
    void ensureTableDataCached() {
        if (!tableDataCached && currentSheet) {
            cachedTableData = parseSheetData(currentSheet->filePath);
            tableDataCached = true;
            
            // 计算维度
            if (!cachedTableData.empty()) {
                RowIndex maxRow = cachedTableData.size();
                ColumnIndex maxCol = 0;
                
                for (const auto& row : cachedTableData) {
                    if (row.size() > maxCol) {
                        maxCol = row.size();
                    }
                }
                
                cachedDimensions = std::make_pair(maxRow, maxCol);
            } else {
                cachedDimensions = std::make_pair(0, 0);
            }
        }
    }
    
    // 读取特定行
    std::optional<RowData> getRow(RowIndex rowIndex, ColumnIndex maxColumns = 0) {
        ensureTableDataCached();
        
        if (rowIndex >= cachedTableData.size()) {
            return std::nullopt;
        }
        
        RowData result = cachedTableData[rowIndex];
        
        // 应用最大列限制
        if (maxColumns > 0 && result.size() > maxColumns) {
            result.resize(maxColumns);
        }
        
        return result;
    }
    
    // 读取特定单元格
    std::optional<CellValue> getCell(const CellPosition& position) {
        ensureTableDataCached();
        
        if (position.row >= cachedTableData.size()) {
            return std::nullopt;
        }
        
        const auto& row = cachedTableData[position.row];
        if (position.column >= row.size()) {
            return std::nullopt;
        }
        
        return row[position.column];
    }
    
    // 读取范围数据
    TableData getRange(const CellRange& range) {
        ensureTableDataCached();
        
        TableData result;
        
        RowIndex startRow = std::min(range.start.row, static_cast<RowIndex>(cachedTableData.size()));
        RowIndex endRow = std::min(range.end.row + 1, static_cast<RowIndex>(cachedTableData.size()));
        
        for (RowIndex r = startRow; r < endRow; ++r) {
            const auto& sourceRow = cachedTableData[r];
            RowData resultRow;
            
            ColumnIndex startCol = std::min(range.start.column, static_cast<ColumnIndex>(sourceRow.size()));
            ColumnIndex endCol = std::min(range.end.column + 1, static_cast<ColumnIndex>(sourceRow.size()));
            
            for (ColumnIndex c = startCol; c < endCol; ++c) {
                resultRow.push_back(sourceRow[c]);
            }
            
            result.push_back(resultRow);
        }
        
        return result;
    }
    
    // 重置工作表状态
    void resetSheet() {
        currentRowIndex = 0;
        atEnd = false;
        parseContext.reset();
        tableDataCached = false;
        cachedTableData.clear();
    }
};

Reader::Reader(const std::string& filePath) 
    : pImpl_(std::make_unique<Impl>()) {
    pImpl_->openFile(filePath);
}

Reader::Reader(Reader&& other) noexcept 
    : pImpl_(std::move(other.pImpl_)) {
}

Reader& Reader::operator=(Reader&& other) noexcept {
    if (this != &other) {
        pImpl_ = std::move(other.pImpl_);
    }
    return *this;
}

Reader::~Reader() = default;

std::vector<std::string> Reader::getSheetNames() const {
    std::vector<std::string> sheetNames;
    for (const auto& sheet : pImpl_->sheets) {
        sheetNames.push_back(sheet.name);
    }
    return sheetNames;
}

bool Reader::openSheet(const std::string& sheetName) {
    return pImpl_->openSheet(sheetName);
}

bool Reader::openSheet(SheetIndex sheetIndex) {
    return pImpl_->openSheet(sheetIndex);
}

std::optional<RowIndex> Reader::getRowCount() const {
    if (!pImpl_->currentSheet) {
        return std::nullopt;
    }
    
    // 确保维度已缓存
    const_cast<Reader::Impl*>(pImpl_.get())->ensureTableDataCached();
    
    if (pImpl_->cachedDimensions) {
        return pImpl_->cachedDimensions->first;
    }
    return std::nullopt;
}

std::optional<ColumnIndex> Reader::getColumnCount() const {
    if (!pImpl_->currentSheet) {
        return std::nullopt;
    }
    
    // 确保维度已缓存
    const_cast<Reader::Impl*>(pImpl_.get())->ensureTableDataCached();
    
    if (pImpl_->cachedDimensions) {
        return pImpl_->cachedDimensions->second;
    }
    return std::nullopt;
}

bool Reader::readNextRow(RowData& rowData, ColumnIndex maxColumns) {
    if (!pImpl_->currentSheet || pImpl_->atEnd) {
        return false;
    }
    
    auto row = pImpl_->getRow(pImpl_->currentRowIndex, maxColumns);
    if (row) {
        rowData = *row;
        pImpl_->currentRowIndex++;
        return true;
    } else {
        pImpl_->atEnd = true;
        return false;
    }
}

std::optional<RowData> Reader::readRow(RowIndex rowIndex, ColumnIndex maxColumns) {
    if (!pImpl_->currentSheet) {
        return std::nullopt;
    }
    
    return pImpl_->getRow(rowIndex, maxColumns);
}

std::optional<CellValue> Reader::readCell(const CellPosition& position) {
    if (!pImpl_->currentSheet) {
        return std::nullopt;
    }
    
    return pImpl_->getCell(position);
}

TableData Reader::readRange(const CellRange& range) {
    if (!range.isValid()) {
        throw InvalidArgumentException("无效的单元格范围");
    }
    
    if (!pImpl_->currentSheet) {
        return {};
    }
    
    return pImpl_->getRange(range);
}

TableData Reader::readAll(RowIndex maxRows, ColumnIndex maxColumns, bool skipEmptyRows) {
    if (!pImpl_->currentSheet) {
        return {};
    }
    
    // 使用新的解析器读取工作表数据
    TableData result = pImpl_->parseSheetData(pImpl_->currentSheet->filePath);
    
    // 应用限制和过滤
    if (maxRows > 0 && result.size() > maxRows) {
        result.resize(maxRows);
    }
    
    if (maxColumns > 0) {
        for (auto& row : result) {
            if (row.size() > maxColumns) {
                row.resize(maxColumns);
            }
        }
    }
    
    if (skipEmptyRows) {
        result.erase(
            std::remove_if(result.begin(), result.end(),
                [](const RowData& row) { return isEmptyRow(row); }),
            result.end()
        );
    }
    
    return result;
}

RowIndex Reader::readAllRows(const RowCallback& callback, ColumnIndex maxColumns, bool skipEmptyRows) {
    if (!callback) {
        return 0;
    }
    
    if (!pImpl_->currentSheet) {
        return 0;
    }
    
    RowIndex rowCount = 0;
    reset();
    
    // 确保数据已缓存
    pImpl_->ensureTableDataCached();
    
    for (RowIndex i = 0; i < pImpl_->cachedTableData.size(); ++i) {
        const auto& row = pImpl_->cachedTableData[i];
        
        // 应用最大列限制
        RowData processedRow = row;
        if (maxColumns > 0 && processedRow.size() > maxColumns) {
            processedRow.resize(maxColumns);
        }
        
        if (skipEmptyRows && isEmptyRow(processedRow)) {
            continue;
        }
        
        if (!callback(rowCount, processedRow)) {
            break;
        }
        
        ++rowCount;
    }
    
    return rowCount;
}

size_t Reader::readAllCells(const CellCallback& callback, const std::optional<CellRange>& range) {
    if (!callback) {
        return 0;
    }
    
    if (!pImpl_->currentSheet) {
        return 0;
    }
    
    size_t cellCount = 0;
    
    // 确定读取范围
    CellRange actualRange;
    if (range) {
        actualRange = *range;
    } else {
        // 读取整个工作表
        pImpl_->ensureTableDataCached();
        RowIndex maxRow = pImpl_->cachedTableData.size();
        ColumnIndex maxCol = 0;
        
        for (const auto& row : pImpl_->cachedTableData) {
            if (row.size() > maxCol) {
                maxCol = row.size();
            }
        }
        
        actualRange = CellRange{0, 0, maxRow > 0 ? maxRow - 1 : 0, maxCol > 0 ? maxCol - 1 : 0};
    }
    
    reset();
    pImpl_->ensureTableDataCached();
    
    RowIndex endRow = std::min(actualRange.end.row + 1, static_cast<RowIndex>(pImpl_->cachedTableData.size()));
    
    for (RowIndex rowIndex = actualRange.start.row; rowIndex < endRow; ++rowIndex) {
        if (rowIndex >= pImpl_->cachedTableData.size()) {
            break;
        }
        
        const auto& row = pImpl_->cachedTableData[rowIndex];
        ColumnIndex endCol = std::min(actualRange.end.column + 1, static_cast<ColumnIndex>(row.size()));
        
        for (ColumnIndex col = actualRange.start.column; col < endCol; ++col) {
            if (col >= row.size()) {
                break;
            }
            
            CellPosition pos(rowIndex, col);
            if (!callback(pos, row[col])) {
                return cellCount;
            }
            ++cellCount;
        }
    }
    
    return cellCount;
}

void Reader::reset() {
    pImpl_->resetSheet();
}

bool Reader::isAtEnd() const {
    if (!pImpl_->currentSheet) {
        return true;
    }
    
    // 如果还没有缓存数据，检查是否有数据
    if (!pImpl_->tableDataCached) {
        const_cast<Reader::Impl*>(pImpl_.get())->ensureTableDataCached();
    }
    
    return pImpl_->atEnd || pImpl_->currentRowIndex >= pImpl_->cachedTableData.size();
}

RowIndex Reader::getCurrentRowIndex() const {
    return pImpl_->currentRowIndex;
}

bool Reader::isEmptyRow(const RowData& rowData) {
    return std::all_of(rowData.begin(), rowData.end(), 
        [](const auto& cell) { return isEmptyCell(cell); });
}

bool Reader::isEmptyCell(const CellValue& value) {
    return std::visit([](const auto& v) -> bool {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, std::string>) {
            return v.empty();
        } else if constexpr (std::is_same_v<T, double>) {
            return v == 0.0;
        } else if constexpr (std::is_same_v<T, Integer>) {
            return v == 0;
        } else if constexpr (std::is_same_v<T, bool>) {
            return !v;
        } else {
            return true; // std::monostate
        }
    }, value);
}

CellValue Reader::stringToCellValue(const std::string& str) {
    if (str.empty()) {
        return std::monostate{};
    }
    
    // 尝试解析为数字
    if (std::all_of(str.begin(), str.end(), [](char c) { 
        return std::isdigit(c) || c == '.' || c == '-' || c == '+' || c == 'e' || c == 'E'; 
    })) {
        if (str.find('.') != std::string::npos || str.find('e') != std::string::npos || str.find('E') != std::string::npos) {
            // 浮点数
            try {
                return std::stod(str);
            } catch (...) {
                return str;
            }
        } else {
            // 整数
            try {
                return static_cast<Integer>(std::stoll(str));
            } catch (...) {
                return str;
            }
        }
    }
    
    // 布尔值
    if (str == "true" || str == "TRUE" || str == "1") {
        return true;
    } else if (str == "false" || str == "FALSE" || str == "0") {
        return false;
    }
    
    return str;
}

std::string Reader::cellValueToString(const CellValue& value) {
    return std::visit([](const auto& v) -> std::string {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, std::string>) {
            return v;
        } else if constexpr (std::is_same_v<T, double>) {
            return std::to_string(v);
        } else if constexpr (std::is_same_v<T, Integer>) {
            return std::to_string(v);
        } else if constexpr (std::is_same_v<T, bool>) {
            return v ? "true" : "false";
        } else {
            return "";
        }
    }, value);
}

} // namespace TinaXlsx
