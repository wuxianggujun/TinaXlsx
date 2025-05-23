/**
 * @file Reader.cpp
 * @brief 高性能Excel读取器实现
 */

#include "TinaXlsx/Reader.hpp"
#include <xlsxio_read.h>
#include <memory>
#include <vector>
#include <string>
#include <algorithm>
#include <charconv>
#include <filesystem>
#include <locale>
#include <codecvt>

namespace TinaXlsx {

struct Reader::Impl {
    xlsxioreader book = nullptr;
    xlsxioreadersheet sheet = nullptr;
    std::string filePath;
    std::string tempFilePath; // 用于处理非ASCII路径
    std::vector<std::string> sheetNames;
    RowIndex currentRowIndex = 0;
    bool atEnd = false;
    
    ~Impl() {
        close();
    }
    
    void close() {
        if (sheet) {
            xlsxioread_sheet_close(sheet);
            sheet = nullptr;
        }
        
        if (book) {
            xlsxioread_close(book);
            book = nullptr;
        }
        
        // 清理临时文件
        if (!tempFilePath.empty() && std::filesystem::exists(tempFilePath)) {
            try {
                std::filesystem::remove(tempFilePath);
            } catch (...) {
                // 忽略清理错误
            }
            tempFilePath.clear();
        }
        
        atEnd = false;
        currentRowIndex = 0;
    }
    
    bool openFile(const std::string& path) {
        close();
        filePath = path;
        
        // 检查文件是否存在
        std::filesystem::path fsPath;
        try {
#ifdef _WIN32
            // 在Windows上，将UTF-8字符串转换为宽字符路径
            std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
            std::wstring widePath = converter.from_bytes(path);
            fsPath = std::filesystem::path(widePath);
#else
            // 在其他平台上，使用UTF-8路径
            fsPath = std::filesystem::u8path(path);
#endif
        } catch (const std::exception& e) {
            throw FileException("Path encoding error: " + path + " (" + e.what() + ")");
        }
        
        if (!std::filesystem::exists(fsPath)) {
            throw FileException("File not found: " + path);
        }
        
        // 在Windows上尝试使用宽字符路径
#ifdef _WIN32
        try {
            std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
            std::wstring widePath = converter.from_bytes(path);
            std::string ansiPath = std::filesystem::path(widePath).string();
            book = xlsxioread_open(ansiPath.c_str());
        } catch (const std::exception&) {
            // 如果转换失败，使用临时文件方案
            book = nullptr;
        }
#else
        // 非Windows平台直接尝试打开
        book = xlsxioread_open(path.c_str());
#endif
        
        // 如果失败，可能是因为路径包含非ASCII字符，创建临时文件
        if (!book) {
            try {
                // 创建临时文件
                auto tempDir = std::filesystem::temp_directory_path();
                auto tempFileName = "TinaXlsx_" + std::to_string(std::hash<std::string>{}(path)) + ".xlsx";
                auto tempPath = tempDir / tempFileName;
                tempFilePath = tempPath.string();
                
                // 复制文件到临时位置
                std::filesystem::copy_file(fsPath, tempPath, std::filesystem::copy_options::overwrite_existing);
                
                // 尝试打开临时文件
                book = xlsxioread_open(tempFilePath.c_str());
            } catch (const std::exception& e) {
                throw FileException("Cannot open file: " + path + " (" + e.what() + ")");
            }
        }
        
        if (!book) {
            throw FileException("Failed to open Excel file: " + path);
        }
        
        // 获取工作表列表
        loadSheetNames();
        
        return true;
    }
    
    void loadSheetNames() {
        sheetNames.clear();
        
        if (!book) {
            return;
        }
        
        // 使用xlsxio_read库的sheet列表功能获取所有工作表名称（参考ExcelUtils.cpp的实现）
        xlsxioreadersheetlist sheetList = xlsxioread_sheetlist_open(book);
        if (!sheetList) {
            return;
        }
        
        // 逐个获取工作表名称
        const char* sheetName;
        while ((sheetName = xlsxioread_sheetlist_next(sheetList)) != NULL) {
            sheetNames.push_back(sheetName);
        }
        
        // 关闭sheet列表
        xlsxioread_sheetlist_close(sheetList);
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
    return pImpl_->sheetNames;
}

bool Reader::openSheet(const std::string& sheetName) {
    if (!pImpl_->book) {
        return false;
    }
    
    // 关闭当前工作表
    if (pImpl_->sheet) {
        xlsxioread_sheet_close(pImpl_->sheet);
        pImpl_->sheet = nullptr;
    }
    
    // 打开指定工作表，使用与原始ExcelParser相同的标志
    const char* name = sheetName.empty() ? nullptr : sheetName.c_str();
    pImpl_->sheet = xlsxioread_sheet_open(pImpl_->book, name, XLSXIOREAD_SKIP_EMPTY_ROWS);
    
    if (pImpl_->sheet) {
        pImpl_->currentRowIndex = 0;
        pImpl_->atEnd = false;
        return true;
    } else {
        // 如果指定名称失败，尝试默认工作表
        if (!sheetName.empty()) {
            pImpl_->sheet = xlsxioread_sheet_open(pImpl_->book, nullptr, XLSXIOREAD_SKIP_EMPTY_ROWS);
            if (pImpl_->sheet) {
                pImpl_->currentRowIndex = 0;
                pImpl_->atEnd = false;
                return true;
            }
        }
    }
    
    return false;
}

bool Reader::openSheet(SheetIndex sheetIndex) {
    if (sheetIndex >= pImpl_->sheetNames.size()) {
        return false;
    }
    
    return openSheet(pImpl_->sheetNames[sheetIndex]);
}

std::optional<RowIndex> Reader::getRowCount() const {
    // xlsxio不提供直接获取行数的方法
    // 返回空值表示未知
    return std::nullopt;
}

std::optional<ColumnIndex> Reader::getColumnCount() const {
    // xlsxio不提供直接获取列数的方法
    // 返回空值表示未知
    return std::nullopt;
}

bool Reader::readNextRow(RowData& rowData, ColumnIndex maxColumns) {
    if (!pImpl_->sheet || pImpl_->atEnd) {
        return false;
    }
    
    // 移动到下一行
    if (!xlsxioread_sheet_next_row(pImpl_->sheet)) {
        pImpl_->atEnd = true;
        return false;
    }
    
    rowData.clear();
    
    // 如果没有指定最大列数，使用一个合理的默认值
    if (maxColumns == 0) {
        maxColumns = 100; // 默认最大100列
    }
    
    ColumnIndex columnIndex = 0;
    XLSXIOCHAR* cellValue = nullptr;
    
    // 读取当前行的所有单元格
    while (xlsxioread_sheet_next_cell_string(pImpl_->sheet, &cellValue) && columnIndex < maxColumns) {
        std::string value = cellValue ? cellValue : "";
        rowData.push_back(stringToCellValue(value));
        
        if (cellValue) {
            xlsxioread_free(cellValue);
            cellValue = nullptr;
        }
        
        ++columnIndex;
    }
    
    ++pImpl_->currentRowIndex;
    return true;
}

std::optional<RowData> Reader::readRow(RowIndex rowIndex, ColumnIndex maxColumns) {
    // xlsxio是流式读取，不支持随机访问
    // 这里需要重置到开头然后顺序读取到指定行
    reset();
    
    RowData rowData;
    for (RowIndex i = 0; i <= rowIndex; ++i) {
        if (!readNextRow(rowData, maxColumns)) {
            return std::nullopt;
        }
    }
    
    return rowData;
}

std::optional<CellValue> Reader::readCell(const CellPosition& position) {
    auto rowData = readRow(position.row);
    if (!rowData || position.column >= rowData->size()) {
        return std::nullopt;
    }
    
    return (*rowData)[position.column];
}

TableData Reader::readRange(const CellRange& range) {
    if (!range.isValid()) {
        throw InvalidArgumentException("无效的单元格范围");
    }
    
    TableData result;
    result.reserve(range.rowCount());
    
    // 重置到开头
    reset();
    
    // 跳过到起始行
    RowData tempRow;
    for (RowIndex i = 0; i < range.start.row; ++i) {
        if (!readNextRow(tempRow)) {
            break;
        }
    }
    
    // 读取范围内的数据
    for (RowIndex row = range.start.row; row <= range.end.row; ++row) {
        RowData rowData;
        if (!readNextRow(rowData, range.end.column + 1)) {
            break;
        }
        
        // 提取指定列范围的数据
        RowData rangeRow;
        for (ColumnIndex col = range.start.column; col <= range.end.column; ++col) {
            if (col < rowData.size()) {
                rangeRow.push_back(rowData[col]);
            } else {
                rangeRow.push_back(std::monostate{}); // 空单元格
            }
        }
        
        result.push_back(std::move(rangeRow));
    }
    
    return result;
}

TableData Reader::readAll(RowIndex maxRows, ColumnIndex maxColumns, bool skipEmptyRows) {
    TableData result;
    
    reset();
    
    RowData rowData;
    RowIndex rowCount = 0;
    
    while (readNextRow(rowData, maxColumns)) {
        if (skipEmptyRows && isEmptyRow(rowData)) {
            continue;
        }
        
        result.push_back(rowData);
        ++rowCount;
        
        if (maxRows > 0 && rowCount >= maxRows) {
            break;
        }
    }
    
    return result;
}

RowIndex Reader::readAllRows(const RowCallback& callback, ColumnIndex maxColumns, bool skipEmptyRows) {
    if (!callback) {
        throw InvalidArgumentException("回调函数不能为空");
    }
    
    reset();
    
    RowData rowData;
    RowIndex rowCount = 0;
    RowIndex currentRow = 0;
    
    while (readNextRow(rowData, maxColumns)) {
        if (skipEmptyRows && isEmptyRow(rowData)) {
            ++currentRow;
            continue;
        }
        
        if (!callback(currentRow, rowData)) {
            break; // 回调函数返回false，停止读取
        }
        
        ++rowCount;
        ++currentRow;
    }
    
    return rowCount;
}

size_t Reader::readAllCells(const CellCallback& callback, const std::optional<CellRange>& range) {
    if (!callback) {
        throw InvalidArgumentException("回调函数不能为空");
    }
    
    size_t cellCount = 0;
    
    if (range) {
        // 读取指定范围
        auto tableData = readRange(*range);
        for (RowIndex row = 0; row < tableData.size(); ++row) {
            for (ColumnIndex col = 0; col < tableData[row].size(); ++col) {
                CellPosition pos(range->start.row + row, range->start.column + col);
                if (!callback(pos, tableData[row][col])) {
                    return cellCount;
                }
                ++cellCount;
            }
        }
    } else {
        // 读取整个工作表
        reset();
        RowData rowData;
        RowIndex rowIndex = 0;
        
        while (readNextRow(rowData)) {
            for (ColumnIndex col = 0; col < rowData.size(); ++col) {
                CellPosition pos(rowIndex, col);
                if (!callback(pos, rowData[col])) {
                    return cellCount;
                }
                ++cellCount;
            }
            ++rowIndex;
        }
    }
    
    return cellCount;
}

void Reader::reset() {
    if (!pImpl_->sheet) {
        return;
    }
    
    // xlsxio不支持重置，需要重新打开工作表
    std::string currentSheetName;
    if (!pImpl_->sheetNames.empty()) {
        currentSheetName = pImpl_->sheetNames[0]; // 假设我们知道当前工作表名
    }
    
    xlsxioread_sheet_close(pImpl_->sheet);
    pImpl_->sheet = xlsxioread_sheet_open(pImpl_->book, 
                                         currentSheetName.empty() ? nullptr : currentSheetName.c_str(), 
                                         XLSXIOREAD_SKIP_EMPTY_ROWS);
    
    pImpl_->currentRowIndex = 0;
    pImpl_->atEnd = false;
}

bool Reader::isAtEnd() const {
    return pImpl_->atEnd;
}

RowIndex Reader::getCurrentRowIndex() const {
    return pImpl_->currentRowIndex;
}

bool Reader::isEmptyRow(const RowData& rowData) {
    return std::all_of(rowData.begin(), rowData.end(), [](const CellValue& value) {
        return isEmptyCell(value);
    });
}

bool Reader::isEmptyCell(const CellValue& value) {
    return std::visit([](const auto& v) -> bool {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            return true;
        } else if constexpr (std::is_same_v<T, std::string>) {
            return v.empty();
        } else {
            return false;
        }
    }, value);
}

CellValue Reader::stringToCellValue(const std::string& str) {
    if (str.empty()) {
        return std::monostate{};
    }
    
    // 尝试解析为数字
    if (str.find_first_not_of("0123456789.-+eE") == std::string::npos) {
        // 尝试解析为整数
        if (str.find('.') == std::string::npos && str.find('e') == std::string::npos && str.find('E') == std::string::npos) {
            int64_t intValue;
            auto result = std::from_chars(str.data(), str.data() + str.size(), intValue);
            if (result.ec == std::errc{}) {
                return intValue;
            }
        }
        
        // 尝试解析为浮点数
        double doubleValue;
        auto result = std::from_chars(str.data(), str.data() + str.size(), doubleValue);
        if (result.ec == std::errc{}) {
            return doubleValue;
        }
    }
    
    // 尝试解析为布尔值
    if (str == "true" || str == "TRUE" || str == "True" || str == "1") {
        return true;
    } else if (str == "false" || str == "FALSE" || str == "False" || str == "0") {
        return false;
    }
    
    // 默认作为字符串
    return str;
}

std::string Reader::cellValueToString(const CellValue& value) {
    return std::visit([](const auto& v) -> std::string {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            return "";
        } else if constexpr (std::is_same_v<T, std::string>) {
            return v;
        } else if constexpr (std::is_same_v<T, double>) {
            return std::to_string(v);
        } else if constexpr (std::is_same_v<T, int64_t>) {
            return std::to_string(v);
        } else if constexpr (std::is_same_v<T, bool>) {
            return v ? "true" : "false";
        } else {
            return "";
        }
    }, value);
}

} // namespace TinaXlsx
