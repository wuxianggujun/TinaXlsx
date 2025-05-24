/**
 * @file Reader.cpp
 * @brief 重构后的高性能Excel读取器实现
 * 
 * 新架构说明：
 * - Reader: 只负责提供读取接口和管理当前状态
 * - ExcelStructureManager: 负责Excel文件结构的解析和管理
 * - WorksheetDataParser: 负责工作表数据的解析
 * - DataCache: 负责数据的缓存管理
 * - ZipReader/ExcelZipReader: 负责ZIP文件操作
 * - XmlParser组件: 负责XML解析
 */

#include "TinaXlsx/Reader.hpp"
#include "TinaXlsx/ZipReader.hpp"
#include <algorithm>
#include <charconv>
#include <sstream>

namespace TinaXlsx {

// =============================================================================
// Reader 构造和销毁
// =============================================================================

Reader::Reader(const std::string& filePath) {
    try {
        // 创建Excel ZIP读取器
        auto zipReader = std::make_unique<ExcelZipReader>(filePath);
        
        // 创建结构管理器
        structureManager_ = std::make_unique<ExcelStructureManager>(std::move(zipReader));
        
        // 创建数据缓存
        cache_ = std::make_unique<DataCache>();
        
        // 创建工作表数据解析器
        parser_ = std::make_unique<WorksheetDataParser>(structureManager_.get(), cache_.get());
        
    } catch (const std::exception& e) {
        throw FileException("Failed to open Excel file '" + filePath + "': " + e.what());
    }
}

Reader::Reader(Reader&& other) noexcept 
    : structureManager_(std::move(other.structureManager_)),
      cache_(std::move(other.cache_)),
      parser_(std::move(other.parser_)),
      currentSheet_(std::move(other.currentSheet_)),
      currentRowIndex_(other.currentRowIndex_) {
}

Reader& Reader::operator=(Reader&& other) noexcept {
    if (this != &other) {
        structureManager_ = std::move(other.structureManager_);
        cache_ = std::move(other.cache_);
        parser_ = std::move(other.parser_);
        currentSheet_ = std::move(other.currentSheet_);
        currentRowIndex_ = other.currentRowIndex_;
    }
    return *this;
}

Reader::~Reader() = default;

// =============================================================================
// 工作表管理
// =============================================================================

std::vector<std::string> Reader::getSheetNames() const {
    if (!structureManager_) {
        return {};
    }
    
    return structureManager_->getSheetNames();
}

bool Reader::openSheet(const std::string& sheetName) {
    if (!structureManager_) {
        return false;
    }
    
    auto sheetInfo = structureManager_->findSheetByName(sheetName);
    if (!sheetInfo) {
        return false;
    }
    
    currentSheet_ = *sheetInfo;
    currentRowIndex_ = 0;
    
    return true;
}

bool Reader::openSheet(SheetIndex sheetIndex) {
    if (!structureManager_) {
        return false;
    }
    
    auto sheetInfo = structureManager_->getSheetByIndex(sheetIndex);
    if (!sheetInfo) {
        return false;
    }
    
    currentSheet_ = *sheetInfo;
    currentRowIndex_ = 0;
    
    return true;
}

// =============================================================================
// 维度信息
// =============================================================================

std::optional<RowIndex> Reader::getRowCount() const {
    if (!currentSheet_ || !parser_) {
        return std::nullopt;
    }
    
    try {
        auto dimensions = parser_->getSheetDimensions(*currentSheet_);
        return dimensions.first;
    } catch (const std::exception&) {
        return std::nullopt;
    }
}

std::optional<ColumnIndex> Reader::getColumnCount() const {
    if (!currentSheet_ || !parser_) {
        return std::nullopt;
    }
    
    try {
        auto dimensions = parser_->getSheetDimensions(*currentSheet_);
        return dimensions.second;
    } catch (const std::exception&) {
        return std::nullopt;
    }
}

// =============================================================================
// 数据读取
// =============================================================================

bool Reader::readNextRow(RowData& rowData, ColumnIndex maxColumns) {
    ensureSheetOpen();
    
    try {
        auto row = parser_->parseRow(*currentSheet_, currentRowIndex_, maxColumns);
        if (!row) {
            return false;
        }
        
        rowData = *row;
        currentRowIndex_++;
        return true;
        
    } catch (const std::exception&) {
        return false;
    }
}

std::optional<RowData> Reader::readRow(RowIndex rowIndex, ColumnIndex maxColumns) {
    ensureSheetOpen();
    
    try {
        return parser_->parseRow(*currentSheet_, rowIndex, maxColumns);
    } catch (const std::exception&) {
        return std::nullopt;
    }
}

std::optional<CellValue> Reader::readCell(const CellPosition& position) {
    ensureSheetOpen();
    
    try {
        return parser_->parseCell(*currentSheet_, position);
    } catch (const std::exception&) {
        return std::nullopt;
    }
}

TableData Reader::readRange(const CellRange& range) {
    ensureSheetOpen();
    
    try {
        return parser_->parseRange(*currentSheet_, range);
    } catch (const std::exception&) {
        return {};
    }
}

TableData Reader::readAll(RowIndex maxRows, ColumnIndex maxColumns, bool skipEmptyRows) {
    ensureSheetOpen();
    
    try {
        auto fullData = parser_->parseWorksheetData(*currentSheet_);
        
        // 应用限制
        if (maxRows > 0 && fullData.size() > maxRows) {
            fullData.resize(maxRows);
        }
        
        for (auto& row : fullData) {
            if (maxColumns > 0 && row.size() > maxColumns) {
                row.resize(maxColumns);
            }
        }
        
        // 跳过空行
        if (skipEmptyRows) {
            fullData.erase(
                std::remove_if(fullData.begin(), fullData.end(),
                    [](const RowData& row) { return isEmptyRow(row); }),
                fullData.end()
            );
        }
        
        return fullData;
        
    } catch (const std::exception&) {
        return {};
    }
}

// =============================================================================
// 流式读取
// =============================================================================

RowIndex Reader::readAllRows(const RowCallback& callback, ColumnIndex maxColumns, bool skipEmptyRows) {
    ensureSheetOpen();
    
    if (!callback) {
        return 0;
    }
    
    try {
        RowIndex processedRows = 0;
        
        auto rowsProcessed = parser_->parseWorksheetDataWithCallback(
            *currentSheet_,
            nullptr, // 不需要单元格回调
            [&](RowIndex rowIndex, const RowData& rowData) -> bool {
                // 应用列数限制
                RowData processedRow = rowData;
                if (maxColumns > 0 && processedRow.size() > maxColumns) {
                    processedRow.resize(maxColumns);
                }
                
                // 跳过空行检查
                if (skipEmptyRows && isEmptyRow(processedRow)) {
                    return true; // 继续处理
                }
                
                processedRows++;
                return callback(rowIndex, processedRow);
            }
        );
        
        return processedRows;
        
    } catch (const std::exception&) {
        return 0;
    }
}

size_t Reader::readAllCells(const CellCallback& callback, const std::optional<CellRange>& range) {
    ensureSheetOpen();
    
    if (!callback) {
        return 0;
    }
    
    try {
        size_t processedCells = 0;
        
        if (range) {
            // 读取指定范围
            auto rangeData = parser_->parseRange(*currentSheet_, *range);
            
            for (RowIndex row = 0; row < rangeData.size(); ++row) {
                for (ColumnIndex col = 0; col < rangeData[row].size(); ++col) {
                    CellPosition position{
                        range->start.row + row,
                        range->start.column + col
                    };
                    
                    if (!callback(position, rangeData[row][col])) {
                        return processedCells;
                    }
                    processedCells++;
                }
            }
        } else {
            // 读取整个工作表
            auto fullData = parser_->parseWorksheetData(*currentSheet_);
            
            for (RowIndex row = 0; row < fullData.size(); ++row) {
                for (ColumnIndex col = 0; col < fullData[row].size(); ++col) {
                    CellPosition position{row, col};
                    
                    if (!callback(position, fullData[row][col])) {
                        return processedCells;
                    }
                    processedCells++;
                }
            }
        }
        
        return processedCells;
        
    } catch (const std::exception&) {
        return 0;
    }
}

// =============================================================================
// 状态管理
// =============================================================================

void Reader::reset() {
    currentRowIndex_ = 0;
}

bool Reader::isAtEnd() const {
    if (!currentSheet_ || !parser_) {
        return true;
    }
    
    try {
        auto dimensions = parser_->getSheetDimensions(*currentSheet_);
        return currentRowIndex_ >= dimensions.first;
    } catch (const std::exception&) {
        return true;
    }
}

RowIndex Reader::getCurrentRowIndex() const {
    return currentRowIndex_;
}

// =============================================================================
// 静态工具函数
// =============================================================================

bool Reader::isEmptyRow(const RowData& rowData) {
    return std::all_of(rowData.begin(), rowData.end(),
        [](const CellValue& value) { return isEmptyCell(value); });
}

bool Reader::isEmptyCell(const CellValue& value) {
    return std::visit([](const auto& v) -> bool {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            return true;
        } else if constexpr (std::is_same_v<T, std::string>) {
            return v.empty();
        } else if constexpr (std::is_arithmetic_v<T>) {
            return v == 0;
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
    if (std::all_of(str.begin(), str.end(), 
        [](char c) { return std::isdigit(c) || c == '.' || c == '-' || c == '+' || c == 'e' || c == 'E'; })) {
        
        // 尝试解析为整数
        long long intVal;
        auto [ptr, ec] = std::from_chars(str.data(), str.data() + str.size(), intVal);
        if (ec == std::errc{} && ptr == str.data() + str.size()) {
            return intVal;
        }
        
        // 尝试解析为浮点数
        try {
            double doubleVal = std::stod(str);
            return doubleVal;
        } catch (...) {
            // 如果解析失败，返回字符串
        }
    }
    
    return str;
}

std::string Reader::cellValueToString(const CellValue& value) {
    return std::visit([](const auto& v) -> std::string {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            return "";
        } else if constexpr (std::is_same_v<T, std::string>) {
            return v;
        } else if constexpr (std::is_arithmetic_v<T>) {
            return std::to_string(v);
        } else {
            return "";
        }
    }, value);
}

// =============================================================================
// 私有方法
// =============================================================================

void Reader::ensureSheetOpen() const {
    if (!currentSheet_) {
        throw InvalidOperationException("No worksheet is currently open. Please call openSheet() first.");
    }
}

TableData Reader::getCurrentSheetData() {
    ensureSheetOpen();
    
    if (!parser_) {
        throw InvalidOperationException("Parser not available");
    }
    
    return parser_->parseWorksheetData(*currentSheet_);
}

} // namespace TinaXlsx
