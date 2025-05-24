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
    
    // 预分配容量以减少内存重分配
    if (maxColumns == 0) {
        maxColumns = 100; // 默认最大100列
    }
    
    rowData.clear();
    rowData.reserve(maxColumns); // 预分配容量
    
    ColumnIndex columnIndex = 0;
    XLSXIOCHAR* cellValue = nullptr;
    
    // 优化：批量读取单元格以减少函数调用开销
    while (xlsxioread_sheet_next_cell_string(pImpl_->sheet, &cellValue) && columnIndex < maxColumns) {
        if (cellValue) {
            // 避免不必要的字符串拷贝
            rowData.emplace_back(stringToCellValue(std::string(cellValue)));
            xlsxioread_free(cellValue);
            cellValue = nullptr;
        } else {
            // 空单元格
            rowData.emplace_back(std::monostate{});
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
    // 使用索引而不是std::visit获得更好的性能
    // 正确的类型索引：0=string, 1=double, 2=int64_t, 3=bool, 4=monostate
    const uint8_t type = value.index();
    
    switch (type) {
        case 0: // std::string
            return std::get<std::string>(value).empty();
        case 4: // std::monostate
            return true;
        default:
            return false;
    }
}

CellValue Reader::stringToCellValue(const std::string& str) {
    if (str.empty()) {
        return std::monostate{};
    }
    
    const char* data = str.data();
    const size_t len = str.length();
    
    // 快速检查第一个字符来优化常见情况
    const char first = data[0];
    
    // 优化数字检测：先快速检查字符集
    bool couldBeNumber = false;
    bool couldBeInt = true;
    bool hasDecimal = false;
    bool hasExponent = false;
    
    // 更高效的数字检查
    for (size_t i = 0; i < len; ++i) {
        const char c = data[i];
        if (c >= '0' && c <= '9') {
            couldBeNumber = true;
        } else if (c == '.') {
            if (hasDecimal) break; // 第二个小数点，不是数字
            hasDecimal = true;
            couldBeInt = false;
            couldBeNumber = true;
        } else if (c == 'e' || c == 'E') {
            if (hasExponent) break; // 第二个指数，不是数字
            hasExponent = true;
            couldBeInt = false;
            couldBeNumber = true;
        } else if (c == '-' || c == '+') {
            if (i != 0 && data[i-1] != 'e' && data[i-1] != 'E') {
                couldBeNumber = false;
                break;
            }
            couldBeNumber = true;
        } else {
            couldBeNumber = false;
            break;
        }
    }
    
    if (couldBeNumber) {
        if (couldBeInt && !hasDecimal && !hasExponent) {
            // 尝试解析为整数
            int64_t intValue;
            auto result = std::from_chars(data, data + len, intValue);
            if (result.ec == std::errc{} && result.ptr == data + len) {
                return intValue;
            }
        }
        
        // 尝试解析为浮点数
        double doubleValue;
        auto result = std::from_chars(data, data + len, doubleValue);
        if (result.ec == std::errc{} && result.ptr == data + len) {
            return doubleValue;
        }
    }
    
    // 优化布尔值检测：直接检查而不创建子字符串
    if (len <= 5) { // 最长的布尔值字符串是"false"
        // 使用预编译的比较函数
        if ((len == 4 && 
             (first == 't' || first == 'T') &&
             (data[1] == 'r' || data[1] == 'R') &&
             (data[2] == 'u' || data[2] == 'U') &&
             (data[3] == 'e' || data[3] == 'E')) ||
            (len == 1 && first == '1')) {
            return true;
        }
        
        if ((len == 5 && 
             (first == 'f' || first == 'F') &&
             (data[1] == 'a' || data[1] == 'A') &&
             (data[2] == 'l' || data[2] == 'L') &&
             (data[3] == 's' || data[3] == 'S') &&
             (data[4] == 'e' || data[4] == 'E')) ||
            (len == 1 && first == '0')) {
            return false;
        }
    }
    
    // 默认作为字符串，避免额外的拷贝
    return str;
}

std::string Reader::cellValueToString(const CellValue& value) {
    // 使用索引而不是std::visit获得更好的性能
    // 正确的类型索引：0=string, 1=double, 2=int64_t, 3=bool, 4=monostate
    const uint8_t type = value.index();
    
    switch (type) {
        case 0: // std::string
            return std::get<std::string>(value);
        case 1: { // double
            const double val = std::get<double>(value);
            // 优化：检查是否为整数以避免不必要的小数点
            if (val == static_cast<int64_t>(val) && 
                val >= -9007199254740992.0 && val <= 9007199254740992.0) {
                return std::to_string(static_cast<int64_t>(val));
            }
            return std::to_string(val);
        }
        case 2: // int64_t
            return std::to_string(std::get<int64_t>(value));
        case 3: // bool
            return std::get<bool>(value) ? "true" : "false";
        case 4: // std::monostate
            return "";
        default:
            return "";
    }
}

} // namespace TinaXlsx
