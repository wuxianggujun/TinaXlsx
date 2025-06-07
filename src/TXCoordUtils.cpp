//
// @file TXCoordUtils.cpp
// @brief 统一的坐标转换工具实现
//

#include "TinaXlsx/TXCoordUtils.hpp"
#include "TinaXlsx/TXHighPerformanceLogger.hpp"
#include <algorithm>
#include <cctype>
#include <charconv>

namespace TinaXlsx {

// ==================== 单个坐标转换 ====================

TXResult<TXCoordinate> TXCoordUtils::parseCoord(std::string_view excel_coord) {
    if (excel_coord.empty()) {
        return TXResult<TXCoordinate>(TXError(TXErrorCode::InvalidArgument, "坐标字符串为空"));
    }
    
    auto [row, col] = parseCoordFast(excel_coord);
    if (row == INVALID_INDEX || col == INVALID_INDEX) {
        return TXResult<TXCoordinate>(TXError(TXErrorCode::InvalidArgument, 
            "无效的Excel坐标格式: " + std::string(excel_coord)));
    }
    
    // parseCoordFast返回0-based，需要转换为1-based给row_t和column_t
    return TXResult<TXCoordinate>(TXCoordinate(row_t(row + 1), column_t(col + 1)));
}

std::pair<uint32_t, uint32_t> TXCoordUtils::parseCoordFast(std::string_view excel_coord) noexcept {
    if (excel_coord.empty()) {
        return {INVALID_INDEX, INVALID_INDEX};
    }
    
    // 分离字母和数字部分
    size_t letter_end = 0;
    while (letter_end < excel_coord.size() && std::isalpha(excel_coord[letter_end])) {
        ++letter_end;
    }
    
    if (letter_end == 0 || letter_end == excel_coord.size()) {
        return {INVALID_INDEX, INVALID_INDEX};
    }
    
    // 解析列字母
    uint32_t col = parseColumnLettersInternal(excel_coord.substr(0, letter_end));
    if (col == INVALID_INDEX) {
        return {INVALID_INDEX, INVALID_INDEX};
    }
    
    // 解析行号
    uint32_t row = parseRowNumberInternal(excel_coord.substr(letter_end));
    if (row == INVALID_INDEX) {
        return {INVALID_INDEX, INVALID_INDEX};
    }
    
    return {row - 1, col - 1}; // 转换为0-based
}

std::string TXCoordUtils::coordToExcel(const TXCoordinate& coord) {
    // TXCoordinate内部存储1-based，需要转换为0-based给coordToExcel
    return coordToExcel(coord.getRow().index() - 1, coord.getCol().index() - 1);
}

std::string TXCoordUtils::coordToExcel(uint32_t row, uint32_t col) {
    std::string result;
    columnIndexToLettersInternal(col + 1, result); // 转换为1-based
    result += std::to_string(row + 1); // 转换为1-based
    return result;
}

// ==================== 范围转换 ====================

TXResult<std::pair<TXCoordinate, TXCoordinate>> TXCoordUtils::parseRange(std::string_view excel_range) {
    // 查找冒号分隔符
    size_t colon_pos = excel_range.find(':');
    if (colon_pos == std::string_view::npos) {
        return TXResult<std::pair<TXCoordinate, TXCoordinate>>(
            TXError(TXErrorCode::InvalidArgument, "范围格式无效，缺少冒号: " + std::string(excel_range)));
    }
    
    std::string_view start_coord = excel_range.substr(0, colon_pos);
    std::string_view end_coord = excel_range.substr(colon_pos + 1);
    
    auto start_result = parseCoord(start_coord);
    if (start_result.isError()) {
        return TXResult<std::pair<TXCoordinate, TXCoordinate>>(start_result.error());
    }
    
    auto end_result = parseCoord(end_coord);
    if (end_result.isError()) {
        return TXResult<std::pair<TXCoordinate, TXCoordinate>>(end_result.error());
    }
    
    return TXResult<std::pair<TXCoordinate, TXCoordinate>>(
        std::make_pair(start_result.value(), end_result.value()));
}

std::string TXCoordUtils::rangeToExcel(const TXCoordinate& start, const TXCoordinate& end) {
    return coordToExcel(start) + ":" + coordToExcel(end);
}

// ==================== 批量转换 ====================

void TXCoordUtils::coordsBatchToExcel(const TXCoordinate* coords, size_t count, 
                                      std::vector<std::string>& output) {
    output.clear();
    output.reserve(count);
    
    for (size_t i = 0; i < count; ++i) {
        output.push_back(coordToExcel(coords[i]));
    }
}

void TXCoordUtils::packedCoordsBatchToExcel(const uint32_t* packed_coords, size_t count,
                                            std::vector<std::string>& output) {
    output.clear();
    output.reserve(count);
    
    for (size_t i = 0; i < count; ++i) {
        uint32_t row = packed_coords[i] >> 16;
        uint32_t col = packed_coords[i] & 0xFFFF;
        output.push_back(coordToExcel(row, col));
    }
}

// ==================== 列转换工具 ====================

uint32_t TXCoordUtils::columnLettersToIndex(std::string_view col_letters) noexcept {
    return parseColumnLettersInternal(col_letters);
}

std::string TXCoordUtils::columnIndexToLetters(uint32_t col_index) {
    std::string result;
    columnIndexToLettersInternal(col_index + 1, result); // 转换为1-based
    return result;
}

// ==================== 验证工具 ====================

bool TXCoordUtils::isValidExcelCoord(std::string_view excel_coord) noexcept {
    auto [row, col] = parseCoordFast(excel_coord);
    return row != INVALID_INDEX && col != INVALID_INDEX && 
           row < MAX_EXCEL_ROWS && col < MAX_EXCEL_COLS;
}

bool TXCoordUtils::isValidExcelRange(std::string_view excel_range) noexcept {
    size_t colon_pos = excel_range.find(':');
    if (colon_pos == std::string_view::npos) {
        return false;
    }
    
    std::string_view start_coord = excel_range.substr(0, colon_pos);
    std::string_view end_coord = excel_range.substr(colon_pos + 1);
    
    return isValidExcelCoord(start_coord) && isValidExcelCoord(end_coord);
}

// ==================== 内部优化方法 ====================

uint32_t TXCoordUtils::parseColumnLettersInternal(std::string_view letters) noexcept {
    if (letters.empty()) {
        return INVALID_INDEX;
    }
    
    uint32_t result = 0;
    for (char c : letters) {
        if (!std::isalpha(c)) {
            return INVALID_INDEX;
        }
        result = result * 26 + (std::toupper(c) - 'A' + 1);
    }
    
    return result; // 1-based
}

uint32_t TXCoordUtils::parseRowNumberInternal(std::string_view numbers) noexcept {
    if (numbers.empty()) {
        return INVALID_INDEX;
    }
    
    uint32_t result = 0;
    for (char c : numbers) {
        if (!std::isdigit(c)) {
            return INVALID_INDEX;
        }
        result = result * 10 + (c - '0');
    }
    
    return result; // 1-based
}

void TXCoordUtils::columnIndexToLettersInternal(uint32_t col_index, std::string& result) {
    result.clear();
    uint32_t temp = col_index; // 1-based输入
    
    while (temp > 0) {
        temp--;
        result = char('A' + (temp % 26)) + result;
        temp /= 26;
    }
}

} // namespace TinaXlsx
