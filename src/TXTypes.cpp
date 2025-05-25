#include "TinaXlsx/TXTypes.hpp"
#include <algorithm>
#include <cctype>
#include <sstream>

namespace TinaXlsx {

std::string TXTypes::colIndexToName(ColIndex col) {
    if (col == 0 || col > MAX_COLS) {
        return "";
    }
    
    std::string result;
    while (col > 0) {
        col--; // 转换为0-based
        result = static_cast<char>('A' + (col % 26)) + result;
        col /= 26;
    }
    return result;
}

TXTypes::ColIndex TXTypes::colNameToIndex(const std::string& name) {
    if (name.empty()) {
        return INVALID_COL;
    }
    
    ColIndex result = 0;
    for (char c : name) {
        if (!std::isalpha(c)) {
            return INVALID_COL;
        }
        c = std::toupper(c);
        if (c < 'A' || c > 'Z') {
            return INVALID_COL;
        }
        result = result * 26 + (c - 'A' + 1);
    }
    
    return (result <= MAX_COLS) ? result : INVALID_COL;
}

std::string TXTypes::coordinateToAddress(RowIndex row, ColIndex col) {
    if (!isValidCoordinate(row, col)) {
        return "";
    }
    
    return colIndexToName(col) + std::to_string(row);
}

std::pair<TXTypes::RowIndex, TXTypes::ColIndex> TXTypes::addressToCoordinate(const std::string& address) {
    if (address.empty()) {
        return {INVALID_ROW, INVALID_COL};
    }
    
    // 分离字母和数字部分
    std::size_t i = 0;
    while (i < address.size() && std::isalpha(address[i])) {
        ++i;
    }
    
    if (i == 0 || i == address.size()) {
        return {INVALID_ROW, INVALID_COL};
    }
    
    std::string colPart = address.substr(0, i);
    std::string rowPart = address.substr(i);
    
    // 验证数字部分
    for (char c : rowPart) {
        if (!std::isdigit(c)) {
            return {INVALID_ROW, INVALID_COL};
        }
    }
    
    ColIndex col = colNameToIndex(colPart);
    RowIndex row = 0;
    
    try {
        unsigned long rowLong = std::stoul(rowPart);
        if (rowLong > MAX_ROWS) {
            return {INVALID_ROW, INVALID_COL};
        }
        row = static_cast<RowIndex>(rowLong);
    } catch (...) {
        return {INVALID_ROW, INVALID_COL};
    }
    
    if (row == 0 || col == INVALID_COL) {
        return {INVALID_ROW, INVALID_COL};
    }
    
    return {row, col};
}

TXTypes::ColorValue TXTypes::createColorFromHex(const std::string& hex) {
    std::string cleanHex = hex;
    
    // 移除开头的 #
    if (!cleanHex.empty() && cleanHex[0] == '#') {
        cleanHex = cleanHex.substr(1);
    }
    
    // 验证16进制字符
    for (char c : cleanHex) {
        if (!std::isxdigit(c)) {
            return DEFAULT_COLOR;
        }
    }
    
    // 支持的格式: RGB (6位), ARGB (8位)
    if (cleanHex.length() == 6) {
        // RGB 格式，添加不透明度
        cleanHex = "FF" + cleanHex;
    } else if (cleanHex.length() != 8) {
        return DEFAULT_COLOR;
    }
    
    try {
        unsigned long colorLong = std::stoul(cleanHex, nullptr, 16);
        return static_cast<ColorValue>(colorLong);
    } catch (...) {
        return DEFAULT_COLOR;
    }
}

} // namespace TinaXlsx 