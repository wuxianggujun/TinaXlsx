/**
 * @file Types.cpp
 * @brief TinaXlsx基础类型实现
 */

#include "TinaXlsx/Types.hpp"
#include <algorithm>
#include <cctype>
#include <regex>
#include <sstream>

namespace TinaXlsx {

std::string columnIndexToName(ColumnIndex column) {
    std::string result;
    ++column; // 转换为1基于
    
    while (column > 0) {
        --column; // 转换回0基于进行计算
        result = static_cast<char>('A' + (column % 26)) + result;
        column /= 26;
    }
    
    return result;
}

std::optional<ColumnIndex> columnNameToIndex(const std::string& columnName) {
    if (columnName.empty()) {
        return std::nullopt;
    }
    
    ColumnIndex result = 0;
    for (char c : columnName) {
        if (c < 'A' || c > 'Z') {
            // 尝试转换小写字母
            if (c >= 'a' && c <= 'z') {
                c = c - 'a' + 'A';
            } else {
                return std::nullopt; // 非法字符
            }
        }
        result = result * 26 + (c - 'A' + 1);
    }
    
    return result - 1; // 转换为0基于
}

std::string cellPositionToString(const CellPosition& pos) {
    return columnIndexToName(pos.column) + std::to_string(pos.row + 1);
}

std::optional<CellPosition> stringToCellPosition(const std::string& cellRef) {
    if (cellRef.empty()) {
        return std::nullopt;
    }
    
    // 使用正则表达式解析单元格引用 (如 A1, B10, AA100)
    static const std::regex cellRegex(R"(^([A-Za-z]+)(\d+)$)");
    std::smatch match;
    
    if (!std::regex_match(cellRef, match, cellRegex)) {
        return std::nullopt;
    }
    
    std::string columnStr = match[1].str();
    std::string rowStr = match[2].str();
    
    // 转换列名为列索引
    auto columnIndex = columnNameToIndex(columnStr);
    if (!columnIndex) {
        return std::nullopt;
    }
    
    // 转换行号
    try {
        int rowNumber = std::stoi(rowStr);
        if (rowNumber <= 0) {
            return std::nullopt;
        }
        return CellPosition(static_cast<RowIndex>(rowNumber - 1), *columnIndex);
    } catch (const std::exception&) {
        return std::nullopt;
    }
}

} // namespace TinaXlsx 
