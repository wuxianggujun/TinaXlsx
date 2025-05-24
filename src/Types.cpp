/**
 * @file Types.cpp
 * @brief TinaXlsx基础类型实现
 */

#include "TinaXlsx/Types.hpp"
#include <charconv>
#include <cstring>
#include <algorithm>
#include <cmath>
#include <limits>
#include <cstdio>
#include <cctype>

namespace TinaXlsx {

// ================================
// 高性能字符串转换实现
// ================================

namespace FastConvert {

std::string integerToString(Integer value) noexcept {
    // 使用栈上缓冲区避免堆分配，64位整数最多20个字符 + 符号
    char buffer[32];
    
    // 使用std::to_chars进行最快的转换
    auto [ptr, ec] = std::to_chars(buffer, buffer + sizeof(buffer), value);
    
    if (ec == std::errc{}) {
        return std::string(buffer, ptr - buffer);
    }
    
    // 备用方案 - 手动转换（理论上不会到达这里）
    if (value == 0) return "0";
    
    bool negative = value < 0;
    if (negative) value = -value;
    
    char* end = buffer + sizeof(buffer) - 1;
    char* start = end;
    *end = '\0';
    
    while (value > 0) {
        *(--start) = '0' + (value % 10);
        value /= 10;
    }
    
    if (negative) {
        *(--start) = '-';
    }
    
    return std::string(start, end - start);
}

std::string doubleToString(double value) noexcept {
    // 特殊值处理
    if (std::isnan(value)) return "NaN";
    if (std::isinf(value)) return value > 0 ? "Infinity" : "-Infinity";
    if (value == 0.0) return "0";
    
    // 检查是否为整数值 - 避免不必要的小数点
    if (value == std::floor(value) && 
        value >= static_cast<double>(std::numeric_limits<Integer>::min()) &&
        value <= static_cast<double>(std::numeric_limits<Integer>::max())) {
        return integerToString(static_cast<Integer>(value));
    }
    
    // 使用固定缓冲区
    char buffer[64];
    
    // 优先使用std::to_chars（C++17）
    auto [ptr, ec] = std::to_chars(buffer, buffer + sizeof(buffer), value);
    
    if (ec == std::errc{}) {
        return std::string(buffer, ptr - buffer);
    }
    
    // 备用方案 - 使用snprintf
    int len = std::snprintf(buffer, sizeof(buffer), "%.15g", value);
    if (len > 0 && len < static_cast<int>(sizeof(buffer))) {
        return std::string(buffer, len);
    }
    
    // 最后的备用方案
    return std::to_string(value);
}

} // namespace FastConvert

// ================================
// 高性能CellValue转换实现
// ================================

std::string cellValueToString(const CellValue& value) noexcept {
    // 使用内联的switch语句替代std::visit，获得最佳性能
    const CellValueType type = getCellValueType(value);
    
    switch (type) {
        case CellValueType::String:
            return std::get<std::string>(value);
            
        case CellValueType::Double:
            return FastConvert::doubleToString(std::get<double>(value));
            
        case CellValueType::Integer:
            return FastConvert::integerToString(std::get<Integer>(value));
            
        case CellValueType::Boolean:
            return FastConvert::boolToString(std::get<bool>(value));
            
        case CellValueType::Empty:
        default:
            return "";
    }
}

// ================================
// Excel工具函数实现
// ================================

std::string columnIndexToName(ColumnIndex column) {
    std::string result;
    
    // Excel列名转换算法 - 类似26进制但从1开始
    while (true) {
        result = static_cast<char>('A' + (column % 26)) + result;
        if (column < 26) break;
        column = (column / 26) - 1;
    }
    
    return result;
}

std::optional<ColumnIndex> columnNameToIndex(const std::string& columnName) {
    if (columnName.empty()) return std::nullopt;
    
    ColumnIndex result = 0;
    
    for (char c : columnName) {
        if (c < 'A' || c > 'Z') {
            // 尝试转换小写
            if (c >= 'a' && c <= 'z') {
                c = c - 'a' + 'A';
            } else {
                return std::nullopt;
            }
        }
        
        result = result * 26 + (c - 'A' + 1);
    }
    
    return result - 1; // 转换为0基准
}

std::string cellPositionToString(const CellPosition& pos) {
    return columnIndexToName(pos.column) + std::to_string(pos.row + 1);
}

std::optional<CellPosition> stringToCellPosition(const std::string& cellRef) {
    if (cellRef.empty()) return std::nullopt;
    
    // 分离字母和数字部分
    size_t numStart = 0;
    while (numStart < cellRef.length() && std::isalpha(cellRef[numStart])) {
        ++numStart;
    }
    
    if (numStart == 0 || numStart == cellRef.length()) {
        return std::nullopt;
    }
    
    std::string columnPart = cellRef.substr(0, numStart);
    std::string rowPart = cellRef.substr(numStart);
    
    // 转换列名
    auto column = columnNameToIndex(columnPart);
    if (!column) return std::nullopt;
    
    // 转换行号
    try {
        RowIndex row = std::stoul(rowPart);
        if (row == 0) return std::nullopt; // Excel行号从1开始
        
        return CellPosition{row - 1, *column}; // 转换为0基准
    } catch (const std::exception&) {
        return std::nullopt;
    }
}

} // namespace TinaXlsx 
