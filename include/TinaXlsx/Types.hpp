/**
 * @file Types.hpp
 * @brief TinaXlsx基础类型定义
 */

#pragma once

#include <string>
#include <vector>
#include <variant>
#include <optional>

namespace TinaXlsx {

// ================================
// 平台相关类型定义
// ================================

#if defined(_WIN32) || defined(_WIN64)
    // Windows平台
    using Integer = long long;           // 64位整数
    using UInteger = unsigned long long; // 64位无符号整数
    using UInt8 = unsigned char;         // 8位无符号整数
    using UInt32 = unsigned int;         // 32位无符号整数
#elif defined(__linux__) || defined(__APPLE__)
    // Linux/macOS平台，优先使用标准类型
    #include <cstdint>
    using Integer = int64_t;
    using UInteger = uint64_t;
    using UInt8 = uint8_t;
    using UInt32 = uint32_t;
#else
    // 其他平台，使用保守的类型定义
    using Integer = long long;
    using UInteger = unsigned long long;
    using UInt8 = unsigned char;
    using UInt32 = unsigned int;
#endif

// ================================
// Excel相关类型定义
// ================================

/**
 * @brief 行索引类型（0基于）
 */
using RowIndex = UInt32;

/**
 * @brief 列索引类型（0基于）
 */
using ColumnIndex = UInt32;

/**
 * @brief 工作表索引类型
 */
using SheetIndex = size_t;

/**
 * @brief 颜色类型（RGB值）
 */
using Color = UInt32;

// ================================
// 单元格数据类型
// ================================

/**
 * @brief 单元格数据类型 - 高性能variant
 */
using CellValue = std::variant<
    std::string,    // 字符串 - 索引0
    double,         // 浮点数 - 索引1
    Integer,        // 整数 - 索引2
    bool,           // 布尔值 - 索引3
    std::monostate  // 空值 - 索引4
>;

/**
 * @brief CellValue类型索引枚举，确保类型安全和高性能
 */
enum class CellValueType : UInt8 {
    String = 0,      // std::string
    Double = 1,      // double  
    Integer = 2,     // Integer (int64_t/long long)
    Boolean = 3,     // bool
    Empty = 4        // std::monostate
};

/**
 * @brief 内联高性能类型获取函数
 * @param value CellValue实例
 * @return CellValueType 对应的类型枚举
 */
[[nodiscard]] inline CellValueType getCellValueType(const CellValue& value) noexcept {
    return static_cast<CellValueType>(value.index());
}

// ================================
// 高性能字符串转换工具
// ================================

/**
 * @brief 高性能整数到字符串转换 - 避免std::to_string的开销
 */
namespace FastConvert {
    
    /**
     * @brief 快速整数转字符串 - 使用栈上缓冲区避免堆分配
     */
    [[nodiscard]] std::string integerToString(Integer value) noexcept;
    
    /**
     * @brief 快速双精度转字符串 - 优化小数位数和科学计数法
     */
    [[nodiscard]] std::string doubleToString(double value) noexcept;
    
    /**
     * @brief 快速布尔值转字符串 - 编译时优化
     */
    [[nodiscard]] constexpr const char* boolToString(bool value) noexcept {
        return value ? "true" : "false";
    }
}

// ================================
// CellValue高性能转换函数
// ================================

/**
 * @brief 高性能CellValue到字符串转换
 * @param value CellValue实例
 * @return std::string 转换后的字符串
 */
[[nodiscard]] std::string cellValueToString(const CellValue& value) noexcept;

/**
 * @brief Excel行数据类型
 */
using RowData = std::vector<CellValue>;

/**
 * @brief Excel表格数据类型
 */
using TableData = std::vector<RowData>;

/**
 * @brief 边框样式枚举
 */
enum class BorderStyle : UInt8 {
    None = 0,
    Thin = 1,
    Medium = 2,
    Thick = 3,
    Double = 4,
    Dotted = 5,
    Dashed = 6
};

/**
 * @brief 对齐方式枚举 - 匹配libxlsxwriter的LXW_ALIGN_*值
 */
enum class Alignment : UInt8 {
    None = 0,                    // LXW_ALIGN_NONE
    Left = 1,                    // LXW_ALIGN_LEFT
    Center = 2,                  // LXW_ALIGN_CENTER
    Right = 3,                   // LXW_ALIGN_RIGHT
    Fill = 4,                    // LXW_ALIGN_FILL
    Justify = 5,                 // LXW_ALIGN_JUSTIFY
    CenterAcrossSelection = 6,   // LXW_ALIGN_CENTER_ACROSS
    Distributed = 7              // LXW_ALIGN_DISTRIBUTED
};

/**
 * @brief 垂直对齐方式枚举 - 匹配libxlsxwriter的LXW_ALIGN_VERTICAL_*值
 */
enum class VerticalAlignment : UInt8 {
    Top = 8,          // LXW_ALIGN_VERTICAL_TOP
    Bottom = 9,       // LXW_ALIGN_VERTICAL_BOTTOM
    VCenter = 10,     // LXW_ALIGN_VERTICAL_CENTER
    VJustify = 11,    // LXW_ALIGN_VERTICAL_JUSTIFY
    VDistributed = 12 // LXW_ALIGN_VERTICAL_DISTRIBUTED
};

/**
 * @brief 单元格位置结构
 */
struct CellPosition {
    RowIndex row;
    ColumnIndex column;
    
    constexpr CellPosition() noexcept : row(0), column(0) {}
    constexpr CellPosition(RowIndex r, ColumnIndex c) noexcept : row(r), column(c) {}
    
    // 比较操作符
    constexpr bool operator==(const CellPosition& other) const noexcept {
        return row == other.row && column == other.column;
    }
    
    constexpr bool operator!=(const CellPosition& other) const noexcept {
        return !(*this == other);
    }
    
    constexpr bool operator<(const CellPosition& other) const noexcept {
        return row < other.row || (row == other.row && column < other.column);
    }
};

/**
 * @brief 单元格范围结构
 */
struct CellRange {
    CellPosition start;
    CellPosition end;
    
    constexpr CellRange() noexcept = default;
    constexpr CellRange(const CellPosition& s, const CellPosition& e) noexcept : start(s), end(e) {}
    constexpr CellRange(RowIndex startRow, ColumnIndex startCol, RowIndex endRow, ColumnIndex endCol) noexcept
        : start(startRow, startCol), end(endRow, endCol) {}
    
    /**
     * @brief 检查范围是否有效
     */
    [[nodiscard]] constexpr bool isValid() const noexcept {
        return start.row <= end.row && start.column <= end.column;
    }
    
    /**
     * @brief 获取范围的行数
     */
    [[nodiscard]] constexpr RowIndex rowCount() const noexcept {
        return isValid() ? (end.row - start.row + 1) : 0;
    }
    
    /**
     * @brief 获取范围的列数
     */
    [[nodiscard]] constexpr ColumnIndex columnCount() const noexcept {
        return isValid() ? (end.column - start.column + 1) : 0;
    }
    
    /**
     * @brief 检查位置是否在范围内
     */
    [[nodiscard]] constexpr bool contains(const CellPosition& pos) const noexcept {
        return pos.row >= start.row && pos.row <= end.row &&
               pos.column >= start.column && pos.column <= end.column;
    }
};

/**
 * @brief 预定义的颜色常量
 */
namespace Colors {
    constexpr Color White      = 0xFFFFFF;
    constexpr Color Black      = 0x000000;
    constexpr Color Red        = 0xFF0000;
    constexpr Color Green      = 0x008000;
    constexpr Color Blue       = 0x0000FF;
    constexpr Color Yellow     = 0xFFFF00;
    constexpr Color Cyan       = 0x00FFFF;
    constexpr Color Magenta    = 0xFF00FF;
    constexpr Color Silver     = 0xC0C0C0;
    constexpr Color Gray       = 0x808080;
    constexpr Color Maroon     = 0x800000;
    constexpr Color Olive      = 0x808000;
    constexpr Color Lime       = 0x00FF00;
    constexpr Color Aqua       = 0x00FFFF;
    constexpr Color Teal       = 0x008080;
    constexpr Color Navy       = 0x000080;
    constexpr Color Fuchsia    = 0xFF00FF;
    constexpr Color Purple     = 0x800080;
}

/**
 * @brief 工作表选项
 */
struct WorksheetOptions {
    std::optional<double> defaultRowHeight;
    std::optional<double> defaultColumnWidth;
    bool showGridlines = true;
    bool showHeaders = true;
    bool rightToLeft = false;
    std::optional<Color> tabColor;
};

/**
 * @brief 列A到列名的转换（如A、B、..、Z、AA、AB、...）
 * @param column 列索引（0基于）
 * @return std::string 列名
 */
[[nodiscard]] std::string columnIndexToName(ColumnIndex column);

/**
 * @brief 列名到列索引的转换
 * @param columnName 列名
 * @return std::optional<ColumnIndex> 列索引，如果无效则返回空
 */
[[nodiscard]] std::optional<ColumnIndex> columnNameToIndex(const std::string& columnName);

/**
 * @brief 将单元格位置转换为Excel格式的字符串（如A1、B2等）
 * @param pos 单元格位置
 * @return std::string Excel格式的单元格引用
 */
[[nodiscard]] std::string cellPositionToString(const CellPosition& pos);

/**
 * @brief 将Excel格式的字符串解析为单元格位置
 * @param cellRef Excel格式的单元格引用
 * @return std::optional<CellPosition> 单元格位置，如果无效则返回空
 */
[[nodiscard]] std::optional<CellPosition> stringToCellPosition(const std::string& cellRef);

} // namespace TinaXlsx 
