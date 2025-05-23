/**
 * @file Types.hpp
 * @brief TinaXlsx基础类型定义
 */

#pragma once

#include <string>
#include <vector>
#include <variant>
#include <optional>
#include <cstdint>

namespace TinaXlsx {

/**
 * @brief 单元格数据类型
 */
using CellValue = std::variant<
    std::string,    // 字符串
    double,         // 数字
    int64_t,        // 整数
    bool,           // 布尔值
    std::monostate  // 空值
>;

/**
 * @brief Excel行数据类型
 */
using RowData = std::vector<CellValue>;

/**
 * @brief Excel表格数据类型
 */
using TableData = std::vector<RowData>;

/**
 * @brief 行索引类型（0基于）
 */
using RowIndex = uint32_t;

/**
 * @brief 列索引类型（0基于）
 */
using ColumnIndex = uint32_t;

/**
 * @brief 工作表索引类型
 */
using SheetIndex = size_t;

/**
 * @brief 颜色类型（RGB值）
 */
using Color = uint32_t;

/**
 * @brief 边框样式枚举
 */
enum class BorderStyle : uint8_t {
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
enum class Alignment : uint8_t {
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
enum class VerticalAlignment : uint8_t {
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