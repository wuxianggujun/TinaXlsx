/**
 * @file Types.hpp
 * @brief TinaXlsx basic type definitions
 */

#pragma once

#include <string>
#include <vector>
#include <variant>
#include <optional>

namespace TinaXlsx {

// ================================
// Platform-specific type definitions
// ================================

#if defined(_WIN32) || defined(_WIN64)
    // Windows platform
    using Integer = long long;           // 64-bit integer
    using UInteger = unsigned long long; // 64-bit unsigned integer
    using UInt8 = unsigned char;         // 8-bit unsigned integer
    using UInt32 = unsigned int;         // 32-bit unsigned integer
#elif defined(__linux__) || defined(__APPLE__)
    // Linux/macOS platform, prefer standard types
    #include <cstdint>
    using Integer = int64_t;
    using UInteger = uint64_t;
    using UInt8 = uint8_t;
    using UInt32 = uint32_t;
#else
    // Other platforms, use conservative type definitions
    using Integer = long long;
    using UInteger = unsigned long long;
    using UInt8 = unsigned char;
    using UInt32 = unsigned int;
#endif

// ================================
// Excel-related type definitions
// ================================

/**
 * @brief Row index type (0-based)
 */
using RowIndex = UInt32;

/**
 * @brief Column index type (0-based)
 */
using ColumnIndex = UInt32;

/**
 * @brief Worksheet index type
 */
using SheetIndex = size_t;

/**
 * @brief Color type (RGB value)
 */
using Color = UInt32;

// ================================
// Cell data types
// ================================

/**
 * @brief Cell data type - high performance variant
 */
using CellValue = std::variant<
    std::string,    // String - index 0
    double,         // Float - index 1
    Integer,        // Integer - index 2
    bool,           // Boolean - index 3
    std::monostate  // Empty - index 4
>;

/**
 * @brief CellValue type index enum, ensures type safety and high performance
 */
enum class CellValueType : UInt8 {
    String = 0,      // std::string
    Double = 1,      // double  
    Integer = 2,     // Integer (int64_t/long long)
    Boolean = 3,     // bool
    Empty = 4        // std::monostate
};

/**
 * @brief Inline high performance type getter function
 * @param value CellValue instance
 * @return CellValueType corresponding type enum
 */
[[nodiscard]] inline CellValueType getCellValueType(const CellValue& value) noexcept {
    return static_cast<CellValueType>(value.index());
}

// ================================
// High performance string conversion utilities
// ================================

/**
 * @brief High performance integer to string conversion - avoid std::to_string overhead
 */
namespace FastConvert {
    
    /**
     * @brief Fast integer to string - use stack buffer to avoid heap allocation
     */
    [[nodiscard]] std::string integerToString(Integer value) noexcept;
    
    /**
     * @brief Fast double to string - optimize decimal places and scientific notation
     */
    [[nodiscard]] std::string doubleToString(double value) noexcept;
    
    /**
     * @brief Fast boolean to string - compile-time optimization
     */
    [[nodiscard]] constexpr const char* boolToString(bool value) noexcept {
        return value ? "true" : "false";
    }
}

// ================================
// CellValue high performance conversion functions
// ================================

/**
 * @brief High performance CellValue to string conversion
 * @param value CellValue instance
 * @return std::string converted string
 */
[[nodiscard]] std::string cellValueToString(const CellValue& value) noexcept;

/**
 * @brief Excel row data type
 */
using RowData = std::vector<CellValue>;

/**
 * @brief Excel table data type
 */
using TableData = std::vector<RowData>;

/**
 * @brief Border style enum
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
 * @brief Alignment enum - matches libxlsxwriter LXW_ALIGN_* values
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
 * @brief Vertical alignment enum - matches libxlsxwriter LXW_ALIGN_VERTICAL_* values
 */
enum class VerticalAlignment : UInt8 {
    Top = 8,          // LXW_ALIGN_VERTICAL_TOP
    Bottom = 9,       // LXW_ALIGN_VERTICAL_BOTTOM
    VCenter = 10,     // LXW_ALIGN_VERTICAL_CENTER
    VJustify = 11,    // LXW_ALIGN_VERTICAL_JUSTIFY
    VDistributed = 12 // LXW_ALIGN_VERTICAL_DISTRIBUTED
};

/**
 * @brief Cell position structure
 */
struct CellPosition {
    RowIndex row;
    ColumnIndex column;
    
    constexpr CellPosition() noexcept : row(0), column(0) {}
    constexpr CellPosition(RowIndex r, ColumnIndex c) noexcept : row(r), column(c) {}
    
    // Comparison operators
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
 * @brief Cell range structure
 */
struct CellRange {
    CellPosition start;
    CellPosition end;
    
    constexpr CellRange() noexcept = default;
    constexpr CellRange(const CellPosition& s, const CellPosition& e) noexcept : start(s), end(e) {}
    constexpr CellRange(RowIndex startRow, ColumnIndex startCol, RowIndex endRow, ColumnIndex endCol) noexcept
        : start(startRow, startCol), end(endRow, endCol) {}
    
    /**
     * @brief Check if range is valid
     */
    [[nodiscard]] constexpr bool isValid() const noexcept {
        return start.row <= end.row && start.column <= end.column;
    }
    
    /**
     * @brief Get number of rows in range
     */
    [[nodiscard]] constexpr RowIndex rowCount() const noexcept {
        return isValid() ? (end.row - start.row + 1) : 0;
    }
    
    /**
     * @brief Get number of columns in range
     */
    [[nodiscard]] constexpr ColumnIndex columnCount() const noexcept {
        return isValid() ? (end.column - start.column + 1) : 0;
    }
    
    /**
     * @brief Check if position is within range
     */
    [[nodiscard]] constexpr bool contains(const CellPosition& pos) const noexcept {
        return pos.row >= start.row && pos.row <= end.row &&
               pos.column >= start.column && pos.column <= end.column;
    }
};

/**
 * @brief Predefined colors
 */
namespace Colors {
    constexpr Color White = 0xFFFFFF;
    constexpr Color Black = 0x000000;
    constexpr Color Red = 0xFF0000;
    constexpr Color Green = 0x00FF00;
    constexpr Color Blue = 0x0000FF;
    constexpr Color Yellow = 0xFFFF00;
    constexpr Color Cyan = 0x00FFFF;
    constexpr Color Magenta = 0xFF00FF;
    constexpr Color Gray = 0x808080;
    constexpr Color LightGray = 0xC0C0C0;
    constexpr Color DarkGray = 0x404040;
    constexpr Color Silver = 0xC0C0C0;
    constexpr Color Maroon = 0x800000;
    constexpr Color Olive = 0x808000;
    constexpr Color Navy = 0x000080;
    constexpr Color Purple = 0x800080;
    constexpr Color Teal = 0x008080;
    constexpr Color Lime = 0x00FF00;
    constexpr Color Aqua = 0x00FFFF;
    constexpr Color Fuchsia = 0xFF00FF;
    constexpr Color Orange = 0xFFA500;
    constexpr Color Pink = 0xFFC0CB;
    constexpr Color Brown = 0xA52A2A;
    constexpr Color Gold = 0xFFD700;
    constexpr Color Violet = 0xEE82EE;
    constexpr Color Indigo = 0x4B0082;
    constexpr Color Turquoise = 0x40E0D0;
    constexpr Color Coral = 0xFF7F50;
    constexpr Color Salmon = 0xFA8072;
    constexpr Color Khaki = 0xF0E68C;
    constexpr Color Lavender = 0xE6E6FA;
    constexpr Color Peach = 0xFFDAB9;
    constexpr Color Mint = 0x98FB98;
    constexpr Color Wheat = 0xF5DEB3;
    constexpr Color Gray25 = 0xC0C0C0;
    constexpr Color Gray50 = 0x808080;
    constexpr Color Gray75 = 0x404040;
}

/**
 * @brief Worksheet options structure
 */
struct WorksheetOptions {
    std::optional<double> defaultRowHeight;
    std::optional<double> defaultColumnWidth;
    bool showGridlines = true;
    bool showHeaders = true;
    bool rightToLeft = false;
    std::optional<Color> tabColor;
};

// ================================
// Utility functions
// ================================

/**
 * @brief Convert column index to Excel column name (A, B, C, ..., AA, AB, ...)
 */
[[nodiscard]] std::string columnIndexToName(ColumnIndex column);

/**
 * @brief Convert Excel column name to column index
 */
[[nodiscard]] std::optional<ColumnIndex> columnNameToIndex(const std::string& columnName);

/**
 * @brief Convert cell position to Excel reference string (A1, B2, etc.)
 */
[[nodiscard]] std::string cellPositionToString(const CellPosition& pos);

/**
 * @brief Convert Excel reference string to cell position
 */
[[nodiscard]] std::optional<CellPosition> stringToCellPosition(const std::string& cellRef);

} // namespace TinaXlsx 
