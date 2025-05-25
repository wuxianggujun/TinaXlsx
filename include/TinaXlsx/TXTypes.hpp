#pragma once

#include <cstdint>
#include <string>

namespace TinaXlsx {

/**
 * @brief 全局类型定义类
 * 
 * 统一管理所有基础类型定义，便于维护和修改
 */
class TXTypes {
public:
    // ==================== 基础数值类型 ====================
    using RowIndex = uint32_t;        ///< 行索引类型 (1-based)
    using ColIndex = uint32_t;        ///< 列索引类型 (1-based)
    using CellIndex = uint64_t;       ///< 单元格索引类型
    using SheetIndex = uint16_t;      ///< 工作表索引类型
    
    // ==================== 样式相关类型 ====================
    using FontSize = uint16_t;        ///< 字体大小类型 (点数)
    using ColorValue = uint32_t;      ///< 颜色值类型 (ARGB)
    using BorderWidth = uint8_t;      ///< 边框宽度类型
    using StyleId = uint32_t;         ///< 样式ID类型
    
    // ==================== 文件相关类型 ====================
    using FileSize = uint64_t;        ///< 文件大小类型
    using CompressionLevel = int8_t;  ///< 压缩级别类型 (-1 到 9)
    
    // ==================== 常量定义 ====================
    static constexpr RowIndex MAX_ROWS = 1048576;     ///< Excel最大行数
    static constexpr ColIndex MAX_COLS = 16384;       ///< Excel最大列数
    static constexpr SheetIndex MAX_SHEETS = 255;     ///< 最大工作表数
    static constexpr FontSize DEFAULT_FONT_SIZE = 11; ///< 默认字体大小
    static constexpr ColorValue DEFAULT_COLOR = 0xFF000000; ///< 默认颜色(黑色)
    
    // ==================== 无效值定义 ====================
    static constexpr RowIndex INVALID_ROW = 0;        ///< 无效行索引
    static constexpr ColIndex INVALID_COL = 0;        ///< 无效列索引
    static constexpr StyleId INVALID_STYLE_ID = 0;    ///< 无效样式ID
    
    // ==================== 工具函数 ====================
    
    /**
     * @brief 检查行索引是否有效
     * @param row 行索引
     * @return 有效返回true，无效返回false
     */
    static bool isValidRow(RowIndex row) {
        return row > 0 && row <= MAX_ROWS;
    }
    
    /**
     * @brief 检查列索引是否有效
     * @param col 列索引
     * @return 有效返回true，无效返回false
     */
    static bool isValidCol(ColIndex col) {
        return col > 0 && col <= MAX_COLS;
    }
    
    /**
     * @brief 检查坐标是否有效
     * @param row 行索引
     * @param col 列索引
     * @return 有效返回true，无效返回false
     */
    static bool isValidCoordinate(RowIndex row, ColIndex col) {
        return isValidRow(row) && isValidCol(col);
    }
    
    /**
     * @brief 列索引转换为Excel列名 (A, B, C, ..., AA, AB, ...)
     * @param col 列索引 (1-based)
     * @return Excel列名
     */
    static std::string colIndexToName(ColIndex col);
    
    /**
     * @brief Excel列名转换为列索引
     * @param name Excel列名 (A, B, C, ..., AA, AB, ...)
     * @return 列索引 (1-based)，无效返回INVALID_COL
     */
    static ColIndex colNameToIndex(const std::string& name);
    
    /**
     * @brief 坐标转换为A1格式地址
     * @param row 行索引 (1-based)
     * @param col 列索引 (1-based)
     * @return A1格式地址 (如 "A1", "B5", "AA10")
     */
    static std::string coordinateToAddress(RowIndex row, ColIndex col);
    
    /**
     * @brief A1格式地址转换为坐标
     * @param address A1格式地址 (如 "A1", "B5", "AA10")
     * @return 坐标对 {row, col}，无效返回 {INVALID_ROW, INVALID_COL}
     */
    static std::pair<RowIndex, ColIndex> addressToCoordinate(const std::string& address);
    
    /**
     * @brief 创建RGB颜色值
     * @param r 红色分量 (0-255)
     * @param g 绿色分量 (0-255)
     * @param b 蓝色分量 (0-255)
     * @param a 透明度分量 (0-255, 默认255不透明)
     * @return ARGB颜色值
     */
    static ColorValue createColor(uint8_t r, uint8_t g, uint8_t b, uint8_t a = 255) {
        return (static_cast<ColorValue>(a) << 24) |
               (static_cast<ColorValue>(r) << 16) |
               (static_cast<ColorValue>(g) << 8)  |
               static_cast<ColorValue>(b);
    }
    
    /**
     * @brief 从16进制字符串创建颜色值
     * @param hex 16进制颜色字符串 (如 "#FF0000", "FF0000", "#FFFF0000")
     * @return 颜色值
     */
    static ColorValue createColorFromHex(const std::string& hex);
    
    /**
     * @brief 提取颜色分量
     * @param color 颜色值
     * @return {r, g, b, a}
     */
    static std::tuple<uint8_t, uint8_t, uint8_t, uint8_t> extractColorComponents(ColorValue color) {
        uint8_t a = static_cast<uint8_t>((color >> 24) & 0xFF);
        uint8_t r = static_cast<uint8_t>((color >> 16) & 0xFF);
        uint8_t g = static_cast<uint8_t>((color >> 8) & 0xFF);
        uint8_t b = static_cast<uint8_t>(color & 0xFF);
        return {r, g, b, a};
    }
};

// ==================== 常用颜色预定义 ====================
namespace Colors {
    constexpr TXTypes::ColorValue BLACK   = 0xFF000000;
    constexpr TXTypes::ColorValue WHITE   = 0xFFFFFFFF;
    constexpr TXTypes::ColorValue RED     = 0xFFFF0000;
    constexpr TXTypes::ColorValue GREEN   = 0xFF00FF00;
    constexpr TXTypes::ColorValue BLUE    = 0xFF0000FF;
    constexpr TXTypes::ColorValue YELLOW  = 0xFFFFFF00;
    constexpr TXTypes::ColorValue CYAN    = 0xFF00FFFF;
    constexpr TXTypes::ColorValue MAGENTA = 0xFFFF00FF;
    constexpr TXTypes::ColorValue GRAY    = 0xFF808080;
    constexpr TXTypes::ColorValue DARK_GRAY = 0xFF404040;
    constexpr TXTypes::ColorValue LIGHT_GRAY = 0xFFC0C0C0;
}

} // namespace TinaXlsx 