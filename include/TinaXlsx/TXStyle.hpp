#pragma once

#include "TXTypes.hpp"
#include "TXColor.hpp"
#include <string>
#include <memory>

namespace TinaXlsx {

/**
 * @brief 字体样式枚举
 */
enum class FontStyle : uint8_t {
    Normal = 0,           ///< 正常
    Bold = 1,             ///< 粗体
    Italic = 2,           ///< 斜体
    BoldItalic = 3,       ///< 粗斜体
    Underline = 4,        ///< 下划线
    Strikethrough = 8     ///< 删除线
};

/**
 * @brief 水平对齐方式
 */
enum class HorizontalAlignment : uint8_t {
    Left = 0,             ///< 左对齐
    Center = 1,           ///< 居中
    Right = 2,            ///< 右对齐
    Justify = 3,          ///< 两端对齐
    Fill = 4,             ///< 填充
    CenterAcrossSelection = 5,  ///< 跨列居中
    General = 6           ///< 常规
};

/**
 * @brief 垂直对齐方式
 */
enum class VerticalAlignment : uint8_t {
    Top = 0,              ///< 顶端对齐
    Middle = 1,           ///< 居中
    Bottom = 2,           ///< 底端对齐
    Justify = 3,          ///< 两端对齐
    Distributed = 4       ///< 分散对齐
};

/**
 * @brief 边框样式
 */
enum class BorderStyle : uint8_t {
    None = 0,             ///< 无边框
    Thin = 1,             ///< 细边框
    Medium = 2,           ///< 中等边框
    Thick = 3,            ///< 粗边框
    Double = 4,           ///< 双线边框
    Dotted = 5,           ///< 点线边框
    Dashed = 6,           ///< 虚线边框
    DashDot = 7,          ///< 点划线边框
    DashDotDot = 8        ///< 双点划线边框
};

/**
 * @brief 填充模式
 */
enum class FillPattern : uint8_t {
    None = 0,             ///< 无填充
    Solid = 1,            ///< 实心填充
    Gray75 = 2,           ///< 75%灰色
    Gray50 = 3,           ///< 50%灰色
    Gray25 = 4,           ///< 25%灰色
    Gray125 = 5,          ///< 12.5%灰色
    Gray0625 = 6          ///< 6.25%灰色
};

/**
 * @brief 字体信息结构
 */
struct TXFont {
    std::string name;                    ///< 字体名称
    font_size_t size;              ///< 字体大小
    TXColor color;                       ///< 字体颜色
    FontStyle style;                     ///< 字体样式
    bool bold;                     ///< 是否粗体
    bool italic;                   ///< 是否斜体
    bool underline;                ///< 是否下划线
    bool strikethrough;            ///< 是否删除线
    
    /**
     * @brief 默认构造函数
     */
    TXFont() : name("Calibri")
        , size(DEFAULT_FONT_SIZE)
        , color(ColorConstants::BLACK)
        , style(FontStyle::Normal)
        , bold(false), italic(false), underline(false), strikethrough(false) {}
    
    /**
     * @brief 构造函数
     * @param fontName 字体名称
     * @param fontSize 字体大小
     */
    TXFont(const std::string& fontName, font_size_t fontSize = DEFAULT_FONT_SIZE)
        : name(fontName), size(fontSize)
        , color(ColorConstants::BLACK)
        , style(FontStyle::Normal)
        , bold(false), italic(false), underline(false), strikethrough(false) {}
    
    // 便捷方法
    TXFont& setName(const std::string& fontName) { name = fontName; return *this; }
    TXFont& setSize(font_size_t fontSize) { size = fontSize; return *this; }
    TXFont& setColor(color_value_t colorValue) { color = TXColor(colorValue); return *this; }
    TXFont& setStyle(FontStyle fontStyle) { 
        style = fontStyle; 
        // 同步布尔字段
        bold = (fontStyle == FontStyle::Bold || fontStyle == FontStyle::BoldItalic);
        italic = (fontStyle == FontStyle::Italic || fontStyle == FontStyle::BoldItalic);
        underline = (fontStyle == FontStyle::Underline);
        strikethrough = (fontStyle == FontStyle::Strikethrough);
        return *this; 
    }
    TXFont& setBold(bool bold = true) { this->bold = bold; return *this; }
    TXFont& setItalic(bool italic = true) { this->italic = italic; return *this; }
    TXFont& setUnderline(bool underline = true) { this->underline = underline; return *this; }
    TXFont& setStrikethrough(bool strikethrough = true) { this->strikethrough = strikethrough; return *this; }
    
    // 查询方法
    bool isBold() const { return bold; }
    bool isItalic() const { return italic; }
    bool hasUnderline() const { return underline; }
    bool hasStrikethrough() const { return strikethrough; }
    
    bool operator==(const TXFont& other) const;
    bool operator!=(const TXFont& other) const { return !(*this == other); }
};

/**
 * @brief 对齐信息结构
 */
struct TXAlignment {
    HorizontalAlignment horizontal;      ///< 水平对齐
    VerticalAlignment vertical;          ///< 垂直对齐
    u32 textRotation;      ///< 文字旋转角度 (0-180)
    u32 indent;            ///< 缩进级别
    bool wrapText;                  ///< 是否自动换行
    bool shrinkToFit;               ///< 是否缩小字体以适应
    
    /**
     * @brief 默认构造函数
     */
    TXAlignment() 
        : horizontal(HorizontalAlignment::Left)
        , vertical(VerticalAlignment::Bottom)
        , textRotation(0)
        , indent(0)
        , wrapText(false)
        , shrinkToFit(false) {}
    
    // 便捷方法
    TXAlignment& setHorizontal(HorizontalAlignment align) { horizontal = align; return *this; }
    TXAlignment& setVertical(VerticalAlignment align) { vertical = align; return *this; }
    TXAlignment& setWrapText(bool wrap) { wrapText = wrap; return *this; }
    TXAlignment& setShrinkToFit(bool shrink) { shrinkToFit = shrink; return *this; }
    TXAlignment& setTextRotation(u32 rotation) { textRotation = rotation; return *this; }
    TXAlignment& setIndent(u32 indentLevel) { indent = indentLevel; return *this; }
    
    bool operator==(const TXAlignment& other) const;
    bool operator!=(const TXAlignment& other) const { return !(*this == other); }
};

/**
 * @brief 边框信息结构
 */
struct TXBorder {
    BorderStyle leftStyle;               ///< 左边框样式
    BorderStyle rightStyle;              ///< 右边框样式
    BorderStyle topStyle;                ///< 上边框样式
    BorderStyle bottomStyle;             ///< 下边框样式
    BorderStyle diagonalStyle;           ///< 对角线样式
    
    TXColor leftColor;                   ///< 左边框颜色
    TXColor rightColor;                  ///< 右边框颜色
    TXColor topColor;                    ///< 上边框颜色
    TXColor bottomColor;                 ///< 下边框颜色
    TXColor diagonalColor;               ///< 对角线颜色
    
    bool diagonalUp;                     ///< 是否显示左下到右上对角线
    bool diagonalDown;                   ///< 是否显示左上到右下对角线
    
    TXBorder()
        : leftStyle(BorderStyle::None)
        , rightStyle(BorderStyle::None)
        , topStyle(BorderStyle::None)
        , bottomStyle(BorderStyle::None)
        , diagonalStyle(BorderStyle::None)
        , leftColor(ColorConstants::BLACK)
        , rightColor(ColorConstants::BLACK)
        , topColor(ColorConstants::BLACK)
        , bottomColor(ColorConstants::BLACK)
        , diagonalColor(ColorConstants::BLACK)
        , diagonalUp(false)
        , diagonalDown(false)
    {}
    
    // 便捷方法
    TXBorder& setAllBorders(BorderStyle style, const TXColor& color = ColorConstants::BLACK);
    TXBorder& setLeftBorder(BorderStyle style, const TXColor& color = ColorConstants::BLACK);
    TXBorder& setRightBorder(BorderStyle style, const TXColor& color = ColorConstants::BLACK);
    TXBorder& setTopBorder(BorderStyle style, const TXColor& color = ColorConstants::BLACK);
    TXBorder& setBottomBorder(BorderStyle style, const TXColor& color = ColorConstants::BLACK);
    TXBorder& setDiagonalBorder(BorderStyle style, const TXColor& color = ColorConstants::BLACK, bool up = true, bool down = true);
    
    bool operator==(const TXBorder& other) const;
    bool operator!=(const TXBorder& other) const { return !(*this == other); }
};

/**
 * @brief 填充信息结构
 */
struct TXFill {
    FillPattern pattern;                 ///< 填充模式
    TXColor foregroundColor;             ///< 前景色
    TXColor backgroundColor;             ///< 背景色
    
    TXFill()
        : pattern(FillPattern::None)
        , foregroundColor(ColorConstants::BLACK)
        , backgroundColor(ColorConstants::WHITE)
    {}
    
    TXFill(FillPattern fillPattern, const TXColor& fgColor = ColorConstants::BLACK, const TXColor& bgColor = ColorConstants::WHITE)
        : pattern(fillPattern)
        , foregroundColor(fgColor)
        , backgroundColor(bgColor)
    {}
    
    // 便捷方法
    TXFill& setPattern(FillPattern fillPattern) { pattern = fillPattern; return *this; }
    TXFill& setForegroundColor(const TXColor& color) { foregroundColor = color; return *this; }
    TXFill& setBackgroundColor(const TXColor& color) { backgroundColor = color; return *this; }
    TXFill& setSolidFill(const TXColor& color);
    
    bool operator==(const TXFill& other) const;
    bool operator!=(const TXFill& other) const { return !(*this == other); }
};

/**
 * @brief 完整的单元格样式
 */
class TXCellStyle {
public:
    TXCellStyle();
    ~TXCellStyle();
    
    // 禁用拷贝构造，使用智能指针管理
    TXCellStyle(const TXCellStyle& other);
    TXCellStyle& operator=(const TXCellStyle& other);
    
    // 支持移动语义
    TXCellStyle(TXCellStyle&& other) noexcept;
    TXCellStyle& operator=(TXCellStyle&& other) noexcept;
    
    // ==================== 样式组件访问 ====================
    
    /**
     * @brief 获取字体样式
     * @return 字体样式引用
     */
    TXFont& getFont();
    const TXFont& getFont() const;
    
    /**
     * @brief 获取对齐样式
     * @return 对齐样式引用
     */
    TXAlignment& getAlignment();
    const TXAlignment& getAlignment() const;
    
    /**
     * @brief 获取边框样式
     * @return 边框样式引用
     */
    TXBorder& getBorder();
    const TXBorder& getBorder() const;
    
    /**
     * @brief 获取填充样式
     * @return 填充样式引用
     */
    TXFill& getFill();
    const TXFill& getFill() const;
    
    // ==================== 便捷设置方法 ====================
    
    /**
     * @brief 设置字体
     * @param font 字体信息
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setFont(const TXFont& font);
    
    /**
     * @brief 设置对齐
     * @param alignment 对齐信息
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setAlignment(const TXAlignment& alignment);
    
    /**
     * @brief 设置边框
     * @param border 边框信息
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setBorder(const TXBorder& border);
    
    /**
     * @brief 设置填充
     * @param fill 填充信息
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setFill(const TXFill& fill);
    
    // ==================== 快捷样式方法 ====================
    
    /**
     * @brief 设置字体名称和大小
     * @param name 字体名称
     * @param size 字体大小
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setFont(const std::string& name, font_size_t size = DEFAULT_FONT_SIZE);
    
    /**
     * @brief 设置字体颜色
     * @param color 颜色对象
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setFontColor(const TXColor& color);
    
    /**
     * @brief 设置字体颜色（从颜色值）
     * @param color 颜色值
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setFontColor(color_value_t color);
    
    /**
     * @brief 设置字体样式
     * @param style 字体样式
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setFontStyle(FontStyle style);
    
    /**
     * @brief 设置水平对齐
     * @param alignment 水平对齐方式
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setHorizontalAlignment(HorizontalAlignment alignment);
    
    /**
     * @brief 设置垂直对齐
     * @param alignment 垂直对齐方式
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setVerticalAlignment(VerticalAlignment alignment);
    
    /**
     * @brief 设置背景颜色
     * @param color 背景颜色
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setBackgroundColor(const TXColor& color);
    
    /**
     * @brief 设置背景颜色（从颜色值）
     * @param color 颜色值
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setBackgroundColor(color_value_t color);
    
    /**
     * @brief 设置所有边框
     * @param style 边框样式
     * @param color 边框颜色
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setAllBorders(BorderStyle style, const TXColor& color = ColorConstants::BLACK);
    
    /**
     * @brief 重置为默认样式
     */
    void reset();
    
    /**
     * @brief 比较两个样式是否相等
     * @param other 另一个样式
     * @return 相等返回true，否则返回false
     */
    bool operator==(const TXCellStyle& other) const;
    bool operator!=(const TXCellStyle& other) const { return !(*this == other); }

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

// ==================== 预定义样式 ====================
namespace Styles {
    /**
     * @brief 获取标题样式 (粗体、大字体、居中)
     * @return 标题样式
     */
    TXCellStyle createHeaderStyle();
    
    /**
     * @brief 获取数据样式 (普通字体、左对齐)
     * @return 数据样式
     */
    TXCellStyle createDataStyle();
    
    /**
     * @brief 获取数字样式 (右对齐)
     * @return 数字样式
     */
    TXCellStyle createNumberStyle();
    
    /**
     * @brief 获取强调样式 (粗体、彩色背景)
     * @param backgroundColor 背景颜色
     * @return 强调样式
     */
    TXCellStyle createHighlightStyle(const TXColor& backgroundColor = ColorConstants::YELLOW);
    
    /**
     * @brief 获取边框表格样式
     * @return 边框表格样式
     */
    TXCellStyle createTableStyle();
}

} // namespace TinaXlsx 