#pragma once

#include "TinaXlsx/TXTypes.hpp"
#include "TinaXlsx/TXColor.hpp"
#include "TinaXlsx/TXFont.hpp"  // 修正路径，移到主目录
#include <string>
#include <memory>

#include "TXNumberFormat.hpp"
#include "TXNumberFormat.hpp"

namespace TinaXlsx
{
    class TXNumberFormat;
}

namespace TinaXlsx {

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
 * @brief 对齐信息结构
 */
struct TXAlignment {
    HorizontalAlignment horizontal;      ///< 水平对齐
    VerticalAlignment vertical;          ///< 垂直对齐
    uint32_t textRotation;               ///< 文字旋转角度 (0-180)
    uint32_t indent;                     ///< 缩进级别
    bool wrapText;                       ///< 是否自动换行
    bool shrinkToFit;                    ///< 是否缩小字体以适应
    
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
    TXAlignment& setTextRotation(uint32_t rotation) { textRotation = rotation; return *this; }
    TXAlignment& setIndent(uint32_t indentLevel) { indent = indentLevel; return *this; }
    
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
 * @brief 单元格样式类 - 整合各种样式组件
 * 
 * 现在使用独立的TXFont类而不是内联结构体
 */
class TXCellStyle {

public:
    // --- 新增: 数字格式定义结构体 ---
    struct NumberFormatDefinition
    {
        TXNumberFormat::FormatType types_ = TXNumberFormat::FormatType::General;
        
    };

    
public:
    // ==================== 构造函数 ====================
    
    /**
     * @brief 默认构造函数
     */
    TXCellStyle();
    
    /**
     * @brief 拷贝构造函数
     */
    TXCellStyle(const TXCellStyle& other);
    
    /**
     * @brief 移动构造函数
     */
    TXCellStyle(TXCellStyle&& other) noexcept;
    
    /**
     * @brief 拷贝赋值操作符
     */
    TXCellStyle& operator=(const TXCellStyle& other);
    
    /**
     * @brief 移动赋值操作符
     */
    TXCellStyle& operator=(TXCellStyle&& other) noexcept;
    
    /**
     * @brief 析构函数
     */
    ~TXCellStyle() = default;

    // ==================== 字体设置（使用新的TXFont类）====================
    
    /**
     * @brief 设置字体
     * @param font 字体对象
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setFont(const TXFont& font);
    
    /**
     * @brief 获取字体
     * @return 字体对象的常量引用
     */
    [[nodiscard]] const TXFont& getFont() const { return font_; }
    
    /**
     * @brief 获取字体（可修改）
     * @return 字体对象的引用
     */
    [[nodiscard]] TXFont& getFont() { return font_; }
    
    // 字体便捷方法
    TXCellStyle& setFontName(const std::string& name);
    TXCellStyle& setFontSize(font_size_t size);
    TXCellStyle& setFontColor(color_value_t color);
    TXCellStyle& setFontColor(const TXColor& color);
    TXCellStyle& setFontBold(bool bold = true);
    TXCellStyle& setFontItalic(bool italic = true);
    TXCellStyle& setFontStyle(FontStyle style);

    // ==================== 对齐设置 ====================
    
    /**
     * @brief 设置对齐方式
     * @param alignment 对齐对象
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setAlignment(const TXAlignment& alignment);
    
    /**
     * @brief 获取对齐方式
     * @return 对齐对象的常量引用
     */
    [[nodiscard]] const TXAlignment& getAlignment() const { return alignment_; }
    
    /**
     * @brief 获取对齐方式（可修改）
     * @return 对齐对象的引用
     */
    [[nodiscard]] TXAlignment& getAlignment() { return alignment_; }
    
    // 对齐便捷方法
    TXCellStyle& setHorizontalAlignment(HorizontalAlignment alignment);
    TXCellStyle& setVerticalAlignment(VerticalAlignment alignment);
    TXCellStyle& setWrapText(bool wrap = true);
    TXCellStyle& setTextRotation(uint32_t rotation);

    // ==================== 边框设置 ====================
    
    /**
     * @brief 设置边框
     * @param border 边框对象
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setBorder(const TXBorder& border);
    
    /**
     * @brief 获取边框
     * @return 边框对象的常量引用
     */
    [[nodiscard]] const TXBorder& getBorder() const { return border_; }
    
    /**
     * @brief 获取边框（可修改）
     * @return 边框对象的引用
     */
    [[nodiscard]] TXBorder& getBorder() { return border_; }
    
    // 边框便捷方法
    TXCellStyle& setAllBorders(BorderStyle style, const TXColor& color = ColorConstants::BLACK);
    TXCellStyle& setLeftBorder(BorderStyle style, const TXColor& color = ColorConstants::BLACK);
    TXCellStyle& setRightBorder(BorderStyle style, const TXColor& color = ColorConstants::BLACK);
    TXCellStyle& setTopBorder(BorderStyle style, const TXColor& color = ColorConstants::BLACK);
    TXCellStyle& setBottomBorder(BorderStyle style, const TXColor& color = ColorConstants::BLACK);

    // ==================== 填充设置 ====================
    
    /**
     * @brief 设置填充
     * @param fill 填充对象
     * @return 自身引用，支持链式调用
     */
    TXCellStyle& setFill(const TXFill& fill);
    
    /**
     * @brief 获取填充
     * @return 填充对象的常量引用
     */
    [[nodiscard]] const TXFill& getFill() const { return fill_; }
    
    /**
     * @brief 获取填充（可修改）
     * @return 填充对象的引用
     */
    [[nodiscard]] TXFill& getFill() { return fill_; }
    
    // 填充便捷方法
    TXCellStyle& setBackgroundColor(color_value_t color);
    TXCellStyle& setBackgroundColor(const TXColor& color);
    TXCellStyle& setFillPattern(FillPattern pattern);
    TXCellStyle& setSolidFill(const TXColor& color);

    // ==================== 比较操作符 ====================
    
    bool operator==(const TXCellStyle& other) const;
    bool operator!=(const TXCellStyle& other) const { return !(*this == other); }

    // ==================== 工具方法 ====================
    
    /**
     * @brief 重置为默认样式
     */
    void reset();
    
    /**
     * @brief 检查是否为默认样式
     * @return 是默认样式返回true
     */
    [[nodiscard]] bool isDefault() const;
    
    /**
     * @brief 获取样式的唯一标识键（用于样式管理器）
     * @return 唯一标识字符串
     */
    [[nodiscard]] std::string getUniqueKey() const;

private:
    // ==================== 私有成员（使用新的独立类）====================
    
    TXFont font_;                        ///< 字体样式（新的独立类）
    TXAlignment alignment_;              ///< 对齐方式
    TXBorder border_;                    ///< 边框样式
    TXFill fill_;                        ///< 填充样式
};

} // namespace TinaXlsx 
