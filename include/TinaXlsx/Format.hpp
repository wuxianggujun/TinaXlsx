/**
 * @file Format.hpp
 * @brief Excel格式处理
 */

#pragma once

#include "Types.hpp"
#include "Exception.hpp"
#include <memory>
#include <string>
#include <optional>

// 前向声明，避免在公共头文件中包含xlsxwriter.h
struct lxw_workbook;
struct lxw_format;

namespace TinaXlsx {

/**
 * @brief Excel格式类
 * 
 * 封装libxlsxwriter的格式功能，提供类型安全的格式设置
 */
class Format {
public:
    /**
     * @brief 数字格式类型
     */
    enum class NumberFormat {
        General,           // 常规
        Number,            // 数字
        Number2,           // 保留2位小数的数字
        Percentage,        // 百分比
        Date,             // 日期
        Time,             // 时间
        DateTime,         // 日期时间
        Currency,         // 货币
        Accounting,       // 会计
        Scientific,       // 科学计数法
        Text              // 文本
    };

private:
    struct Impl;
    std::unique_ptr<Impl> pImpl_;
    
public:
    /**
     * @brief 构造函数
     * @param workbook 关联的工作簿（内部使用）
     */
    explicit Format(lxw_workbook* workbook);
    
    /**
     * @brief 移动构造函数
     */
    Format(Format&& other) noexcept;
    
    /**
     * @brief 移动赋值操作符
     */
    Format& operator=(Format&& other) noexcept;
    
    /**
     * @brief 析构函数
     */
    ~Format();
    
    // 禁用拷贝构造和拷贝赋值
    Format(const Format&) = delete;
    Format& operator=(const Format&) = delete;
    
    /**
     * @brief 设置字体名称
     * @param fontName 字体名称
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setFontName(const std::string& fontName);
    
    /**
     * @brief 设置字体大小
     * @param fontSize 字体大小
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setFontSize(double fontSize);
    
    /**
     * @brief 设置字体颜色
     * @param color 字体颜色
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setFontColor(Color color);
    
    /**
     * @brief 设置是否粗体
     * @param bold 是否粗体
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setBold(bool bold = true);
    
    /**
     * @brief 设置是否斜体
     * @param italic 是否斜体
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setItalic(bool italic = true);
    
    /**
     * @brief 设置下划线
     * @param underline 是否有下划线
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setUnderline(bool underline = true);
    
    /**
     * @brief 设置删除线
     * @param strikeout 是否有删除线
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setStrikeout(bool strikeout = true);
    
    /**
     * @brief 设置背景颜色
     * @param color 背景颜色
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setBackgroundColor(Color color);
    
    /**
     * @brief 设置前景颜色（用于填充图案）
     * @param color 前景颜色
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setForegroundColor(Color color);
    
    /**
     * @brief 设置水平对齐方式
     * @param alignment 水平对齐方式
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setAlignment(Alignment alignment);
    
    /**
     * @brief 设置垂直对齐方式
     * @param alignment 垂直对齐方式
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setVerticalAlignment(VerticalAlignment alignment);
    
    /**
     * @brief 设置文本换行
     * @param wrap 是否自动换行
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setTextWrap(bool wrap = true);
    
    /**
     * @brief 设置文本缩进
     * @param indent 缩进级别
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setIndent(int indent);
    
    /**
     * @brief 设置文本旋转角度
     * @param angle 旋转角度（-90到90度）
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setRotation(int angle);
    
    /**
     * @brief 设置边框样式
     * @param style 边框样式
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setBorder(BorderStyle style);
    
    /**
     * @brief 设置左边框样式
     * @param style 边框样式
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setLeftBorder(BorderStyle style);
    
    /**
     * @brief 设置右边框样式
     * @param style 边框样式
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setRightBorder(BorderStyle style);
    
    /**
     * @brief 设置上边框样式
     * @param style 边框样式
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setTopBorder(BorderStyle style);
    
    /**
     * @brief 设置下边框样式
     * @param style 边框样式
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setBottomBorder(BorderStyle style);
    
    /**
     * @brief 设置边框颜色
     * @param color 边框颜色
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setBorderColor(Color color);
    
    /**
     * @brief 设置数字格式
     * @param format 数字格式类型
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setNumberFormat(NumberFormat format);
    
    /**
     * @brief 设置自定义数字格式
     * @param formatString 格式字符串
     * @return Format& 返回自身引用，支持链式调用
     * @example
     * format.setCustomNumberFormat("0.00");      // 两位小数
     * format.setCustomNumberFormat("#,##0.00");  // 千分位分隔符
     * format.setCustomNumberFormat("yyyy-mm-dd"); // 日期格式
     */
    Format& setCustomNumberFormat(const std::string& formatString);
    
    /**
     * @brief 设置单元格锁定状态
     * @param locked 是否锁定
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setLocked(bool locked = true);
    
    /**
     * @brief 设置单元格隐藏状态
     * @param hidden 是否隐藏
     * @return Format& 返回自身引用，支持链式调用
     */
    Format& setHidden(bool hidden = true);
    
    /**
     * @brief 获取底层的libxlsxwriter格式对象
     * @return 格式对象指针（内部使用）
     */
    void* getNativeFormat() const;
    
    /**
     * @brief 获取内部格式对象（用于兼容性）
     * @return 内部格式对象指针
     */
    lxw_format* getInternalFormat() const;
    
    /**
     * @brief 克隆格式对象
     * @param workbook 目标工作簿
     * @return 克隆的格式对象
     */
    [[nodiscard]] std::unique_ptr<Format> clone(lxw_workbook* workbook) const;
};

/**
 * @brief 格式构建器
 * 提供便捷的方法来创建常用格式
 */
class FormatBuilder {
private:
    lxw_workbook* workbook_;
    
public:
    /**
     * @brief 构造函数
     * @param workbook 关联的工作簿
     */
    explicit FormatBuilder(lxw_workbook* workbook) : workbook_(workbook) {}
    
    /**
     * @brief 创建基本格式
     * @return 基本格式对象
     */
    [[nodiscard]] std::unique_ptr<Format> createBasic() const;
    
    /**
     * @brief 创建标题格式
     * @param fontSize 字体大小
     * @param bold 是否粗体
     * @return 标题格式对象
     */
    [[nodiscard]] std::unique_ptr<Format> createTitle(double fontSize = 16, bool bold = true) const;
    
    /**
     * @brief 创建表头格式
     * @param backgroundColor 背景颜色
     * @param fontColor 字体颜色
     * @return 表头格式对象
     */
    [[nodiscard]] std::unique_ptr<Format> createHeader(
        Color backgroundColor = Colors::Gray25,
        Color fontColor = Colors::Black) const;
    
    /**
     * @brief 创建数字格式
     * @param decimalPlaces 小数位数
     * @return 数字格式对象
     */
    [[nodiscard]] std::unique_ptr<Format> createNumber(int decimalPlaces = 2) const;
    
    /**
     * @brief 创建百分比格式
     * @param decimalPlaces 小数位数
     * @return 百分比格式对象
     */
    [[nodiscard]] std::unique_ptr<Format> createPercentage(int decimalPlaces = 2) const;
    
    /**
     * @brief 创建日期格式
     * @return 日期格式对象
     */
    [[nodiscard]] std::unique_ptr<Format> createDate() const;
    
    /**
     * @brief 创建货币格式
     * @param currency 货币符号
     * @return 货币格式对象
     */
    [[nodiscard]] std::unique_ptr<Format> createCurrency(const std::string& currency = "¥") const;
    
    /**
     * @brief 创建带边框的格式
     * @param borderStyle 边框样式
     * @param borderColor 边框颜色
     * @return 带边框的格式对象
     */
    [[nodiscard]] std::unique_ptr<Format> createBorder(
        BorderStyle borderStyle = BorderStyle::Thin,
        Color borderColor = Colors::Black) const;
};

} // namespace TinaXlsx 