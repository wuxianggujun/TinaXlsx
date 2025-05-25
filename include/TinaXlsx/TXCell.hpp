#pragma once

#include <string>
#include <variant>
#include <memory>

namespace TinaXlsx {

/**
 * @brief Excel单元格类
 * 
 * 单元格数据的抽象，提供读写单个单元格的能力
 */
class TXCell {
public:
    /**
     * @brief 单元格值类型
     */
    using CellValue = std::variant<std::monostate, std::string, double, int64_t, bool>;

    /**
     * @brief 单元格类型枚举
     */
    enum class CellType {
        Empty,      ///< 空单元格
        String,     ///< 字符串
        Number,     ///< 数字
        Integer,    ///< 整数
        Boolean,    ///< 布尔值
        Formula,    ///< 公式
        Error       ///< 错误值
    };

    /**
     * @brief 数字格式类型
     */
    enum class NumberFormat {
        General,        ///< 常规
        Number,         ///< 数字
        Currency,       ///< 货币
        Percentage,     ///< 百分比
        Date,           ///< 日期
        Time,           ///< 时间
        DateTime,       ///< 日期时间
        Scientific,     ///< 科学计数法
        Text           ///< 文本
    };

public:
    TXCell();
    explicit TXCell(const CellValue& value);
    ~TXCell();

    // 支持拷贝构造和赋值
    TXCell(const TXCell& other);
    TXCell& operator=(const TXCell& other);

    // 支持移动构造和赋值
    TXCell(TXCell&& other) noexcept;
    TXCell& operator=(TXCell&& other) noexcept;

    /**
     * @brief 获取单元格值
     * @return 单元格值
     */
    const CellValue& getValue() const;

    /**
     * @brief 设置单元格值
     * @param value 单元格值
     */
    void setValue(const CellValue& value);

    /**
     * @brief 获取单元格类型
     * @return 单元格类型
     */
    CellType getType() const;

    /**
     * @brief 检查单元格是否为空
     * @return 为空返回true，否则返回false
     */
    bool isEmpty() const;

    /**
     * @brief 获取字符串值
     * @return 字符串值，如果类型不匹配返回空字符串
     */
    std::string getStringValue() const;

    /**
     * @brief 获取数字值
     * @return 数字值，如果类型不匹配返回0.0
     */
    double getNumberValue() const;

    /**
     * @brief 获取整数值
     * @return 整数值，如果类型不匹配返回0
     */
    int64_t getIntegerValue() const;

    /**
     * @brief 获取布尔值
     * @return 布尔值，如果类型不匹配返回false
     */
    bool getBooleanValue() const;

    /**
     * @brief 设置字符串值
     * @param value 字符串值
     */
    void setStringValue(const std::string& value);

    /**
     * @brief 设置数字值
     * @param value 数字值
     */
    void setNumberValue(double value);

    /**
     * @brief 设置整数值
     * @param value 整数值
     */
    void setIntegerValue(int64_t value);

    /**
     * @brief 设置布尔值
     * @param value 布尔值
     */
    void setBooleanValue(bool value);

    /**
     * @brief 获取公式
     * @return 公式字符串，如果不是公式返回空字符串
     */
    std::string getFormula() const;

    /**
     * @brief 设置公式
     * @param formula 公式字符串（不包含等号）
     */
    void setFormula(const std::string& formula);

    /**
     * @brief 检查是否为公式
     * @return 是公式返回true，否则返回false
     */
    bool isFormula() const;

    /**
     * @brief 获取数字格式
     * @return 数字格式类型
     */
    NumberFormat getNumberFormat() const;

    /**
     * @brief 设置数字格式
     * @param format 数字格式类型
     */
    void setNumberFormat(NumberFormat format);

    /**
     * @brief 获取自定义格式字符串
     * @return 自定义格式字符串
     */
    std::string getCustomFormat() const;

    /**
     * @brief 设置自定义格式
     * @param format_string 格式字符串
     */
    void setCustomFormat(const std::string& format_string);

    /**
     * @brief 清空单元格
     */
    void clear();

    /**
     * @brief 转换为字符串表示
     * @return 字符串表示
     */
    std::string toString() const;

    /**
     * @brief 从字符串解析值
     * @param str 字符串
     * @param auto_detect_type 是否自动检测类型
     * @return 成功返回true，失败返回false
     */
    bool fromString(const std::string& str, bool auto_detect_type = true);

    /**
     * @brief 类型转换操作符
     */
    operator std::string() const { return getStringValue(); }
    operator double() const { return getNumberValue(); }
    operator int64_t() const { return getIntegerValue(); }
    operator bool() const { return getBooleanValue(); }

    /**
     * @brief 赋值操作符
     */
    TXCell& operator=(const std::string& value);
    TXCell& operator=(const char* value);
    TXCell& operator=(double value);
    TXCell& operator=(int64_t value);
    TXCell& operator=(int value);
    TXCell& operator=(bool value);

    /**
     * @brief 比较操作符
     */
    bool operator==(const TXCell& other) const;
    bool operator!=(const TXCell& other) const;

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

} // namespace TinaXlsx 