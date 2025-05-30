#pragma once

#include "TXTypes.hpp"
#include <string>
#include <vector>
#include <memory>
#include <unordered_map>
#include <variant>
#include <regex>

namespace TinaXlsx {

/**
 * @brief 数字格式化类
 * 
 * 提供Excel兼容的数字、日期、百分比等格式化功能
 */
class TXNumberFormat {
public:
    /**
     * @brief 预定义格式类型
     */
    enum class FormatType {
        General = 0,        ///< 常规
        Number = 1,         ///< 数字
        Decimal = 2,        ///< 小数
        Currency = 3,       ///< 货币
        Accounting = 4,     ///< 会计
        Date = 14,          ///< 日期
        Time = 18,          ///< 时间
        DateTime = 22,      ///< 日期时间
        Percentage = 9,     ///< 百分比
        Fraction = 12,      ///< 分数
        Scientific = 11,    ///< 科学计数法
        Text = 49,          ///< 文本
        Custom = -1         ///< 自定义
    };

    /**
     * @brief 格式选项
     */
    struct FormatOptions {
        int decimalPlaces;          ///< 小数位数
        bool useThousandSeparator; ///< 使用千位分隔符
        std::string currencySymbol; ///< 货币符号
        std::string dateFormat; ///< 日期格式
        std::string timeFormat;   ///< 时间格式
        bool showNegativeInRed;  ///< 负数显示为红色
        bool showZero;            ///< 显示零值
        
        FormatOptions() 
            : decimalPlaces(2)
            , useThousandSeparator(true)
            , currencySymbol("$")
            , dateFormat("yyyy-mm-dd")
            , timeFormat("hh:mm:ss")
            , showNegativeInRed(false)
            , showZero(true) {}
    };

    /**
     * @brief 值类型
     */
    using Value = cell_value_t;

public:
    TXNumberFormat();
    explicit TXNumberFormat(FormatType type, const FormatOptions& options = FormatOptions{});
    explicit TXNumberFormat(const std::string& customFormat);
    ~TXNumberFormat() = default;
    
    // 支持拷贝和移动
    TXNumberFormat(const TXNumberFormat& other) = default;
    TXNumberFormat& operator=(const TXNumberFormat& other) = default;
    TXNumberFormat(TXNumberFormat&& other) noexcept = default;
    TXNumberFormat& operator=(TXNumberFormat&& other) noexcept = default;

    // ==================== 格式设置 ====================

    /**
     * @brief 设置格式类型
     * @param type 格式类型
     * @param options 格式选项
     */
    void setFormat(FormatType type, const FormatOptions& options = FormatOptions{});

    /**
     * @brief 设置自定义格式
     * @param formatString 格式字符串
     */
    void setCustomFormat(const std::string& formatString);

    /**
     * @brief 获取格式类型
     * @return 格式类型
     */
    FormatType getFormatType() const;

    /**
     * @brief 获取格式字符串
     * @return 格式字符串
     */
    std::string getFormatString() const;

    /**
     * @brief 获取格式选项
     * @return 格式选项
     */
    const FormatOptions& getFormatOptions() const;

    // ==================== 格式化方法 ====================

    /**
     * @brief 格式化值
     * @param value 要格式化的值
     * @return 格式化后的字符串
     */
    std::string format(const Value& value) const;

    /**
     * @brief 格式化数字
     * @param number 数字
     * @return 格式化后的字符串
     */
    std::string formatNumber(double number) const;

    /**
     * @brief 格式化整数
     * @param integer 整数
     * @return 格式化后的字符串
     */
    std::string formatInteger(int64_t integer) const;

    /**
     * @brief 格式化百分比
     * @param value 值（0.5 = 50%）
     * @return 格式化后的字符串
     */
    std::string formatPercentage(double value) const;

    /**
     * @brief 格式化货币
     * @param amount 金额
     * @return 格式化后的字符串
     */
    std::string formatCurrency(double amount) const;

    /**
     * @brief 格式化日期
     * @param excelDate Excel日期序列号
     * @return 格式化后的字符串
     */
    std::string formatDate(double excelDate) const;

    /**
     * @brief 格式化时间
     * @param excelTime Excel时间序列号
     * @return 格式化后的字符串
     */
    std::string formatTime(double excelTime) const;

    /**
     * @brief 格式化科学计数法
     * @param number 数字
     * @return 格式化后的字符串
     */
    std::string formatScientific(double number) const;

    // ==================== 解析方法 ====================

    /**
     * @brief 解析格式化的字符串回数值
     * @param formattedStr 格式化的字符串
     * @return 解析后的值
     */
    Value parse(const std::string& formattedStr) const;

    /**
     * @brief 检查字符串是否匹配当前格式
     * @param str 字符串
     * @return 匹配返回true，否则返回false
     */
    bool matches(const std::string& str) const;

    // ==================== 预定义格式 ====================

    /**
     * @brief 创建数字格式
     * @param decimalPlaces 小数位数
     * @param useThousandSeparator 是否使用千位分隔符
     * @return 数字格式对象
     */
    static TXNumberFormat createNumberFormat(int decimalPlaces = 2, bool useThousandSeparator = true);

    /**
     * @brief 创建货币格式
     * @param currencySymbol 货币符号
     * @param decimalPlaces 小数位数
     * @return 货币格式对象
     */
    static TXNumberFormat createCurrencyFormat(const std::string& currencySymbol = "$", int decimalPlaces = 2);

    /**
     * @brief 创建百分比格式
     * @param decimalPlaces 小数位数
     * @return 百分比格式对象
     */
    static TXNumberFormat createPercentageFormat(int decimalPlaces = 2);

    /**
     * @brief 创建日期格式
     * @param dateFormat 日期格式字符串
     * @return 日期格式对象
     */
    static TXNumberFormat createDateFormat(const std::string& dateFormat = "yyyy-mm-dd");

    /**
     * @brief 创建时间格式
     * @param timeFormat 时间格式字符串
     * @return 时间格式对象
     */
    static TXNumberFormat createTimeFormat(const std::string& timeFormat = "hh:mm:ss");

    /**
     * @brief 创建科学计数法格式
     * @param decimalPlaces 小数位数
     * @return 科学计数法格式对象
     */
    static TXNumberFormat createScientificFormat(int decimalPlaces = 2);

    // ==================== 日期时间工具 ====================

    /**
     * @brief 将系统时间转换为Excel日期序列号
     * @param timeT 系统时间
     * @return Excel日期序列号
     */
    static double systemTimeToExcelDate(time_t timeT);

    /**
     * @brief 将Excel日期序列号转换为系统时间
     * @param excelDate Excel日期序列号
     * @return 系统时间
     */
    static time_t excelDateToSystemTime(double excelDate);

    /**
     * @brief 获取当前时间的Excel日期序列号
     * @return Excel日期序列号
     */
    static double getCurrentExcelDate();

    /**
     * @brief 解析日期字符串为Excel日期序列号
     * @param dateStr 日期字符串
     * @param format 日期格式
     * @return Excel日期序列号
     */
    static double parseDateString(const std::string& dateStr, const std::string& format = "yyyy-mm-dd");

    // ==================== 格式代码生成 ====================

    /**
     * @brief 获取Excel格式代码
     * @return Excel格式代码
     */
    int getExcelFormatCode() const;

    /**
     * @brief 生成Excel格式字符串
     * @return Excel格式字符串
     */
    std::string generateExcelFormatString() const;

    // ==================== 验证和工具 ====================

    /**
     * @brief 验证格式字符串是否有效
     * @param formatString 格式字符串
     * @return 有效返回true，否则返回false
     */
    static bool isValidFormatString(const std::string& formatString);

    /**
     * @brief 获取所有预定义格式
     * @return 格式类型到描述的映射
     */
    static std::unordered_map<FormatType, std::string> getPredefinedFormats();

    /**
     * @brief 获取格式类型的描述
     * @param type 格式类型
     * @return 描述字符串
     */
    static std::string getFormatDescription(FormatType type);

    /**
     * @brief 检查值是否为数字类型
     * @param value 值
     * @return 是数字返回true，否则返回false
     */
    static bool isNumericValue(const Value& value);

    /**
     * @brief 将值转换为数字
     * @param value 值
     * @return 数字值
     */
    static double valueToNumber(const Value& value);

private:
    FormatType formatType_ = FormatType::General;
    FormatOptions options_;
    std::string customFormatString_;
    
    // 预编译的格式模式
    std::regex numberPattern_;
    std::regex datePattern_;
    std::regex timePattern_;

    // ==================== 私有辅助方法 ====================

    /**
     * @brief 更新格式模式
     */
    void updatePatterns();

    /**
     * @brief 格式化值的内部实现
     */
    std::string formatValue(const Value& value) const;

    /**
     * @brief 格式化常规类型
     */
    std::string formatGeneral(const Value& value) const;

    /**
     * @brief 格式化日期时间
     */
    std::string formatDateTime(double excelDateTime) const;

    /**
     * @brief 格式化文本
     */
    std::string formatText(const Value& value) const;

    /**
     * @brief 格式化自定义格式
     */
    std::string formatCustom(const Value& value) const;

    /**
     * @brief 解析值的内部实现
     */
    Value parseValue(const std::string& formattedStr) const;

    /**
     * @brief 解析常规格式
     */
    Value parseGeneral(const std::string& str) const;

    /**
     * @brief 解析数字
     */
    Value parseNumber(const std::string& str) const;

    /**
     * @brief 解析货币
     */
    Value parseCurrency(const std::string& str) const;

    /**
     * @brief 解析百分比
     */
    Value parsePercentage(const std::string& str) const;

    /**
     * @brief 解析日期
     */
    Value parseDate(const std::string& str) const;

    /**
     * @brief 解析时间
     */
    Value parseTime(const std::string& str) const;

    /**
     * @brief 将Excel日期转换为系统时间（内部使用）
     */
    static time_t excelDateToSystemTimeInternal(double excelDate);

    /**
     * @brief 将系统时间转换为Excel日期（内部使用）
     */
    static double systemTimeToExcelDateInternal(time_t timeT);

    /**
     * @brief 解析日期字符串（内部使用）
     */
    static double parseDateStringInternal(const std::string& dateStr, const std::string& format);
};

} // namespace TinaXlsx 