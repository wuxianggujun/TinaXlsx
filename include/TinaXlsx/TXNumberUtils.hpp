//
// @file TXNumberUtils.hpp
// @brief 高性能数值解析和格式化工具类
// @author wuxianggujun
//

#pragma once

#include <string>
#include <string_view>
#include <optional>
#include <sstream>
#include <iomanip>
#include <fast_float/fast_float.h>

namespace TinaXlsx {

/**
 * @brief 高性能数值解析和格式化工具类
 * 
 * 使用 fast_float 库提供高性能的浮点数解析，
 * 并提供统一的数值格式化功能，确保Excel兼容性。
 */
class TXNumberUtils {
public:
    /**
     * @brief 解析结果枚举
     */
    enum class ParseResult {
        Success,        ///< 解析成功
        InvalidFormat,  ///< 格式无效
        OutOfRange,     ///< 超出范围
        Empty          ///< 空字符串
    };

    /**
     * @brief 数值格式化选项
     */
    struct FormatOptions {
        int precision = -1;              ///< 精度（-1表示自动）
        bool removeTrailingZeros = true; ///< 移除尾随零
        bool useScientific = false;      ///< 使用科学计数法
        char decimalPoint = '.';         ///< 小数点字符
        char thousandSeparator = ',';    ///< 千位分隔符
        bool useThousandSeparator = false; ///< 是否使用千位分隔符
    };

    // ==================== 解析方法 ====================

    /**
     * @brief 高性能解析浮点数
     * @param str 要解析的字符串
     * @param result 解析结果（输出参数）
     * @return 解析状态
     */
    static ParseResult parseDouble(std::string_view str, double& result);

    /**
     * @brief 高性能解析浮点数（返回optional）
     * @param str 要解析的字符串
     * @return 解析成功返回数值，失败返回nullopt
     */
    static std::optional<double> parseDouble(std::string_view str);

    /**
     * @brief 高性能解析单精度浮点数
     * @param str 要解析的字符串
     * @param result 解析结果（输出参数）
     * @return 解析状态
     */
    static ParseResult parseFloat(std::string_view str, float& result);

    /**
     * @brief 高性能解析单精度浮点数（返回optional）
     * @param str 要解析的字符串
     * @return 解析成功返回数值，失败返回nullopt
     */
    static std::optional<float> parseFloat(std::string_view str);

    /**
     * @brief 解析整数（64位）
     * @param str 要解析的字符串
     * @param result 解析结果（输出参数）
     * @return 解析状态
     */
    static ParseResult parseInt64(std::string_view str, int64_t& result);

    /**
     * @brief 解析整数（64位，返回optional）
     * @param str 要解析的字符串
     * @return 解析成功返回数值，失败返回nullopt
     */
    static std::optional<int64_t> parseInt64(std::string_view str);

    // ==================== 格式化方法 ====================

    /**
     * @brief 格式化双精度浮点数为Excel兼容格式
     * @param value 要格式化的数值
     * @param options 格式化选项
     * @return 格式化后的字符串
     */
    static std::string formatDouble(double value, const FormatOptions& options = {});

    /**
     * @brief 格式化单精度浮点数为Excel兼容格式
     * @param value 要格式化的数值
     * @param options 格式化选项
     * @return 格式化后的字符串
     */
    static std::string formatFloat(float value, const FormatOptions& options = {});

    /**
     * @brief 格式化整数
     * @param value 要格式化的数值
     * @param options 格式化选项
     * @return 格式化后的字符串
     */
    static std::string formatInt64(int64_t value, const FormatOptions& options = {});

    /**
     * @brief 格式化数值为Excel XML兼容格式
     * 
     * 这个方法确保生成的数值格式与Excel筛选条件兼容：
     * - 整数：不带小数点（如 "3000"）
     * - 小数：最小精度，移除尾随零（如 "123.45"）
     * 
     * @param value 要格式化的数值
     * @return Excel XML兼容的字符串
     */
    static std::string formatForExcelXml(double value);

    // ==================== 工具方法 ====================

    /**
     * @brief 检查字符串是否为有效数字
     * @param str 要检查的字符串
     * @return 如果是有效数字返回true
     */
    static bool isValidNumber(std::string_view str);

    /**
     * @brief 检查数值是否为整数
     * @param value 要检查的数值
     * @return 如果是整数返回true
     */
    static bool isInteger(double value);

    /**
     * @brief 移除字符串中的尾随零
     * @param str 要处理的字符串
     * @return 移除尾随零后的字符串
     */
    static std::string removeTrailingZeros(const std::string& str);

    /**
     * @brief 获取解析错误的描述
     * @param result 解析结果
     * @return 错误描述字符串
     */
    static std::string getParseErrorDescription(ParseResult result);

private:
    // 私有辅助方法
    static std::string formatDoubleInternal(double value, const FormatOptions& options);
    static bool isWhitespace(char c);
    static std::string_view trimWhitespace(std::string_view str);
};

} // namespace TinaXlsx
