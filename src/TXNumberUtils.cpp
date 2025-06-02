//
// @file TXNumberUtils.cpp
// @brief 高性能数值解析和格式化工具类实现
// @author wuxianggujun
//

#include "TinaXlsx/TXNumberUtils.hpp"
#include <algorithm>
#include <cmath>
#include <limits>

namespace TinaXlsx {

// ==================== 解析方法 ====================

TXNumberUtils::ParseResult TXNumberUtils::parseDouble(std::string_view str, double& result) {
    if (str.empty()) {
        return ParseResult::Empty;
    }

    // 去除前后空白字符
    str = trimWhitespace(str);
    if (str.empty()) {
        return ParseResult::Empty;
    }

    // 使用 fast_float 进行高性能解析
    auto parseResult = fast_float::from_chars(str.data(), str.data() + str.size(), result);
    
    if (parseResult.ec == std::errc{}) {
        // 检查是否解析了整个字符串
        if (parseResult.ptr == str.data() + str.size()) {
            // 检查范围
            if (std::isfinite(result)) {
                return ParseResult::Success;
            } else {
                return ParseResult::OutOfRange;
            }
        } else {
            return ParseResult::InvalidFormat;
        }
    } else if (parseResult.ec == std::errc::result_out_of_range) {
        return ParseResult::OutOfRange;
    } else {
        return ParseResult::InvalidFormat;
    }
}

std::optional<double> TXNumberUtils::parseDouble(std::string_view str) {
    double result;
    if (parseDouble(str, result) == ParseResult::Success) {
        return result;
    }
    return std::nullopt;
}

TXNumberUtils::ParseResult TXNumberUtils::parseFloat(std::string_view str, float& result) {
    if (str.empty()) {
        return ParseResult::Empty;
    }

    str = trimWhitespace(str);
    if (str.empty()) {
        return ParseResult::Empty;
    }

    auto parseResult = fast_float::from_chars(str.data(), str.data() + str.size(), result);
    
    if (parseResult.ec == std::errc{}) {
        if (parseResult.ptr == str.data() + str.size()) {
            if (std::isfinite(result)) {
                return ParseResult::Success;
            } else {
                return ParseResult::OutOfRange;
            }
        } else {
            return ParseResult::InvalidFormat;
        }
    } else if (parseResult.ec == std::errc::result_out_of_range) {
        return ParseResult::OutOfRange;
    } else {
        return ParseResult::InvalidFormat;
    }
}

std::optional<float> TXNumberUtils::parseFloat(std::string_view str) {
    float result;
    if (parseFloat(str, result) == ParseResult::Success) {
        return result;
    }
    return std::nullopt;
}

TXNumberUtils::ParseResult TXNumberUtils::parseInt64(std::string_view str, int64_t& result) {
    if (str.empty()) {
        return ParseResult::Empty;
    }

    str = trimWhitespace(str);
    if (str.empty()) {
        return ParseResult::Empty;
    }

    // 对于整数，我们仍然可以使用 fast_float，然后检查是否为整数
    double doubleResult;
    auto parseResult = fast_float::from_chars(str.data(), str.data() + str.size(), doubleResult);
    
    if (parseResult.ec == std::errc{}) {
        if (parseResult.ptr == str.data() + str.size()) {
            // 检查是否为整数
            if (isInteger(doubleResult)) {
                // 检查范围
                if (doubleResult >= std::numeric_limits<int64_t>::min() && 
                    doubleResult <= std::numeric_limits<int64_t>::max()) {
                    result = static_cast<int64_t>(doubleResult);
                    return ParseResult::Success;
                } else {
                    return ParseResult::OutOfRange;
                }
            } else {
                return ParseResult::InvalidFormat;
            }
        } else {
            return ParseResult::InvalidFormat;
        }
    } else {
        return ParseResult::InvalidFormat;
    }
}

std::optional<int64_t> TXNumberUtils::parseInt64(std::string_view str) {
    int64_t result;
    if (parseInt64(str, result) == ParseResult::Success) {
        return result;
    }
    return std::nullopt;
}

// ==================== 格式化方法 ====================

std::string TXNumberUtils::formatDouble(double value, const FormatOptions& options) {
    return formatDoubleInternal(value, options);
}

std::string TXNumberUtils::formatFloat(float value, const FormatOptions& options) {
    return formatDoubleInternal(static_cast<double>(value), options);
}

std::string TXNumberUtils::formatInt64(int64_t value, const FormatOptions& options) {
    if (options.useThousandSeparator) {
        // 使用千位分隔符
        std::string result = std::to_string(value);
        
        // 找到数字开始位置（跳过负号）
        size_t start = (result[0] == '-') ? 1 : 0;
        
        // 从右到左添加千位分隔符
        for (size_t i = result.length() - 3; i > start; i -= 3) {
            result.insert(i, 1, options.thousandSeparator);
        }
        
        return result;
    } else {
        return std::to_string(value);
    }
}

std::string TXNumberUtils::formatForExcelXml(double value) {
    if (isInteger(value)) {
        // 整数：使用整数格式
        return std::to_string(static_cast<long long>(value));
    } else {
        // 小数：使用最小精度格式
        FormatOptions options;
        options.precision = 2;
        options.removeTrailingZeros = true;
        options.useThousandSeparator = false;
        
        return formatDoubleInternal(value, options);
    }
}

// ==================== 工具方法 ====================

bool TXNumberUtils::isValidNumber(std::string_view str) {
    double dummy;
    return parseDouble(str, dummy) == ParseResult::Success;
}

bool TXNumberUtils::isInteger(double value) {
    return std::isfinite(value) && (value == std::floor(value));
}

std::string TXNumberUtils::removeTrailingZeros(const std::string& str) {
    if (str.find('.') == std::string::npos) {
        return str; // 没有小数点，直接返回
    }
    
    std::string result = str;
    
    // 移除尾随的零
    result.erase(result.find_last_not_of('0') + 1, std::string::npos);
    
    // 如果最后一个字符是小数点，也移除它
    if (!result.empty() && result.back() == '.') {
        result.pop_back();
    }
    
    return result;
}

std::string TXNumberUtils::getParseErrorDescription(ParseResult result) {
    switch (result) {
        case ParseResult::Success:
            return "Success";
        case ParseResult::InvalidFormat:
            return "Invalid number format";
        case ParseResult::OutOfRange:
            return "Number out of range";
        case ParseResult::Empty:
            return "Empty string";
        default:
            return "Unknown error";
    }
}

// ==================== 私有方法 ====================

std::string TXNumberUtils::formatDoubleInternal(double value, const FormatOptions& options) {
    if (!std::isfinite(value)) {
        if (std::isnan(value)) {
            return "NaN";
        } else if (std::isinf(value)) {
            return value > 0 ? "Infinity" : "-Infinity";
        }
    }

    std::ostringstream oss;
    
    if (options.useScientific) {
        oss << std::scientific;
    } else {
        oss << std::fixed;
    }
    
    if (options.precision >= 0) {
        oss << std::setprecision(options.precision);
    } else {
        // 自动精度：对于整数使用0位小数，其他使用2位
        if (isInteger(value)) {
            oss << std::setprecision(0);
        } else {
            oss << std::setprecision(2);
        }
    }
    
    oss << value;
    std::string result = oss.str();
    
    if (options.removeTrailingZeros && !options.useScientific) {
        result = removeTrailingZeros(result);
    }
    
    return result;
}

bool TXNumberUtils::isWhitespace(char c) {
    return c == ' ' || c == '\t' || c == '\n' || c == '\r';
}

std::string_view TXNumberUtils::trimWhitespace(std::string_view str) {
    // 去除前导空白
    while (!str.empty() && isWhitespace(str.front())) {
        str.remove_prefix(1);
    }
    
    // 去除尾随空白
    while (!str.empty() && isWhitespace(str.back())) {
        str.remove_suffix(1);
    }
    
    return str;
}

} // namespace TinaXlsx
