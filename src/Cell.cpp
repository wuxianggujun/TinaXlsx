/**
 * @file Cell.cpp
 * @brief Excel单元格类实现
 */

#include "TinaXlsx/Cell.hpp"
#include <sstream>
#include <charconv>
#include <algorithm>
#include <cctype>

namespace TinaXlsx {

std::string Cell::toString() const {
    return std::visit([](const auto& v) -> std::string {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            return "";
        } else if constexpr (std::is_same_v<T, std::string>) {
            return v;
        } else if constexpr (std::is_same_v<T, double>) {
            return std::to_string(v);
        } else if constexpr (std::is_same_v<T, int64_t>) {
            return std::to_string(v);
        } else if constexpr (std::is_same_v<T, bool>) {
            return v ? "true" : "false";
        } else {
            return "";
        }
    }, value_);
}

std::optional<double> Cell::toNumber() const {
    return std::visit([](const auto& v) -> std::optional<double> {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            return std::nullopt;
        } else if constexpr (std::is_same_v<T, std::string>) {
            if (v.empty()) {
                return std::nullopt;
            }
            
            // 尝试转换字符串为数字
            double result;
            auto convResult = std::from_chars(v.data(), v.data() + v.size(), result);
            if (convResult.ec == std::errc{}) {
                return result;
            }
            
            // 如果from_chars失败，尝试使用stod
            try {
                return std::stod(v);
            } catch (const std::exception&) {
                return std::nullopt;
            }
        } else if constexpr (std::is_same_v<T, double>) {
            return v;
        } else if constexpr (std::is_same_v<T, int64_t>) {
            return static_cast<double>(v);
        } else if constexpr (std::is_same_v<T, bool>) {
            return v ? 1.0 : 0.0;
        } else {
            return std::nullopt;
        }
    }, value_);
}

std::optional<int64_t> Cell::toInteger() const {
    return std::visit([](const auto& v) -> std::optional<int64_t> {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            return std::nullopt;
        } else if constexpr (std::is_same_v<T, std::string>) {
            if (v.empty()) {
                return std::nullopt;
            }
            
            // 尝试转换字符串为整数
            int64_t result;
            auto convResult = std::from_chars(v.data(), v.data() + v.size(), result);
            if (convResult.ec == std::errc{}) {
                return result;
            }
            
            // 如果from_chars失败，尝试使用stoll
            try {
                return std::stoll(v);
            } catch (const std::exception&) {
                return std::nullopt;
            }
        } else if constexpr (std::is_same_v<T, double>) {
            // 检查是否为整数值
            if (v == static_cast<int64_t>(v)) {
                return static_cast<int64_t>(v);
            }
            return std::nullopt;
        } else if constexpr (std::is_same_v<T, int64_t>) {
            return v;
        } else if constexpr (std::is_same_v<T, bool>) {
            return v ? 1 : 0;
        } else {
            return std::nullopt;
        }
    }, value_);
}

std::optional<bool> Cell::toBoolean() const {
    return std::visit([](const auto& v) -> std::optional<bool> {
        using T = std::decay_t<decltype(v)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            return std::nullopt;
        } else if constexpr (std::is_same_v<T, std::string>) {
            if (v.empty()) {
                return std::nullopt;
            }
            
            // 转换为小写进行比较
            std::string lower = v;
            std::transform(lower.begin(), lower.end(), lower.begin(), 
                         [](unsigned char c) { return std::tolower(c); });
            
            if (lower == "true" || lower == "yes" || lower == "1" || lower == "on") {
                return true;
            } else if (lower == "false" || lower == "no" || lower == "0" || lower == "off") {
                return false;
            }
            
            return std::nullopt;
        } else if constexpr (std::is_same_v<T, double>) {
            return v != 0.0;
        } else if constexpr (std::is_same_v<T, int64_t>) {
            return v != 0;
        } else if constexpr (std::is_same_v<T, bool>) {
            return v;
        } else {
            return std::nullopt;
        }
    }, value_);
}

} // namespace TinaXlsx 