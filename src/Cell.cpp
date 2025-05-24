/**
 * @file Cell.cpp
 * @brief Excel单元格类实现 - 性能优化版本
 */

#include "TinaXlsx/Cell.hpp"
#include <sstream>
#include <charconv>
#include <algorithm>
#include <cctype>
#include <cstring>

namespace TinaXlsx {

// 内联工具函数，避免重复的std::visit调用
namespace {
    // 快速类型判断，避免std::visit
    inline constexpr uint8_t getValueType(const CellValue& value) noexcept {
        return value.index();
    }
    
    // 正确的类型索引常量 - 与Types.hpp中CellValue的定义顺序匹配
    constexpr uint8_t TYPE_STRING = 0;      // std::string
    constexpr uint8_t TYPE_DOUBLE = 1;      // double
    constexpr uint8_t TYPE_INT64 = 2;       // int64_t
    constexpr uint8_t TYPE_BOOL = 3;        // bool
    constexpr uint8_t TYPE_MONOSTATE = 4;   // std::monostate
    
    // 优化的字符串转数字函数
    inline bool fastStringToDouble(const std::string& str, double& result) noexcept {
        if (str.empty()) return false;
        
        // 快速路径：使用std::from_chars（最快）
        auto [ptr, ec] = std::from_chars(str.data(), str.data() + str.size(), result);
        if (ec == std::errc{} && ptr == str.data() + str.size()) {
            return true;
        }
        
        // 备用路径：处理科学计数法等复杂格式
        try {
            size_t pos;
            result = std::stod(str, &pos);
            return pos == str.size(); // 确保整个字符串都被解析
        } catch (...) {
            return false;
        }
    }
    
    inline bool fastStringToInt64(const std::string& str, int64_t& result) noexcept {
        if (str.empty()) return false;
        
        // 快速路径：使用std::from_chars
        auto [ptr, ec] = std::from_chars(str.data(), str.data() + str.size(), result);
        if (ec == std::errc{} && ptr == str.data() + str.size()) {
            return true;
        }
        
        // 备用路径
        try {
            size_t pos;
            result = std::stoll(str, &pos);
            return pos == str.size();
        } catch (...) {
            return false;
        }
    }
    
    // 预编译的字符串比较，用于布尔值转换
    inline bool isTrueString(const std::string& str) noexcept {
        // 快速长度检查
        const size_t len = str.length();
        if (len == 0 || len > 5) return false;
        
        // 使用编译时字符串匹配
        switch (len) {
            case 1: return str[0] == '1';
            case 2: return (str[0] == 'o' || str[0] == 'O') && 
                          (str[1] == 'n' || str[1] == 'N');
            case 3: return (str[0] == 'y' || str[0] == 'Y') &&
                          (str[1] == 'e' || str[1] == 'E') &&
                          (str[2] == 's' || str[2] == 'S');
            case 4: return (str[0] == 't' || str[0] == 'T') &&
                          (str[1] == 'r' || str[1] == 'R') &&
                          (str[2] == 'u' || str[2] == 'U') &&
                          (str[3] == 'e' || str[3] == 'E');
            default: return false;
        }
    }
    
    inline bool isFalseString(const std::string& str) noexcept {
        const size_t len = str.length();
        if (len == 0 || len > 5) return false;
        
        switch (len) {
            case 1: return str[0] == '0';
            case 2: return (str[0] == 'n' || str[0] == 'N') && 
                          (str[1] == 'o' || str[1] == 'O');
            case 3: return (str[0] == 'o' || str[0] == 'O') &&
                          (str[1] == 'f' || str[1] == 'F') &&
                          (str[2] == 'f' || str[2] == 'F');
            case 5: return (str[0] == 'f' || str[0] == 'F') &&
                          (str[1] == 'a' || str[1] == 'A') &&
                          (str[2] == 'l' || str[2] == 'L') &&
                          (str[3] == 's' || str[3] == 'S') &&
                          (str[4] == 'e' || str[4] == 'E');
            default: return false;
        }
    }
}

std::string Cell::toString() const {
    // 使用索引而不是std::visit以获得更好的性能
    const uint8_t type = getValueType(value_);
    
    switch (type) {
        case TYPE_STRING:
            return std::get<std::string>(value_);
        case TYPE_DOUBLE: {
            const double val = std::get<double>(value_);
            // 优化：检查是否为整数以避免不必要的小数点
            if (val == static_cast<int64_t>(val) && val >= -9007199254740992.0 && val <= 9007199254740992.0) {
                return std::to_string(static_cast<int64_t>(val));
            }
            return std::to_string(val);
        }
        case TYPE_INT64:
            return std::to_string(std::get<int64_t>(value_));
        case TYPE_BOOL:
            return std::get<bool>(value_) ? "true" : "false";
        case TYPE_MONOSTATE:
            return "";
        default:
            return "";
    }
}

std::optional<double> Cell::toNumber() const {
    const uint8_t type = getValueType(value_);
    
    switch (type) {
        case TYPE_STRING: {
            const auto& str = std::get<std::string>(value_);
            double result;
            return fastStringToDouble(str, result) ? std::optional<double>(result) : std::nullopt;
        }
        case TYPE_DOUBLE:
            return std::get<double>(value_);
        case TYPE_INT64:
            return static_cast<double>(std::get<int64_t>(value_));
        case TYPE_BOOL:
            return std::get<bool>(value_) ? 1.0 : 0.0;
        case TYPE_MONOSTATE:
            return std::nullopt;
        default:
            return std::nullopt;
    }
}

std::optional<int64_t> Cell::toInteger() const {
    const uint8_t type = getValueType(value_);
    
    switch (type) {
        case TYPE_STRING: {
            const auto& str = std::get<std::string>(value_);
            int64_t result;
            return fastStringToInt64(str, result) ? std::optional<int64_t>(result) : std::nullopt;
        }
        case TYPE_DOUBLE: {
            const double val = std::get<double>(value_);
            // 优化：更精确的整数检查
            if (val >= static_cast<double>(std::numeric_limits<int64_t>::min()) &&
                val <= static_cast<double>(std::numeric_limits<int64_t>::max()) &&
                val == std::floor(val)) {
                return static_cast<int64_t>(val);
            }
            return std::nullopt;
        }
        case TYPE_INT64:
            return std::get<int64_t>(value_);
        case TYPE_BOOL:
            return std::get<bool>(value_) ? 1 : 0;
        case TYPE_MONOSTATE:
            return std::nullopt;
        default:
            return std::nullopt;
    }
}

std::optional<bool> Cell::toBoolean() const {
    const uint8_t type = getValueType(value_);
    
    switch (type) {
        case TYPE_STRING: {
            const auto& str = std::get<std::string>(value_);
            if (str.empty()) return std::nullopt;
            
            // 使用优化的字符串比较
            if (isTrueString(str)) return true;
            if (isFalseString(str)) return false;
            return std::nullopt;
        }
        case TYPE_DOUBLE:
            return std::get<double>(value_) != 0.0;
        case TYPE_INT64:
            return std::get<int64_t>(value_) != 0;
        case TYPE_BOOL:
            return std::get<bool>(value_);
        case TYPE_MONOSTATE:
            return std::nullopt;
        default:
            return std::nullopt;
    }
}

} // namespace TinaXlsx 