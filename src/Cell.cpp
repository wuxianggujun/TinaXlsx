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
    
    inline bool fastStringToInteger(const std::string& str, Integer& result) noexcept {
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
    // 使用高性能的cellValueToString全局函数
    return cellValueToString(value_);
}

std::optional<double> Cell::toNumber() const {
    const CellValueType type = getCellValueType(value_);
    
    switch (type) {
        case CellValueType::String: {
            const auto& str = std::get<std::string>(value_);
            double result;
            return fastStringToDouble(str, result) ? std::optional<double>(result) : std::nullopt;
        }
        case CellValueType::Double:
            return std::get<double>(value_);
        case CellValueType::Integer:
            return static_cast<double>(std::get<Integer>(value_));
        case CellValueType::Boolean:
            return std::get<bool>(value_) ? 1.0 : 0.0;
        case CellValueType::Empty:
            return std::nullopt;
        default:
            return std::nullopt;
    }
}

std::optional<Integer> Cell::toInteger() const {
    const CellValueType type = getCellValueType(value_);
    
    switch (type) {
        case CellValueType::String: {
            const auto& str = std::get<std::string>(value_);
            Integer result;
            return fastStringToInteger(str, result) ? std::optional<Integer>(result) : std::nullopt;
        }
        case CellValueType::Double: {
            const double val = std::get<double>(value_);
            // 优化：更精确的整数检查
            if (val >= static_cast<double>(std::numeric_limits<Integer>::min()) &&
                val <= static_cast<double>(std::numeric_limits<Integer>::max()) &&
                val == std::floor(val)) {
                return static_cast<Integer>(val);
            }
            return std::nullopt;
        }
        case CellValueType::Integer:
            return std::get<Integer>(value_);
        case CellValueType::Boolean:
            return std::get<bool>(value_) ? 1 : 0;
        case CellValueType::Empty:
            return std::nullopt;
        default:
            return std::nullopt;
    }
}

std::optional<bool> Cell::toBoolean() const {
    const CellValueType type = getCellValueType(value_);
    
    switch (type) {
        case CellValueType::String: {
            const auto& str = std::get<std::string>(value_);
            if (str.empty()) return std::nullopt;
            
            // 使用优化的字符串比较
            if (isTrueString(str)) return true;
            if (isFalseString(str)) return false;
            return std::nullopt;
        }
        case CellValueType::Double:
            return std::get<double>(value_) != 0.0;
        case CellValueType::Integer:
            return std::get<Integer>(value_) != 0;
        case CellValueType::Boolean:
            return std::get<bool>(value_);
        case CellValueType::Empty:
            return std::nullopt;
        default:
            return std::nullopt;
    }
}

} // namespace TinaXlsx 