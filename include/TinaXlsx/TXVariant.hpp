//
// @file TXVariant.hpp
// @brief 通用数据类型变体 - 支持Excel所有数据类型
//

#pragma once

#include <string>
#include <variant>
#include <type_traits>
#include <stdexcept>

namespace TinaXlsx {

/**
 * @brief Excel单元格数据类型枚举
 */
enum class TXCellType : uint8_t {
    Empty = 0,
    Number = 1,
    String = 2,
    Boolean = 3,
    Formula = 4,
    Error = 5
};

/**
 * @brief 通用数据变体类 - 支持Excel所有数据类型
 */
class TXVariant {
public:
    enum class Type {
        Empty,
        Number,
        String,
        Boolean
    };
    
    // 默认构造函数 - 空值
    TXVariant() : type_(Type::Empty) {}
    
    // 数值构造
    TXVariant(double value) : type_(Type::Number), number_value_(value) {}
    TXVariant(int value) : type_(Type::Number), number_value_(static_cast<double>(value)) {}
    TXVariant(float value) : type_(Type::Number), number_value_(static_cast<double>(value)) {}
    
    // 字符串构造
    TXVariant(const std::string& value) : type_(Type::String), string_value_(value) {}
    TXVariant(const char* value) : type_(Type::String), string_value_(value) {}
    TXVariant(std::string&& value) : type_(Type::String), string_value_(std::move(value)) {}
    
    // 布尔构造
    TXVariant(bool value) : type_(Type::Boolean), boolean_value_(value) {}
    
    // 拷贝构造
    TXVariant(const TXVariant& other) : type_(other.type_) {
        switch (type_) {
            case Type::Number: number_value_ = other.number_value_; break;
            case Type::String: string_value_ = other.string_value_; break;
            case Type::Boolean: boolean_value_ = other.boolean_value_; break;
            case Type::Empty: break;
        }
    }
    
    // 移动构造
    TXVariant(TXVariant&& other) noexcept : type_(other.type_) {
        switch (type_) {
            case Type::Number: number_value_ = other.number_value_; break;
            case Type::String: string_value_ = std::move(other.string_value_); break;
            case Type::Boolean: boolean_value_ = other.boolean_value_; break;
            case Type::Empty: break;
        }
        other.type_ = Type::Empty;
    }
    
    // 赋值操作符
    TXVariant& operator=(const TXVariant& other) {
        if (this != &other) {
            type_ = other.type_;
            switch (type_) {
                case Type::Number: number_value_ = other.number_value_; break;
                case Type::String: string_value_ = other.string_value_; break;
                case Type::Boolean: boolean_value_ = other.boolean_value_; break;
                case Type::Empty: break;
            }
        }
        return *this;
    }
    
    TXVariant& operator=(TXVariant&& other) noexcept {
        if (this != &other) {
            type_ = other.type_;
            switch (type_) {
                case Type::Number: number_value_ = other.number_value_; break;
                case Type::String: string_value_ = std::move(other.string_value_); break;
                case Type::Boolean: boolean_value_ = other.boolean_value_; break;
                case Type::Empty: break;
            }
            other.type_ = Type::Empty;
        }
        return *this;
    }
    
    // 类型检查
    Type getType() const { return type_; }
    bool isEmpty() const { return type_ == Type::Empty; }
    bool isNumber() const { return type_ == Type::Number; }
    bool isString() const { return type_ == Type::String; }
    bool isBoolean() const { return type_ == Type::Boolean; }
    
    // 值获取
    double getNumber() const {
        if (type_ != Type::Number) {
            throw std::runtime_error("TXVariant is not a number");
        }
        return number_value_;
    }
    
    const std::string& getString() const {
        if (type_ != Type::String) {
            throw std::runtime_error("TXVariant is not a string");
        }
        return string_value_;
    }
    
    bool getBoolean() const {
        if (type_ != Type::Boolean) {
            throw std::runtime_error("TXVariant is not a boolean");
        }
        return boolean_value_;
    }
    
    // 安全的值获取（带默认值）
    double getNumberOr(double default_value) const {
        return (type_ == Type::Number) ? number_value_ : default_value;
    }
    
    std::string getStringOr(const std::string& default_value) const {
        return (type_ == Type::String) ? string_value_ : default_value;
    }
    
    bool getBooleanOr(bool default_value) const {
        return (type_ == Type::Boolean) ? boolean_value_ : default_value;
    }
    
    // 转换为字符串表示
    std::string toString() const {
        switch (type_) {
            case Type::Empty: return "";
            case Type::Number: return std::to_string(number_value_);
            case Type::String: return string_value_;
            case Type::Boolean: return boolean_value_ ? "TRUE" : "FALSE";
        }
        return "";
    }
    
    // 比较操作符
    bool operator==(const TXVariant& other) const {
        if (type_ != other.type_) return false;
        
        switch (type_) {
            case Type::Empty: return true;
            case Type::Number: return number_value_ == other.number_value_;
            case Type::String: return string_value_ == other.string_value_;
            case Type::Boolean: return boolean_value_ == other.boolean_value_;
        }
        return false;
    }
    
    bool operator!=(const TXVariant& other) const {
        return !(*this == other);
    }

private:
    Type type_ = Type::Empty;
    
    union {
        double number_value_;
        bool boolean_value_;
    };
    std::string string_value_;
};

} // namespace TinaXlsx 
