/**
 * @file Cell.hpp
 * @brief Excel单元格类
 */

#pragma once

#include "Types.hpp"
#include "Exception.hpp"
#include <string>
#include <variant>
#include <optional>

namespace TinaXlsx {

/**
 * @brief Excel单元格类
 * 
 * 提供单元格数据的高级操作接口
 */
class Cell {
private:
    CellValue value_;
    CellPosition position_;
    
public:
    /**
     * @brief 默认构造函数
     */
    Cell() : value_(std::monostate{}), position_() {}
    
    /**
     * @brief 构造函数
     * @param position 单元格位置
     * @param value 单元格值
     */
    Cell(const CellPosition& position, const CellValue& value)
        : value_(value), position_(position) {}
    
    /**
     * @brief 获取单元格位置
     * @return CellPosition 单元格位置
     */
    [[nodiscard]] const CellPosition& getPosition() const { return position_; }
    
    /**
     * @brief 设置单元格位置
     * @param position 单元格位置
     */
    void setPosition(const CellPosition& position) { position_ = position; }
    
    /**
     * @brief 获取单元格值
     * @return CellValue 单元格值
     */
    [[nodiscard]] const CellValue& getValue() const { return value_; }
    
    /**
     * @brief 设置单元格值
     * @param value 单元格值
     */
    void setValue(const CellValue& value) { value_ = value; }
    
    /**
     * @brief 检查单元格是否为空
     * @return bool 是否为空
     */
    [[nodiscard]] bool isEmpty() const {
        return std::holds_alternative<std::monostate>(value_);
    }
    
    /**
     * @brief 检查单元格是否包含字符串
     * @return bool 是否包含字符串
     */
    [[nodiscard]] bool isString() const {
        return std::holds_alternative<std::string>(value_);
    }
    
    /**
     * @brief 检查单元格是否包含数字
     * @return bool 是否包含数字
     */
    [[nodiscard]] bool isNumber() const {
        return std::holds_alternative<double>(value_);
    }
    
    /**
     * @brief 检查单元格是否包含整数
     * @return bool 是否包含整数
     */
    [[nodiscard]] bool isInteger() const {
        return std::holds_alternative<int64_t>(value_);
    }
    
    /**
     * @brief 检查单元格是否包含布尔值
     * @return bool 是否包含布尔值
     */
    [[nodiscard]] bool isBoolean() const {
        return std::holds_alternative<bool>(value_);
    }
    
    /**
     * @brief 获取字符串值
     * @return std::optional<std::string> 字符串值，如果类型不匹配则返回空
     */
    [[nodiscard]] std::optional<std::string> getString() const {
        if (isString()) {
            return std::get<std::string>(value_);
        }
        return std::nullopt;
    }
    
    /**
     * @brief 获取数字值
     * @return std::optional<double> 数字值，如果类型不匹配则返回空
     */
    [[nodiscard]] std::optional<double> getNumber() const {
        if (isNumber()) {
            return std::get<double>(value_);
        }
        return std::nullopt;
    }
    
    /**
     * @brief 获取整数值
     * @return std::optional<int64_t> 整数值，如果类型不匹配则返回空
     */
    [[nodiscard]] std::optional<int64_t> getInteger() const {
        if (isInteger()) {
            return std::get<int64_t>(value_);
        }
        return std::nullopt;
    }
    
    /**
     * @brief 获取布尔值
     * @return std::optional<bool> 布尔值，如果类型不匹配则返回空
     */
    [[nodiscard]] std::optional<bool> getBoolean() const {
        if (isBoolean()) {
            return std::get<bool>(value_);
        }
        return std::nullopt;
    }
    
    /**
     * @brief 转换为字符串表示
     * @return std::string 字符串表示
     */
    [[nodiscard]] std::string toString() const;
    
    /**
     * @brief 尝试转换为数字
     * @return std::optional<double> 转换后的数字，如果转换失败则返回空
     */
    [[nodiscard]] std::optional<double> toNumber() const;
    
    /**
     * @brief 尝试转换为整数
     * @return std::optional<int64_t> 转换后的整数，如果转换失败则返回空
     */
    [[nodiscard]] std::optional<int64_t> toInteger() const;
    
    /**
     * @brief 尝试转换为布尔值
     * @return std::optional<bool> 转换后的布尔值，如果转换失败则返回空
     */
    [[nodiscard]] std::optional<bool> toBoolean() const;
    
    /**
     * @brief 比较两个单元格是否相等
     * @param other 另一个单元格
     * @return bool 是否相等
     */
    [[nodiscard]] bool operator==(const Cell& other) const {
        return position_ == other.position_ && value_ == other.value_;
    }
    
    /**
     * @brief 比较两个单元格是否不相等
     * @param other 另一个单元格
     * @return bool 是否不相等
     */
    [[nodiscard]] bool operator!=(const Cell& other) const {
        return !(*this == other);
    }
};

} // namespace TinaXlsx