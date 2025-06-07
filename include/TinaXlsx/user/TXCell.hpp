//
// @file TXCell.hpp
// @brief 🚀 用户层单元格类 - 轻量级16字节高性能设计
//

#pragma once

#include "TinaXlsx/TXCoordinate.hpp"
#include "TinaXlsx/TXVariant.hpp"
#include "TinaXlsx/TXResult.hpp"
#include "TinaXlsx/TXCoordUtils.hpp"
#include <string>
#include <string_view>

namespace TinaXlsx {

// 前向声明
class TXInMemorySheet;

/**
 * @brief 🚀 用户层单元格类 - 轻量级引用设计
 * 
 * 设计理念：
 * - 16字节轻量级：只存储引用和坐标，不存储数据
 * - 高性能：直接委托给底层TXInMemorySheet
 * - 实时性：总是获取最新数据
 * - 链式调用：支持流畅的API设计
 * 
 * 内存布局：
 * - TXInMemorySheet& sheet_  (8字节引用)
 * - TXCoordinate coord_       (8字节：row+col各4字节)
 * - 总计：16字节
 */
class TXCell {
public:
    // ==================== 构造和析构 ====================
    
    /**
     * @brief 构造单元格引用
     * @param sheet 底层工作表引用
     * @param coord 单元格坐标
     */
    TXCell(TXInMemorySheet& sheet, const TXCoordinate& coord);
    
    /**
     * @brief 从Excel格式坐标构造
     * @param sheet 底层工作表引用
     * @param excel_coord Excel格式坐标 (如 "A1", "B2")
     */
    TXCell(TXInMemorySheet& sheet, std::string_view excel_coord);
    
    // 默认拷贝和移动语义
    TXCell(const TXCell&) = default;
    TXCell& operator=(const TXCell&) = default;
    TXCell(TXCell&&) = default;
    TXCell& operator=(TXCell&&) = default;
    
    ~TXCell() = default;

    // ==================== 值操作 ====================
    
    /**
     * @brief 🚀 设置数值 - 支持链式调用
     */
    TXCell& setValue(double value);
    
    /**
     * @brief 🚀 设置字符串 - 支持链式调用
     */
    TXCell& setValue(const std::string& value);
    TXCell& setValue(std::string_view value);
    TXCell& setValue(const char* value);
    
    /**
     * @brief 🚀 设置布尔值 - 支持链式调用
     */
    TXCell& setValue(bool value);
    
    /**
     * @brief 🚀 设置TXVariant值 - 支持链式调用
     */
    TXCell& setValue(const TXVariant& value);
    
    /**
     * @brief 🚀 设置公式 - 支持链式调用
     */
    TXCell& setFormula(const std::string& formula);
    
    /**
     * @brief 🚀 获取值
     */
    TXVariant getValue() const;
    
    /**
     * @brief 🚀 获取公式
     */
    std::string getFormula() const;
    
    /**
     * @brief 🚀 获取单元格类型
     */
    TXVariant::Type getType() const;
    
    /**
     * @brief 🚀 检查是否为空
     */
    bool isEmpty() const;
    
    /**
     * @brief 🚀 清除单元格内容
     */
    TXCell& clear();

    // ==================== 坐标信息 ====================
    
    /**
     * @brief 🚀 获取坐标对象
     */
    const TXCoordinate& getCoordinate() const { return coord_; }
    
    /**
     * @brief 🚀 获取Excel格式地址 (如 "A1", "B2")
     */
    std::string getAddress() const;
    
    /**
     * @brief 🚀 获取行索引 (0-based)
     */
    uint32_t getRow() const { return coord_.getRow().index() - 1; }

    /**
     * @brief 🚀 获取列索引 (0-based)
     */
    uint32_t getColumn() const { return coord_.getCol().index() - 1; }

    // ==================== 便捷操作符 ====================
    
    /**
     * @brief 🚀 赋值操作符 - 支持多种类型
     */
    TXCell& operator=(double value) { return setValue(value); }
    TXCell& operator=(const std::string& value) { return setValue(value); }
    TXCell& operator=(std::string_view value) { return setValue(value); }
    TXCell& operator=(const char* value) { return setValue(value); }
    TXCell& operator=(bool value) { return setValue(value); }
    TXCell& operator=(const TXVariant& value) { return setValue(value); }
    
    /**
     * @brief 🚀 类型转换操作符
     */
    operator TXVariant() const { return getValue(); }
    
    /**
     * @brief 🚀 比较操作符
     */
    bool operator==(const TXCell& other) const;
    bool operator!=(const TXCell& other) const { return !(*this == other); }

    // ==================== 数学操作 ====================
    
    /**
     * @brief 🚀 数学运算 - 支持链式调用
     */
    TXCell& add(double value);
    TXCell& subtract(double value);
    TXCell& multiply(double value);
    TXCell& divide(double value);
    
    /**
     * @brief 🚀 数学运算操作符
     */
    TXCell& operator+=(double value) { return add(value); }
    TXCell& operator-=(double value) { return subtract(value); }
    TXCell& operator*=(double value) { return multiply(value); }
    TXCell& operator/=(double value) { return divide(value); }

    // ==================== 格式化 (预留接口) ====================
    
    /**
     * @brief 🚀 设置数字格式 (预留)
     */
    TXCell& setNumberFormat(const std::string& format);
    
    /**
     * @brief 🚀 设置字体颜色 (预留)
     */
    TXCell& setFontColor(uint32_t color);
    
    /**
     * @brief 🚀 设置背景颜色 (预留)
     */
    TXCell& setBackgroundColor(uint32_t color);

    // ==================== 调试和诊断 ====================
    
    /**
     * @brief 🚀 获取调试信息
     */
    std::string toString() const;
    
    /**
     * @brief 🚀 验证单元格引用是否有效
     */
    bool isValid() const;

private:
    TXInMemorySheet& sheet_;  // 8字节：底层工作表引用
    TXCoordinate coord_;      // 8字节：单元格坐标
    
    // 总计：16字节轻量级设计
    
    // ==================== 内部辅助方法 ====================
    
    /**
     * @brief 内部错误处理
     */
    void handleError(const std::string& operation, const TXError& error) const;
};

/**
 * @brief 🚀 便捷的单元格创建函数
 */
inline TXCell makeCell(TXInMemorySheet& sheet, const TXCoordinate& coord) {
    return TXCell(sheet, coord);
}

inline TXCell makeCell(TXInMemorySheet& sheet, std::string_view excel_coord) {
    return TXCell(sheet, excel_coord);
}

inline TXCell makeCell(TXInMemorySheet& sheet, uint32_t row, uint32_t col) {
    // 转换0-based输入为1-based内部表示
    return TXCell(sheet, TXCoordinate(row_t(row + 1), column_t(static_cast<uint32_t>(col + 1))));
}

} // namespace TinaXlsx
