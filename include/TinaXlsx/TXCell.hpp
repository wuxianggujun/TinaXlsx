#pragma once

#include "TXTypes.hpp"
#include <string>
#include <variant>
#include <memory>
#include <utility>

namespace TinaXlsx {

// Forward declarations
class TXFormula;
class TXNumberFormat;
class TXSheet;

/**
 * @brief Excel单元格类
 * 
 * 单元格数据的抽象，提供读写单个单元格的能力，包括公式、格式化和合并功能
 */
class TXCell {
public:
    // ==================== 类型别名 ====================
    using CellValue = cell_value_t;

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
     * @brief 数字格式类型（兼容性保留，建议使用TXNumberFormat）
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

    // ==================== 公式功能 ====================

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
     * @brief 获取公式对象
     * @return 公式对象指针，如果不是公式返回nullptr
     */
    const TXFormula* getFormulaObject() const;

    /**
     * @brief 设置公式对象
     * @param formula 公式对象
     */
    void setFormulaObject(std::unique_ptr<TXFormula> formula);

    /**
     * @brief 计算公式结果
     * @param sheet 当前工作表
     * @param currentRow 当前行号
     * @param currentCol 当前列号
     * @return 计算结果
     */
    CellValue evaluateFormula(const class TXSheet* sheet, uint32_t currentRow, uint32_t currentCol);

    // ==================== 数字格式化功能 ====================

    /**
     * @brief 获取数字格式（兼容性）
     * @return 数字格式类型
     */
    NumberFormat getNumberFormat() const;

    /**
     * @brief 设置数字格式（兼容性）
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
     * @brief 获取数字格式化对象
     * @return 格式化对象指针
     */
    const TXNumberFormat* getNumberFormatObject() const;

    /**
     * @brief 设置数字格式化对象
     * @param numberFormat 格式化对象
     */
    void setNumberFormatObject(std::unique_ptr<TXNumberFormat> numberFormat);

    /**
     * @brief 获取格式化后的显示值
     * @return 格式化后的字符串
     */
    std::string getFormattedValue() const;

    /**
     * @brief 设置预定义格式
     * @param type 格式类型
     * @param decimalPlaces 小数位数
     * @param useThousandSeparator 是否使用千位分隔符
     */
    void setPredefinedFormat(NumberFormat type, int decimalPlaces = 2, bool useThousandSeparator = true);

    // ==================== 合并单元格功能 ====================

    /**
     * @brief 检查是否为合并单元格的一部分
     * @return 是合并单元格返回true，否则返回false
     */
    bool isMerged() const;

    /**
     * @brief 设置合并状态
     * @param merged 是否为合并单元格
     */
    void setMerged(bool merged);

    /**
     * @brief 检查是否为合并区域的主单元格（左上角）
     * @return 是主单元格返回true，否则返回false
     */
    bool isMasterCell() const;

    /**
     * @brief 设置为主单元格
     * @param master 是否为主单元格
     */
    void setMasterCell(bool master);

    /**
     * @brief 获取主单元格的行列位置
     * @return 主单元格位置，如果不是合并单元格返回{0,0}
     */
    std::pair<uint32_t, uint32_t> getMasterCellPosition() const;

    /**
     * @brief 设置主单元格位置
     * @param row 主单元格行号
     * @param col 主单元格列号
     */
    void setMasterCellPosition(uint32_t row, uint32_t col);

    // ==================== 样式方法 ====================
    
    /**
     * @brief 检查单元格是否有样式
     * @return 有样式返回true，否则返回false
     */
    bool hasStyle() const;
    
    /**
     * @brief 获取单元格样式索引
     * @return 样式索引，如果没有样式返回0
     */
    uint32_t getStyleIndex() const;
    
    /**
     * @brief 设置单元格样式索引
     * @param index 样式索引
     */
    void setStyleIndex(uint32_t index);

    // ==================== 工具方法 ====================

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
     * @brief 克隆单元格
     * @return 新的单元格副本
     */
    std::unique_ptr<TXCell> clone() const;

    /**
     * @brief 复制格式到另一个单元格
     * @param target 目标单元格
     */
    void copyFormatTo(TXCell& target) const;

    /**
     * @brief 比较单元格值（忽略格式）
     * @param other 另一个单元格
     * @return 值相等返回true，否则返回false
     */
    bool isValueEqual(const TXCell& other) const;

    // ==================== 类型转换操作符 ====================

    /**
     * @brief 类型转换操作符
     */
    operator std::string() const { return getFormattedValue(); }
    operator double() const { return getNumberValue(); }
    operator int64_t() const { return getIntegerValue(); }
    operator bool() const { return getBooleanValue(); }

    // ==================== 赋值操作符 ====================

    /**
     * @brief 赋值操作符
     */
    TXCell& operator=(const std::string& value);
    TXCell& operator=(const char* value);
    TXCell& operator=(double value);
    TXCell& operator=(int64_t value);
    TXCell& operator=(int value);
    TXCell& operator=(bool value);

    // ==================== 比较操作符 ====================

    /**
     * @brief 比较操作符
     */
    bool operator==(const TXCell& other) const;
    bool operator!=(const TXCell& other) const;
    bool operator<(const TXCell& other) const;
    bool operator<=(const TXCell& other) const;
    bool operator>(const TXCell& other) const;
    bool operator>=(const TXCell& other) const;

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

} // namespace TinaXlsx 