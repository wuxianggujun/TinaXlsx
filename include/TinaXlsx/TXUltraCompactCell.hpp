//
// @file TXUltraCompactCell.hpp
// @brief 超紧凑单元格实现 - 16字节批处理优化版本
//

#pragma once

#include "TXTypes.hpp"
#include "TXCoordinate.hpp"
#include <cstdint>
#include <string>
#include <variant>
#include <vector>

namespace TinaXlsx {

// 前向声明
class TXFormula;
class TXNumberFormat;

/**
 * @brief 超紧凑单元格数据结构（16字节）
 * 
 * 设计目标：
 * - 固定16字节大小，缓存友好
 * - 支持所有基本数据类型
 * - 高效的批处理操作
 * - 内存局部性优化
 */
struct UltraCompactCell {
    // ==================== 数据布局（16字节总计）====================
    
    // 第一个8字节：主要数据
    union PrimaryData {
        // 数值类型直接存储
        double number_value;
        int64_t integer_value;

        // 字符串类型存储偏移和元数据
        struct StringData {
            uint32_t offset;        // 字符串缓冲区偏移量
            uint16_t length;        // 字符串长度
            uint16_t padding;       // 填充到8字节
        } string;

        // 布尔类型
        struct BooleanData {
            uint8_t  value;         // 布尔值（0或1）
            uint8_t  padding[7];    // 填充到8字节
        } boolean;

        // 原始数据访问
        uint64_t raw_primary;
    } primary;
    
    // 第二个8字节：元信息和坐标
    struct SecondaryData {
        uint8_t  cell_type;        // 单元格类型
        uint8_t  style_index;      // 样式索引
        uint8_t  flags;            // 标志位（has_style, is_formula, is_merged等）
        uint8_t  formula_offset_low; // 公式偏移量低8位
        uint16_t row;              // 行号
        uint16_t col;              // 列号
    } secondary;

    // 公式偏移量的高24位存储在primary的padding中（仅当类型为Formula时）
    // 这样我们可以存储完整的32位偏移量而不破坏元数据
    
    // ==================== 单元格类型枚举 ====================
    
    enum class CellType : uint8_t {
        Empty = 0,
        String = 1,
        Number = 2,
        Integer = 3,
        Boolean = 4,
        Formula = 5,
        Error = 6,
        Reserved = 7
    };
    
    // ==================== 构造函数 ====================
    
    /**
     * @brief 默认构造函数 - 创建空单元格
     */
    UltraCompactCell();
    
    /**
     * @brief 从字符串构造
     */
    UltraCompactCell(const std::string& value, uint32_t string_offset);
    
    /**
     * @brief 从数值构造
     */
    UltraCompactCell(double value);
    
    /**
     * @brief 从整数构造
     */
    UltraCompactCell(int64_t value);
    
    /**
     * @brief 从布尔值构造
     */
    UltraCompactCell(bool value);
    
    // ==================== 数据访问方法 ====================
    
    /**
     * @brief 获取单元格类型
     */
    CellType getType() const;
    
    /**
     * @brief 设置单元格类型
     */
    void setType(CellType type);
    
    /**
     * @brief 检查是否为空
     */
    bool isEmpty() const { return getType() == CellType::Empty; }
    
    /**
     * @brief 获取数值（仅当类型为Number时有效）
     */
    double getNumberValue() const;
    
    /**
     * @brief 获取整数值（仅当类型为Integer时有效）
     */
    int64_t getIntegerValue() const;
    
    /**
     * @brief 获取布尔值（仅当类型为Boolean时有效）
     */
    bool getBooleanValue() const;
    
    /**
     * @brief 获取字符串偏移量（仅当类型为String时有效）
     */
    uint32_t getStringOffset() const;
    
    /**
     * @brief 获取字符串长度（仅当类型为String时有效）
     */
    uint16_t getStringLength() const;
    
    // ==================== 样式和属性 ====================
    
    /**
     * @brief 检查是否有样式
     */
    bool hasStyle() const;
    
    /**
     * @brief 设置样式状态
     */
    void setHasStyle(bool has_style);
    
    /**
     * @brief 获取样式索引
     */
    uint8_t getStyleIndex() const;
    
    /**
     * @brief 设置样式索引
     */
    void setStyleIndex(uint8_t index);
    
    /**
     * @brief 检查是否为公式
     */
    bool isFormula() const;
    
    /**
     * @brief 设置公式状态
     */
    void setIsFormula(bool is_formula);
    
    /**
     * @brief 获取公式偏移量
     */
    uint32_t getFormulaOffset() const;
    
    /**
     * @brief 设置公式偏移量
     */
    void setFormulaOffset(uint32_t offset);
    
    /**
     * @brief 检查是否为合并单元格
     */
    bool isMerged() const;
    
    /**
     * @brief 设置合并状态
     */
    void setIsMerged(bool is_merged);
    
    // ==================== 位置信息 ====================
    
    /**
     * @brief 获取行号
     */
    uint16_t getRow() const { return secondary.row; }

    /**
     * @brief 设置行号
     */
    void setRow(uint16_t row) { secondary.row = row; }

    /**
     * @brief 获取列号
     */
    uint16_t getCol() const { return secondary.col; }

    /**
     * @brief 设置列号
     */
    void setCol(uint16_t col) { secondary.col = col; }

    /**
     * @brief 获取坐标
     */
    TXCoordinate getCoordinate() const {
        row_t row_obj{static_cast<u32>(getRow())};
        column_t col_obj{static_cast<u32>(getCol())};
        return TXCoordinate{row_obj, col_obj};
    }

    /**
     * @brief 设置坐标
     */
    void setCoordinate(const TXCoordinate& coord) {
        setRow(coord.getRow().index());
        setCol(coord.getCol().index());
    }
    
    // ==================== 批处理优化方法 ====================
    
    /**
     * @brief 批量编码数据（SIMD优化）
     */
    static void encodeBatch(const std::vector<cell_value_t>& values,
                           const std::vector<TXCoordinate>& coords,
                           const char* string_buffer,
                           UltraCompactCell* output,
                           size_t count);
    
    /**
     * @brief 批量解码数据（SIMD优化）
     */
    static void decodeBatch(const UltraCompactCell* input,
                           const char* string_buffer,
                           std::vector<cell_value_t>& values,
                           std::vector<TXCoordinate>& coords,
                           size_t count);
    
    // ==================== 内存和性能 ====================
    
    /**
     * @brief 获取内存使用量（固定16字节）
     */
    static constexpr size_t getMemoryUsage() { return 16; }
    
    /**
     * @brief 清空单元格数据
     */
    void clear();

    // ==================== 比较操作符 ====================

    /**
     * @brief 相等比较操作符
     */
    bool operator==(const UltraCompactCell& other) const;

    /**
     * @brief 不等比较操作符
     */
    bool operator!=(const UltraCompactCell& other) const;

private:
    // ==================== 内部辅助方法 ====================
    
    /**
     * @brief 获取类型字段的指针（根据当前类型）
     */
    uint8_t* getTypeField();
    const uint8_t* getTypeField() const;
    
    /**
     * @brief 获取样式字段的指针（根据当前类型）
     */
    uint8_t* getStyleField();
    const uint8_t* getStyleField() const;
};

// ==================== 编译时验证 ====================

/**
 * @brief 验证数据结构大小
 */
static_assert(sizeof(UltraCompactCell) == 16, "UltraCompactCell must be exactly 16 bytes");

/**
 * @brief 验证内存对齐
 */
static_assert(alignof(UltraCompactCell) <= 8, "UltraCompactCell alignment should be <= 8 bytes");

} // namespace TinaXlsx
