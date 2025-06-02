//
// @file TXCompactCell.hpp
// @brief 紧凑型单元格实现 - 内存优化版本
//

#pragma once

#include "TXTypes.hpp"
#include "TXCoordinate.hpp"
#include <variant>
#include <memory>
#include <string>
#include <unordered_map>

namespace TinaXlsx {

// 前向声明
class TXFormula;
class TXNumberFormat;

/**
 * @brief 紧凑型单元格值存储
 * 
 * 使用位域和联合体优化内存布局
 */
class TXCompactCell {
public:
    // 单元格类型（使用更少的位）
    enum class CellType : uint8_t {
        Empty = 0,
        String = 1,
        Number = 2,
        Integer = 3,
        Boolean = 4,
        Formula = 5
    };
    
    // 紧凑的单元格值（24字节）
    using CellValue = std::variant<std::monostate, std::string, double, int64_t, bool>;

public:
    TXCompactCell();
    explicit TXCompactCell(const cell_value_t& value);
    ~TXCompactCell(); // 需要在实现文件中定义，因为使用了Pimpl
    
    // 移动构造和赋值（避免不必要的拷贝）
    TXCompactCell(TXCompactCell&& other) noexcept;
    TXCompactCell& operator=(TXCompactCell&& other) noexcept;
    
    // 拷贝构造和赋值
    TXCompactCell(const TXCompactCell& other);
    TXCompactCell& operator=(const TXCompactCell& other);

    // ==================== 值操作 ====================
    
    /**
     * @brief 设置单元格值
     */
    void setValue(const cell_value_t& value);
    
    /**
     * @brief 获取单元格值
     */
    cell_value_t getValue() const;
    
    /**
     * @brief 获取单元格类型
     */
    CellType getType() const { return static_cast<CellType>(flags_.type_); }

    /**
     * @brief 检查是否为空
     */
    bool isEmpty() const { return flags_.type_ == static_cast<uint8_t>(CellType::Empty); }

    // ==================== 样式操作 ====================
    
    /**
     * @brief 设置样式索引
     */
    void setStyleIndex(uint32_t index);
    
    /**
     * @brief 获取样式索引
     */
    uint32_t getStyleIndex() const;
    
    /**
     * @brief 检查是否有样式
     */
    bool hasStyle() const { return flags_.has_style_ == 1; }

    // ==================== 合并单元格操作 ====================
    
    /**
     * @brief 设置为合并单元格
     */
    void setMerged(bool is_master, uint16_t master_row = 0, uint16_t master_col = 0);
    
    /**
     * @brief 检查是否为合并单元格
     */
    bool isMerged() const { return flags_.is_merged_ == 1; }

    /**
     * @brief 检查是否为主单元格
     */
    bool isMasterCell() const { return flags_.is_master_cell_ == 1; }

    // ==================== 扩展数据操作 ====================

    /**
     * @brief 设置公式
     */
    void setFormula(std::unique_ptr<TXFormula> formula);

    /**
     * @brief 获取公式
     */
    const TXFormula* getFormula() const;

    /**
     * @brief 设置数字格式
     */
    void setNumberFormat(std::unique_ptr<TXNumberFormat> format);

    /**
     * @brief 获取数字格式
     */
    const TXNumberFormat* getNumberFormat() const;

    // ==================== 兼容性接口 ====================

    /**
     * @brief 设置锁定状态
     */
    void setLocked(bool locked) { flags_.is_locked_ = locked ? 1 : 0; }

    /**
     * @brief 获取锁定状态
     */
    bool isLocked() const { return flags_.is_locked_ == 1; }

    /**
     * @brief 获取格式化值（兼容接口）
     */
    std::string getFormattedValue() const;

    /**
     * @brief 检查是否有公式（兼容接口）
     */
    bool hasFormula() const;

    /**
     * @brief 设置公式（字符串版本，兼容接口）
     */
    void setFormula(const std::string& formulaText);

    /**
     * @brief 获取公式文本（兼容接口）
     */
    std::string getFormulaText() const;

    /**
     * @brief 设置数字格式对象（兼容接口）
     */
    void setNumberFormatObject(std::unique_ptr<TXNumberFormat> format);

    /**
     * @brief 获取数字格式对象（兼容接口）
     */
    const TXNumberFormat* getNumberFormatObject() const;

    /**
     * @brief 获取字符串值（兼容接口）
     */
    std::string getStringValue() const;

    /**
     * @brief 获取数字值（兼容接口）
     */
    double getNumberValue() const;

    /**
     * @brief 获取整数值（兼容接口）
     */
    int64_t getIntegerValue() const;

    /**
     * @brief 获取布尔值（兼容接口）
     */
    bool getBooleanValue() const;

    /**
     * @brief 检查是否为公式（兼容接口）
     */
    bool isFormula() const;

    /**
     * @brief 获取公式对象（兼容接口）
     */
    const TXFormula* getFormulaObject() const;

    // ==================== 内存统计 ====================
    
    /**
     * @brief 获取内存使用量
     */
    size_t getMemoryUsage() const;
    
    /**
     * @brief 获取紧凑度（相比原始TXCell的内存节省比例）
     */
    static double getCompactRatio();

private:
    // ==================== 紧凑存储结构 ====================
    
    // 核心数据（32字节）
    CellValue value_;                    // 24字节（std::variant）
    
    // 位域压缩状态（4字节）
    struct {
        uint8_t type_ : 3;               // 单元格类型（0-7）
        uint8_t has_style_ : 1;          // 是否有样式
        uint8_t is_merged_ : 1;          // 是否合并
        uint8_t is_master_cell_ : 1;     // 是否主单元格
        uint8_t is_locked_ : 1;          // 是否锁定
        uint8_t reserved_ : 1;           // 保留位
        
        uint8_t master_row_high_;        // 主单元格行号高8位
        uint8_t master_row_low_;         // 主单元格行号低8位
        uint8_t master_col_;             // 主单元格列号（0-255）
    } flags_;
    
    // 扩展数据指针（仅在需要时分配）
    // 使用前向声明避免包含完整定义
    struct ExtendedData;
    std::unique_ptr<ExtendedData> extended_data_;  // 8字节指针
    
    // ==================== 辅助方法 ====================
    
    /**
     * @brief 确保扩展数据存在
     */
    void ensureExtendedData();
    
    /**
     * @brief 清理不需要的扩展数据
     */
    void cleanupExtendedData();
    
    /**
     * @brief 从cell_value_t推断类型
     */
    static CellType inferType(const cell_value_t& value);
};

/**
 * @brief 紧凑单元格管理器
 * 
 * 专门管理紧凑型单元格的高性能管理器
 */
class TXCompactCellManager {
public:
    using Coordinate = TXCoordinate;
    using CellValue = cell_value_t;
    
    // 使用标准哈希函数（已在TXCoordinate.hpp中定义）

public:
    TXCompactCellManager();
    ~TXCompactCellManager() = default;

    // ==================== 单元格访问 ====================
    
    /**
     * @brief 获取单元格（不创建）
     */
    const TXCompactCell* getCell(const Coordinate& coord) const;
    
    /**
     * @brief 获取或创建单元格
     */
    TXCompactCell* getOrCreateCell(const Coordinate& coord);
    
    /**
     * @brief 设置单元格值
     */
    bool setCellValue(const Coordinate& coord, const CellValue& value);

    // ==================== 批量操作 ====================
    
    /**
     * @brief 批量设置单元格值（高性能版本）
     */
    size_t setCellValues(const std::vector<std::pair<Coordinate, CellValue>>& values);
    
    /**
     * @brief 批量设置范围值
     */
    size_t setRangeValues(row_t startRow, column_t startCol, 
                         const std::vector<std::vector<CellValue>>& values);

    // ==================== 内存管理 ====================
    
    /**
     * @brief 获取内存使用统计
     */
    struct MemoryStats {
        size_t total_cells = 0;
        size_t memory_used = 0;
        size_t memory_saved = 0;
        double compact_ratio = 0.0;
    };
    
    MemoryStats getMemoryStats() const;
    
    /**
     * @brief 内存压缩（清理不需要的扩展数据）
     */
    size_t compactMemory();
    
    /**
     * @brief 预分配内存
     */
    void reserve(size_t expected_cells);

private:
    // 使用标准哈希表
    std::unordered_map<Coordinate, TXCompactCell> cells_;
    
    // 内存统计
    mutable MemoryStats cached_stats_;
    mutable bool stats_dirty_ = true;
};

} // namespace TinaXlsx
