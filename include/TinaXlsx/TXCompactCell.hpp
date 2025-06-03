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
#include <mutex>
#include <vector>

namespace TinaXlsx {

// 前向声明
class TXFormula;
class TXNumberFormat;

/**
 * @brief 全局字符串池 - 用于优化字符串存储
 *
 * 将重复的字符串只存储一次，单元格中只保存4字节索引
 */
class TXStringPool {
public:
    /**
     * @brief 获取全局字符串池实例
     */
    static TXStringPool& getInstance();

    /**
     * @brief 将字符串加入池中，返回索引
     * @param str 要存储的字符串
     * @return 字符串在池中的索引
     */
    uint32_t intern(const std::string& str);

    /**
     * @brief 根据索引获取字符串
     * @param index 字符串索引
     * @return 字符串引用
     */
    const std::string& get(uint32_t index) const;

    /**
     * @brief 获取池中字符串数量
     */
    size_t size() const { return strings_.size(); }

    /**
     * @brief 清空字符串池
     */
    void clear();

    /**
     * @brief 获取内存使用统计
     */
    struct PoolStats {
        size_t string_count = 0;
        size_t total_memory = 0;
        size_t saved_memory = 0;
        double compression_ratio = 0.0;
    };

    PoolStats getStats() const;

private:
    TXStringPool() = default;

    std::vector<std::string> strings_;
    std::unordered_map<std::string, uint32_t> index_map_;
    mutable std::mutex mutex_;  // 线程安全

    static const std::string EMPTY_STRING;
};

/**
 * @brief 扩展数据池 - 存储公式和数字格式等大对象
 */
class TXExtendedDataPool {
public:
    struct ExtendedData {
        std::unique_ptr<TXFormula> formula;
        std::unique_ptr<TXNumberFormat> number_format;
        uint32_t style_index = 0;
    };

    /**
     * @brief 获取全局扩展数据池实例
     */
    static TXExtendedDataPool& getInstance();

    /**
     * @brief 分配扩展数据，返回偏移量
     */
    uint32_t allocate();

    /**
     * @brief 释放扩展数据
     */
    void deallocate(uint32_t offset);

    /**
     * @brief 获取扩展数据
     */
    ExtendedData* get(uint32_t offset);

    /**
     * @brief 获取扩展数据（只读）
     */
    const ExtendedData* get(uint32_t offset) const;

    /**
     * @brief 清空所有扩展数据
     */
    void clear();

private:
    TXExtendedDataPool() = default;

    std::vector<std::unique_ptr<ExtendedData>> pool_;
    std::vector<uint32_t> free_list_;
    mutable std::mutex mutex_;
};

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
    
    // 超紧凑的单元格值（12-16字节）- 使用字符串池索引
    using CompactCellValue = std::variant<
        std::monostate,     // 0字节 - 空单元格
        uint32_t,          // 4字节 - 字符串池索引
        double,            // 8字节 - 浮点数
        int64_t,           // 8字节 - 整数
        bool               // 1字节 - 布尔值
    >;

    // 保持兼容性的原始接口
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
    // ==================== 超紧凑存储结构（32字节总计）====================

    // 核心数据（16字节）- 使用小variant优化
    CompactCellValue compact_value_;     // 12-16字节（小variant）

    // 超压缩标志位（2字节）
    struct {
        uint16_t type_ : 3;              // 单元格类型（0-7）
        uint16_t has_style_ : 1;         // 是否有样式
        uint16_t is_merged_ : 1;         // 是否合并
        uint16_t is_master_cell_ : 1;    // 是否主单元格
        uint16_t is_locked_ : 1;         // 是否锁定
        uint16_t reserved_flags_ : 1;    // 保留标志位
        uint16_t style_index_ : 8;       // 内联样式索引（0-255）
    } flags_;

    // 合并单元格信息（4字节）
    struct {
        uint16_t master_row_;            // 主单元格行号（0-65535）
        uint16_t master_col_;            // 主单元格列号（0-65535）
    } merge_info_;

    // 扩展数据偏移量（4字节）- 使用偏移量而非指针
    uint32_t extended_offset_;           // 扩展数据池中的偏移量，0表示无扩展数据

    // 预留对齐（4字节）
    uint32_t reserved_;                  // 预留空间，确保32字节对齐

    // 静态字符串池引用
    static TXStringPool& getStringPool() { return TXStringPool::getInstance(); }
    
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

    /**
     * @brief 将兼容性CellValue转换为CompactCellValue
     */
    CompactCellValue convertToCompact(const CellValue& value);

    /**
     * @brief 将CompactCellValue转换为兼容性CellValue
     */
    CellValue convertFromCompact(const CompactCellValue& compact_value) const;
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
