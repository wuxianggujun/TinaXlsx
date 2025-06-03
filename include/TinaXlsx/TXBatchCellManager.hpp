//
// @file TXBatchCellManager.hpp
// @brief 批处理单元格管理器 - 高性能批量操作
//

#pragma once

#include "TXUltraCompactCell.hpp"
#include "TXTypes.hpp"
#include "TXCoordinate.hpp"
#include <vector>
#include <memory>
#include <atomic>
#include <mutex>
#include <chrono>
#include <unordered_map>
#include <array>

namespace TinaXlsx {

/**
 * @brief 单元格数据结构（用于批处理输入）
 */
struct CellData {
    cell_value_t value;
    TXCoordinate coordinate;
    uint8_t style_index = 0;
    bool is_formula = false;
    std::string formula_text;
    
    CellData() = default;
    CellData(const cell_value_t& val, const TXCoordinate& coord) 
        : value(val), coordinate(coord) {}
    CellData(const cell_value_t& val, uint16_t row, uint16_t col)
        : value(val), coordinate(row_t{static_cast<u32>(row)}, column_t{static_cast<u32>(col)}) {}
};

/**
 * @brief 单元格范围
 */
struct CellRange {
    uint16_t start_row = 0;
    uint16_t start_col = 0;
    uint16_t end_row = 0;
    uint16_t end_col = 0;
    
    CellRange() = default;
    CellRange(uint16_t sr, uint16_t sc, uint16_t er, uint16_t ec)
        : start_row(sr), start_col(sc), end_row(er), end_col(ec) {}
    
    size_t getCellCount() const {
        return static_cast<size_t>(end_row - start_row + 1) * 
               static_cast<size_t>(end_col - start_col + 1);
    }
};

/**
 * @brief 内存块管理器
 */
class TXMemoryChunk {
public:
    static constexpr size_t CHUNK_SIZE = 64 * 1024 * 1024;  // 64MB块
    static constexpr size_t MAX_CHUNKS = 64;                // 最大64块 = 4GB
    
    TXMemoryChunk();
    ~TXMemoryChunk();
    
    /**
     * @brief 分配内存
     */
    void* allocate(size_t size);
    
    /**
     * @brief 释放所有内存
     */
    void clear();
    
    /**
     * @brief 获取当前内存使用量
     */
    size_t getCurrentUsage() const { return total_allocated_.load(); }
    
    /**
     * @brief 检查内存限制
     */
    bool checkMemoryLimit(size_t requested_size) const;

private:
    struct Chunk {
        std::unique_ptr<char[]> data;
        size_t used = 0;
    };
    
    std::array<Chunk, MAX_CHUNKS> chunks_;
    std::atomic<size_t> current_chunk_{0};
    std::atomic<size_t> total_allocated_{0};
    mutable std::mutex mutex_;
    
    static constexpr size_t MAX_MEMORY = 4ULL * 1024 * 1024 * 1024;  // 4GB
};

/**
 * @brief 字符串缓冲区管理器
 */
class TXStringBuffer {
public:
    TXStringBuffer();
    
    /**
     * @brief 添加字符串，返回偏移量
     */
    uint32_t addString(const std::string& str);
    
    /**
     * @brief 根据偏移量获取字符串
     */
    std::string_view getString(uint32_t offset) const;
    
    /**
     * @brief 获取缓冲区指针
     */
    const char* getBuffer() const { return buffer_.data(); }
    
    /**
     * @brief 获取缓冲区大小
     */
    size_t getSize() const { return buffer_.size(); }
    
    /**
     * @brief 预分配空间
     */
    void reserve(size_t size);
    
    /**
     * @brief 清空缓冲区
     */
    void clear();

private:
    std::vector<char> buffer_;
    std::unordered_map<std::string, uint32_t> offset_map_;
    mutable std::mutex mutex_;
};

/**
 * @brief 批处理单元格管理器
 * 
 * 核心功能：
 * - 批量处理单元格数据
 * - 内存使用控制（4GB限制）
 * - 高性能数据编码/解码
 * - 缓存友好的数据布局
 */
class TXBatchCellManager {
public:
    /**
     * @brief 批处理统计信息
     */
    struct BatchStats {
        size_t cells_processed = 0;
        double avg_time_per_cell = 0.0;
        size_t memory_used = 0;
        double memory_efficiency = 0.0;
        size_t string_pool_size = 0;
        size_t cache_hit_rate = 0;
        
        std::chrono::steady_clock::time_point start_time;
        std::chrono::steady_clock::time_point end_time;
    };
    
    TXBatchCellManager();
    ~TXBatchCellManager();
    
    // ==================== 核心批处理方法 ====================
    
    /**
     * @brief 批量设置单元格（主要API）
     * @param cells 单元格数据数组
     * @return 处理的单元格数量
     */
    size_t setBatchCells(const std::vector<CellData>& cells);
    
    /**
     * @brief 批量获取单元格
     * @param range 单元格范围
     * @return 单元格数据数组
     */
    std::vector<CellData> getBatchCells(const CellRange& range) const;
    
    /**
     * @brief 获取单个单元格
     * @param coord 坐标
     * @return 单元格数据
     */
    CellData getCell(const TXCoordinate& coord) const;
    
    /**
     * @brief 设置单个单元格
     * @param data 单元格数据
     */
    void setCell(const CellData& data);
    
    // ==================== 内存管理 ====================
    
    /**
     * @brief 压缩内存
     */
    void compactMemory();
    
    /**
     * @brief 清空所有数据
     */
    void clear();
    
    /**
     * @brief 获取内存使用量
     */
    size_t getMemoryUsage() const;
    
    /**
     * @brief 检查内存限制
     */
    bool checkMemoryLimit(size_t additional_size) const;
    
    // ==================== 性能统计 ====================
    
    /**
     * @brief 获取统计信息
     */
    BatchStats getStats() const;
    
    /**
     * @brief 重置统计信息
     */
    void resetStats();
    
    /**
     * @brief 开始性能计时
     */
    void startTiming();
    
    /**
     * @brief 结束性能计时
     */
    void endTiming();
    
    // ==================== 配置选项 ====================
    
    /**
     * @brief 设置批处理大小
     */
    void setBatchSize(size_t size) { batch_size_ = size; }
    
    /**
     * @brief 获取批处理大小
     */
    size_t getBatchSize() const { return batch_size_; }
    
    /**
     * @brief 启用/禁用SIMD优化
     */
    void enableSIMD(bool enable) { simd_enabled_ = enable; }
    
    /**
     * @brief 检查SIMD是否启用
     */
    bool isSIMDEnabled() const { return simd_enabled_; }

private:
    // ==================== 内部数据结构 ====================
    
    // 单元格存储
    std::vector<UltraCompactCell> cells_;
    
    // 内存管理
    std::unique_ptr<TXMemoryChunk> memory_chunk_;
    std::unique_ptr<TXStringBuffer> string_buffer_;
    
    // 索引和查找
    std::unordered_map<uint32_t, size_t> coordinate_index_;  // 坐标->索引映射
    
    // 统计信息
    mutable BatchStats stats_;
    mutable std::mutex stats_mutex_;
    
    // 配置
    size_t batch_size_ = 10000;
    bool simd_enabled_ = true;
    
    // 线程安全
    mutable std::mutex data_mutex_;
    
    // ==================== 内部辅助方法 ====================
    
    /**
     * @brief 坐标转换为键值
     */
    uint32_t coordinateToKey(const TXCoordinate& coord) const;
    
    /**
     * @brief 键值转换为坐标
     */
    TXCoordinate keyToCoordinate(uint32_t key) const;
    
    /**
     * @brief 查找单元格索引
     */
    size_t findCellIndex(const TXCoordinate& coord) const;
    
    /**
     * @brief 添加新单元格
     */
    size_t addNewCell(const CellData& data);
    
    /**
     * @brief 更新现有单元格
     */
    void updateExistingCell(size_t index, const CellData& data);
    
    /**
     * @brief 编码单元格数据
     */
    UltraCompactCell encodeCellData(const CellData& data);
    
    /**
     * @brief 解码单元格数据
     */
    CellData decodeCellData(const UltraCompactCell& cell) const;
    
    /**
     * @brief 批量编码（SIMD优化）
     */
    void encodeBatchSIMD(const std::vector<CellData>& input,
                        std::vector<UltraCompactCell>& output);
    
    /**
     * @brief 批量解码（SIMD优化）
     */
    void decodeBatchSIMD(const std::vector<UltraCompactCell>& input,
                        std::vector<CellData>& output) const;
    
    /**
     * @brief 更新统计信息
     */
    void updateStats(size_t cells_count, 
                    std::chrono::steady_clock::time_point start,
                    std::chrono::steady_clock::time_point end);
};

} // namespace TinaXlsx
