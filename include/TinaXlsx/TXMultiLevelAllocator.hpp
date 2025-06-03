//
// @file TXMultiLevelAllocator.hpp
// @brief 多级内存分配器 - 整合Slab、Block、Chunk三级分配
//

#pragma once

#include "TXSlabAllocator.hpp"
#include "TXChunkAllocator.hpp"
#include <memory>
#include <atomic>
#include <mutex>
#include <chrono>

namespace TinaXlsx {

/**
 * @brief 多级分配器配置
 */
struct MultiLevelConfig {
    // 分级阈值
    static constexpr size_t TINY_THRESHOLD = 2048;      // <=2KB 使用Slab
    static constexpr size_t SMALL_THRESHOLD = 64 * 1024; // <=64KB 使用Block
    static constexpr size_t LARGE_THRESHOLD = 4 * 1024 * 1024; // >4MB 使用Chunk
    
    // 性能优化
    static constexpr bool ENABLE_TLS_CACHE = true;      // 启用线程本地缓存
    static constexpr bool ENABLE_STATISTICS = true;     // 启用详细统计
    static constexpr bool ENABLE_AUTO_COMPACT = true;   // 启用自动压缩
    
    // 压缩触发条件
    static constexpr double FRAGMENTATION_THRESHOLD = 0.3; // 30%碎片率触发压缩
    static constexpr size_t MIN_COMPACT_INTERVAL_MS = 5000; // 最小压缩间隔5秒
};

/**
 * @brief 块分配器（中等大小对象）
 */
class TXBlockAllocator {
public:
    static constexpr size_t BLOCK_SIZE = 1024 * 1024;   // 1MB块
    static constexpr size_t MAX_BLOCKS = 32;            // 最多32块
    
    TXBlockAllocator();
    ~TXBlockAllocator();
    
    void* allocate(size_t size);
    bool deallocate(void* ptr);
    size_t compact();
    void clear();
    
    size_t getTotalMemoryUsage() const;
    size_t getUsedMemorySize() const;
    double getFragmentationRatio() const;
    
    struct BlockStats {
        size_t total_blocks = 0;
        size_t active_blocks = 0;
        size_t total_memory = 0;
        size_t used_memory = 0;
        double memory_efficiency = 0.0;
        size_t allocation_count = 0;
        size_t deallocation_count = 0;
    };
    
    BlockStats getStats() const;

private:
    struct Block {
        std::unique_ptr<char[]> data;
        size_t used = 0;
        std::vector<std::pair<size_t, size_t>> free_chunks; // offset, size
    };
    
    std::array<std::unique_ptr<Block>, MAX_BLOCKS> blocks_;
    std::atomic<size_t> block_count_{0};
    std::atomic<size_t> allocation_count_{0};
    std::atomic<size_t> deallocation_count_{0};
    mutable std::mutex mutex_;
    
    Block* findAvailableBlock(size_t size);
    Block* createNewBlock();
    void* allocateFromBlock(Block* block, size_t size);
    bool deallocateFromBlock(Block* block, void* ptr);
};

/**
 * @brief 多级内存分配器
 */
class TXMultiLevelAllocator {
public:
    /**
     * @brief 综合统计信息
     */
    struct ComprehensiveStats {
        // 各级分配器统计
        TXSlabAllocator::SlabStats slab_stats;
        TXBlockAllocator::BlockStats block_stats;
        TXChunkAllocator::AllocationStats chunk_stats;
        
        // 综合指标
        size_t total_memory_usage = 0;
        size_t total_used_memory = 0;
        double overall_efficiency = 0.0;
        double overall_fragmentation = 0.0;
        
        // 分配分布
        size_t tiny_allocations = 0;    // Slab处理
        size_t small_allocations = 0;   // Block处理
        size_t medium_allocations = 0;  // Chunk处理
        size_t large_allocations = 0;   // Chunk处理
        
        // 性能指标
        double avg_allocation_time_us = 0.0;
        size_t allocations_per_second = 0;
        
        std::chrono::steady_clock::time_point start_time;
        std::chrono::steady_clock::time_point last_update_time;
    };
    
    TXMultiLevelAllocator();
    ~TXMultiLevelAllocator();
    
    // ==================== 分配接口 ====================
    
    /**
     * @brief 智能分配 - 自动选择最优分配器
     */
    void* allocate(size_t size);
    
    /**
     * @brief 智能释放 - 自动识别分配器类型
     */
    bool deallocate(void* ptr);
    
    /**
     * @brief 模板分配
     */
    template<typename T>
    T* allocate(size_t count = 1) {
        size_t size = sizeof(T) * count;
        return static_cast<T*>(allocate(size));
    }
    
    /**
     * @brief 批量分配
     */
    std::vector<void*> allocateBatch(const std::vector<size_t>& sizes);
    
    // ==================== 内存管理 ====================
    
    /**
     * @brief 全面压缩 - 压缩所有级别
     */
    size_t compactAll();
    
    /**
     * @brief 智能压缩 - 根据碎片率选择性压缩
     */
    size_t smartCompact();
    
    /**
     * @brief 清空所有内存
     */
    void clear();
    
    /**
     * @brief 获取总内存使用量
     */
    size_t getTotalMemoryUsage() const;
    
    /**
     * @brief 获取实际使用内存
     */
    size_t getUsedMemorySize() const;
    
    /**
     * @brief 获取整体内存效率
     */
    double getOverallEfficiency() const;
    
    /**
     * @brief 获取整体碎片率
     */
    double getOverallFragmentation() const;
    
    // ==================== 统计和监控 ====================
    
    /**
     * @brief 获取综合统计信息
     */
    ComprehensiveStats getComprehensiveStats() const;
    
    /**
     * @brief 生成详细报告
     */
    std::string generateDetailedReport() const;
    
    /**
     * @brief 获取性能指标
     */
    struct PerformanceMetrics {
        double allocations_per_second = 0.0;
        double avg_allocation_time_us = 0.0;
        double cache_hit_ratio = 0.0;
        size_t peak_memory_usage = 0;
        size_t current_memory_usage = 0;
    };
    
    PerformanceMetrics getPerformanceMetrics() const;
    
    // ==================== 配置管理 ====================
    
    /**
     * @brief 启用/禁用自动压缩
     */
    void enableAutoCompact(bool enable) { auto_compact_enabled_ = enable; }
    
    /**
     * @brief 设置碎片率阈值
     */
    void setFragmentationThreshold(double threshold) { fragmentation_threshold_ = threshold; }
    
    /**
     * @brief 启用/禁用TLS缓存
     */
    void enableTLSCache(bool enable) { tls_cache_enabled_ = enable; }

private:
    // ==================== 分配器实例 ====================
    
    std::unique_ptr<TXSlabAllocator> slab_allocator_;   // 微小对象 (<=2KB)
    std::unique_ptr<TXBlockAllocator> block_allocator_; // 小对象 (2KB-64KB)
    std::unique_ptr<TXChunkAllocator> chunk_allocator_; // 大对象 (>64KB)
    
    // ==================== 性能优化 ====================
    
    // 线程本地缓存
    thread_local static std::unique_ptr<TXSlabTLSCache> tls_cache_;
    bool tls_cache_enabled_ = MultiLevelConfig::ENABLE_TLS_CACHE;
    
    // 自动压缩
    bool auto_compact_enabled_ = MultiLevelConfig::ENABLE_AUTO_COMPACT;
    double fragmentation_threshold_ = MultiLevelConfig::FRAGMENTATION_THRESHOLD;
    std::atomic<std::chrono::steady_clock::time_point> last_compact_time_;
    
    // ==================== 统计信息 ====================
    
    mutable std::mutex stats_mutex_;
    mutable ComprehensiveStats cached_stats_;
    mutable std::chrono::steady_clock::time_point last_stats_update_;
    
    // 分配计数
    std::atomic<size_t> tiny_allocation_count_{0};
    std::atomic<size_t> small_allocation_count_{0};
    std::atomic<size_t> medium_allocation_count_{0};
    std::atomic<size_t> large_allocation_count_{0};
    
    // 性能计时
    std::atomic<size_t> total_allocation_time_us_{0};
    std::atomic<size_t> total_allocations_{0};
    std::atomic<size_t> peak_memory_usage_{0};
    
    // ==================== 内部方法 ====================
    
    /**
     * @brief 选择分配器类型
     */
    enum class AllocatorType {
        SLAB,   // 微小对象
        BLOCK,  // 小对象  
        CHUNK   // 大对象
    };
    
    AllocatorType selectAllocatorType(size_t size) const;
    
    /**
     * @brief 更新统计信息
     */
    void updateStats(AllocatorType type, size_t size, 
                    std::chrono::steady_clock::time_point start_time,
                    std::chrono::steady_clock::time_point end_time) const;
    
    /**
     * @brief 检查是否需要自动压缩
     */
    bool shouldAutoCompact() const;
    
    /**
     * @brief 更新峰值内存使用
     */
    void updatePeakMemoryUsage() const;
    
    /**
     * @brief 刷新缓存统计
     */
    void refreshCachedStats() const;
};

} // namespace TinaXlsx
