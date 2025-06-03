//
// @file TXSlabCache.hpp
// @brief 第三阶段：智能缓存池 - 生产环境优化
//

#pragma once

#include "TXSlabAllocator.hpp"
#include <map>
#include <vector>
#include <memory>
#include <mutex>
#include <atomic>
#include <chrono>

namespace TinaXlsx {

/**
 * @brief 智能Slab缓存池
 */
class TXSlabCache {
public:
    /**
     * @brief 缓存配置
     */
    struct CacheConfig {
        size_t max_cached_slabs_per_size = 8;      // 每种大小最多缓存8个slab
        size_t min_cached_slabs_per_size = 2;      // 每种大小最少缓存2个slab
        std::chrono::seconds cache_timeout{300};   // 缓存超时5分钟
        bool enable_preallocation = true;          // 启用预分配
        bool enable_warmup = true;                 // 启用预热
    };
    
    /**
     * @brief 缓存统计
     */
    struct CacheStats {
        size_t cache_hits = 0;
        size_t cache_misses = 0;
        size_t total_cached_slabs = 0;
        size_t preallocation_count = 0;
        double hit_ratio = 0.0;
        
        std::array<size_t, SlabConfig::OBJECT_SIZES.size()> hits_per_size{};
        std::array<size_t, SlabConfig::OBJECT_SIZES.size()> misses_per_size{};
        std::array<size_t, SlabConfig::OBJECT_SIZES.size()> cached_per_size{};
    };
    
    explicit TXSlabCache(const CacheConfig& config = CacheConfig{});
    ~TXSlabCache();
    
    // ==================== 缓存接口 ====================
    
    /**
     * @brief 获取slab（优先使用缓存）
     */
    std::unique_ptr<TXSlab> getSlab(size_t object_size);
    
    /**
     * @brief 归还slab到缓存
     */
    bool returnSlab(std::unique_ptr<TXSlab> slab);
    
    /**
     * @brief 批量获取slab
     */
    std::vector<std::unique_ptr<TXSlab>> getSlabBatch(size_t object_size, size_t count);
    
    // ==================== 预分配和预热 ====================
    
    /**
     * @brief 预分配高频尺寸的slab
     */
    void preallocateHighFrequencySlabs();
    
    /**
     * @brief 预热缓存
     */
    void warmupCache();
    
    /**
     * @brief 设置高频尺寸
     */
    void setHighFrequencySizes(const std::vector<size_t>& sizes);
    
    // ==================== 缓存管理 ====================
    
    /**
     * @brief 清理过期缓存
     */
    size_t cleanupExpiredCache();
    
    /**
     * @brief 压缩缓存
     */
    size_t compactCache();
    
    /**
     * @brief 清空所有缓存
     */
    void clearCache();
    
    // ==================== 统计和监控 ====================
    
    /**
     * @brief 获取缓存统计
     */
    CacheStats getStats() const;
    
    /**
     * @brief 生成缓存报告
     */
    std::string generateCacheReport() const;
    
    /**
     * @brief 获取缓存命中率
     */
    double getHitRatio() const;
    
    // ==================== 配置管理 ====================
    
    /**
     * @brief 更新配置
     */
    void updateConfig(const CacheConfig& config);
    
    /**
     * @brief 获取配置
     */
    const CacheConfig& getConfig() const { return config_; }

private:
    // ==================== 内部数据结构 ====================
    
    /**
     * @brief 缓存条目
     */
    struct CacheEntry {
        std::unique_ptr<TXSlab> slab;
        std::chrono::steady_clock::time_point timestamp;
        size_t access_count = 0;
        
        CacheEntry(std::unique_ptr<TXSlab> s) 
            : slab(std::move(s)), timestamp(std::chrono::steady_clock::now()) {}
    };
    
    // 按对象大小分组的缓存
    std::map<size_t, std::vector<CacheEntry>> cache_;
    
    // 高频尺寸列表
    std::vector<size_t> high_frequency_sizes_;
    
    // 配置
    CacheConfig config_;
    
    // 统计信息
    mutable CacheStats stats_;
    mutable std::mutex stats_mutex_;
    
    // 线程安全
    mutable std::mutex cache_mutex_;
    
    // ==================== 内部方法 ====================
    
    /**
     * @brief 创建新slab
     */
    std::unique_ptr<TXSlab> createNewSlab(size_t object_size);
    
    /**
     * @brief 检查缓存条目是否过期
     */
    bool isExpired(const CacheEntry& entry) const;
    
    /**
     * @brief 更新统计信息
     */
    void updateStats(size_t object_size, bool cache_hit) const;
    
    /**
     * @brief 获取对象大小索引
     */
    size_t getSizeIndex(size_t object_size) const;
    
    /**
     * @brief 清理指定大小的过期缓存
     */
    size_t cleanupExpiredCacheForSize(size_t object_size);
};

/**
 * @brief 批量分配专用通道
 */
class TXBatchAllocator {
public:
    /**
     * @brief 批量分配配置
     */
    struct BatchConfig {
        size_t batch_size_threshold = 10;          // 批量分配阈值
        size_t max_batch_size = 1000;              // 最大批量大小
        bool enable_batch_optimization = true;     // 启用批量优化
        bool enable_parallel_allocation = false;   // 启用并行分配
    };
    
    explicit TXBatchAllocator(TXSlabCache& cache, const BatchConfig& config = BatchConfig{});
    ~TXBatchAllocator();
    
    /**
     * @brief 批量分配对象
     */
    std::vector<void*> allocateBatch(size_t object_size, size_t count);
    
    /**
     * @brief 批量释放对象
     */
    size_t deallocateBatch(const std::vector<void*>& ptrs);
    
    /**
     * @brief 混合大小批量分配
     */
    std::vector<void*> allocateMixedBatch(const std::vector<size_t>& sizes);
    
    /**
     * @brief 获取批量分配统计
     */
    struct BatchStats {
        size_t total_batch_allocations = 0;
        size_t total_objects_allocated = 0;
        size_t parallel_allocations = 0;
        double avg_batch_size = 0.0;
        double batch_efficiency = 0.0;
    };
    
    BatchStats getBatchStats() const;

private:
    TXSlabCache& cache_;
    BatchConfig config_;
    mutable BatchStats stats_;
    mutable std::mutex stats_mutex_;
    
    /**
     * @brief 并行分配实现
     */
    std::vector<void*> parallelAllocate(size_t object_size, size_t count);
    
    /**
     * @brief 串行分配实现
     */
    std::vector<void*> serialAllocate(size_t object_size, size_t count);
};

/**
 * @brief 无锁小对象分配路径
 */
class TXLockFreeAllocator {
public:
    /**
     * @brief 无锁分配器配置
     */
    struct LockFreeConfig {
        size_t thread_cache_size = 64;             // 线程缓存大小
        size_t max_object_size = 512;              // 最大无锁对象大小
        bool enable_thread_cache = true;          // 启用线程缓存
        bool enable_lock_free_path = true;        // 启用无锁路径
    };
    
    explicit TXLockFreeAllocator(TXSlabCache& cache, const LockFreeConfig& config = LockFreeConfig{});
    ~TXLockFreeAllocator();
    
    /**
     * @brief 无锁分配（仅支持小对象）
     */
    void* allocate(size_t size);
    
    /**
     * @brief 无锁释放
     */
    bool deallocate(void* ptr, size_t size);
    
    /**
     * @brief 检查是否支持无锁分配
     */
    bool canAllocateLockFree(size_t size) const;
    
    /**
     * @brief 刷新线程缓存
     */
    void flushThreadCache();
    
    /**
     * @brief 获取无锁分配统计
     */
    struct LockFreeStats {
        size_t lock_free_allocations = 0;
        size_t fallback_allocations = 0;
        size_t thread_cache_hits = 0;
        size_t thread_cache_misses = 0;
        double lock_free_ratio = 0.0;
    };
    
    LockFreeStats getLockFreeStats() const;

private:
    TXSlabCache& cache_;
    LockFreeConfig config_;
    
    // 线程本地存储
    thread_local static std::unique_ptr<TXSlabTLSCache> tls_cache_;
    
    mutable LockFreeStats stats_;
    mutable std::mutex stats_mutex_;
    
    /**
     * @brief 获取或创建线程缓存
     */
    TXSlabTLSCache* getOrCreateTLSCache();
    
    /**
     * @brief 回退到有锁分配
     */
    void* fallbackAllocate(size_t size);
};

} // namespace TinaXlsx
