//
// @file TXSlabAllocator.hpp
// @brief Slab分配器 - 专门处理小对象分配，解决内存碎片问题
//

#pragma once

#include <memory>
#include <vector>
#include <array>
#include <atomic>
#include <mutex>
#include <bitset>
#include <cstdint>

namespace TinaXlsx {

/**
 * @brief Slab配置
 */
struct SlabConfig {
    // 第三阶段：生产环境优化配置
    static constexpr std::array<size_t, 10> OBJECT_SIZES = {
        16, 32, 64, 128,        // 微小对象 (<128B)
        256, 512, 1024,         // 中等对象 (128B-1KB)
        2048, 4096, 8192        // 大对象 (>1KB)
    };

    // 生产环境分级策略配置表 - 确保高效率
    struct OptimalConfig {
        size_t slab_size;
        size_t objects_per_slab;
        double target_efficiency;
    };

    static constexpr std::array<OptimalConfig, 10> OPTIMAL_CONFIGS = {{
        {2048,  128, 1.0},   // 16B:  2KB存128对象 → 100%效率
        {2048,  64,  1.0},   // 32B:  2KB存64对象  → 100%效率
        {2048,  32,  1.0},   // 64B:  2KB存32对象  → 100%效率
        {2048,  16,  1.0},   // 128B: 2KB存16对象  → 100%效率
        {8192,  32,  1.0},   // 256B: 8KB存32对象  → 100%效率
        {8192,  16,  1.0},   // 512B: 8KB存16对象  → 100%效率
        {8192,  8,   1.0},   // 1KB:  8KB存8对象   → 100%效率
        {16384, 8,   1.0},   // 2KB:  16KB存8对象  → 100%效率
        {32768, 8,   1.0},   // 4KB:  32KB存8对象  → 100%效率
        {65536, 8,   1.0}    // 8KB:  64KB存8对象  → 100%效率
    }};

    // 获取优化配置
    static constexpr OptimalConfig getOptimalConfig(size_t size_index) {
        return OPTIMAL_CONFIGS[size_index];
    }

    static constexpr size_t MAX_SLABS_PER_SIZE = 64;    // 每种大小最多64个slab
    static constexpr size_t ALIGNMENT = 16;             // 16字节对齐

    // 高频缓存配置
    static constexpr size_t CACHE_SLABS_COUNT = 3;      // 保留3个高频slab作缓存
    static constexpr double FRAGMENTATION_THRESHOLD = 0.3; // 30%碎片率触发合并

    // 第三阶段：使用优化配置表的slab大小获取
    static constexpr size_t getSlabSize(size_t object_size) {
        for (size_t i = 0; i < OBJECT_SIZES.size(); ++i) {
            if (object_size <= OBJECT_SIZES[i]) {
                return OPTIMAL_CONFIGS[i].slab_size;
            }
        }
        return OPTIMAL_CONFIGS.back().slab_size; // 默认最大配置
    }

    // 根据大小索引获取slab大小
    static constexpr size_t getSlabSizeByIndex(size_t size_index) {
        return OPTIMAL_CONFIGS[size_index].slab_size;
    }

    // 获取对象分组类型
    enum class ObjectGroup {
        MICRO,   // <128B
        SMALL,   // 128B-1KB
        MEDIUM   // >1KB
    };

    static constexpr ObjectGroup getObjectGroup(size_t object_size) {
        if (object_size < 128) return ObjectGroup::MICRO;
        if (object_size <= 1024) return ObjectGroup::SMALL;
        return ObjectGroup::MEDIUM;
    }

    // 计算每个slab的对象容量
    static constexpr size_t getObjectsPerSlab(size_t object_size) {
        return getSlabSize(object_size) / object_size;
    }
};

/**
 * @brief 单个Slab
 */
class TXSlab {
public:
    explicit TXSlab(size_t object_size);
    ~TXSlab();
    
    /**
     * @brief 分配对象
     */
    void* allocate();
    
    /**
     * @brief 释放对象
     */
    bool deallocate(void* ptr);
    
    /**
     * @brief 检查是否可以分配
     */
    bool canAllocate() const { return free_count_.load() > 0; }
    
    /**
     * @brief 检查是否为空
     */
    bool isEmpty() const { return free_count_.load() == max_objects_; }
    
    /**
     * @brief 检查是否已满
     */
    bool isFull() const { return free_count_.load() == 0; }
    
    /**
     * @brief 获取使用率
     */
    double getUsageRatio() const { 
        return 1.0 - static_cast<double>(free_count_.load()) / max_objects_; 
    }
    
    /**
     * @brief 获取对象大小
     */
    size_t getObjectSize() const { return object_size_; }
    
    /**
     * @brief 获取最大对象数
     */
    size_t getMaxObjects() const { return max_objects_; }
    
    /**
     * @brief 获取空闲对象数
     */
    size_t getFreeCount() const { return free_count_.load(); }
    
    /**
     * @brief 检查指针是否属于此slab
     */
    bool contains(const void* ptr) const;

private:
    std::unique_ptr<char[]> data_;
    size_t object_size_;
    size_t max_objects_;
    std::atomic<size_t> free_count_;
    
    // 空闲对象链表
    struct FreeNode {
        FreeNode* next;
    };
    std::atomic<FreeNode*> free_list_;
    
    mutable std::mutex mutex_;
    
    /**
     * @brief 初始化空闲链表
     */
    void initializeFreeList();
};

/**
 * @brief Slab分配器
 */
class TXSlabAllocator {
public:
    /**
     * @brief 分配统计
     */
    struct SlabStats {
        size_t total_slabs = 0;
        size_t active_slabs = 0;
        size_t total_objects = 0;
        size_t allocated_objects = 0;
        size_t total_memory = 0;
        size_t used_memory = 0;
        double memory_efficiency = 0.0;
        double fragmentation_ratio = 0.0;
        
        std::array<size_t, SlabConfig::OBJECT_SIZES.size()> slabs_per_size{};
        std::array<size_t, SlabConfig::OBJECT_SIZES.size()> objects_per_size{};
        std::array<double, SlabConfig::OBJECT_SIZES.size()> efficiency_per_size{};
    };
    
    TXSlabAllocator();
    ~TXSlabAllocator();
    
    // ==================== 分配接口 ====================
    
    /**
     * @brief 分配小对象
     * @param size 请求大小（必须 <= 2048字节）
     * @return 分配的指针，失败返回nullptr
     */
    void* allocate(size_t size);
    
    /**
     * @brief 释放小对象
     * @param ptr 要释放的指针
     * @return 是否成功释放
     */
    bool deallocate(void* ptr);
    
    /**
     * @brief 批量分配
     */
    std::vector<void*> allocateBatch(const std::vector<size_t>& sizes);
    
    /**
     * @brief 检查是否可以处理指定大小
     */
    static bool canHandle(size_t size) {
        return size > 0 && size <= SlabConfig::OBJECT_SIZES.back();
    }
    
    // ==================== 内存管理 ====================
    
    /**
     * @brief 智能压缩内存（第二阶段优化）
     */
    size_t smartCompact();

    /**
     * @brief 传统压缩内存（移除空slab）
     */
    size_t compact();

    /**
     * @brief 清空所有内存
     */
    void clear();

    /**
     * @brief 自动回收管理
     */
    void enableAutoReclaim(bool enable) { auto_reclaim_enabled_ = enable; }
    bool isAutoReclaimEnabled() const { return auto_reclaim_enabled_; }

    /**
     * @brief 检查并执行自动回收
     */
    size_t checkAndReclaim();
    
    /**
     * @brief 获取内存使用量
     */
    size_t getTotalMemoryUsage() const;
    
    /**
     * @brief 获取实际使用的内存
     */
    size_t getUsedMemorySize() const;
    
    // ==================== 统计信息 ====================
    
    /**
     * @brief 获取统计信息
     */
    SlabStats getStats() const;
    
    /**
     * @brief 生成详细报告
     */
    std::string generateReport() const;
    
    /**
     * @brief 获取碎片率
     */
    double getFragmentationRatio() const;

private:
    // ==================== 内部数据结构 ====================
    
    // 每种大小的slab列表
    std::array<std::vector<std::unique_ptr<TXSlab>>, SlabConfig::OBJECT_SIZES.size()> slabs_;
    
    // 线程安全
    mutable std::array<std::mutex, SlabConfig::OBJECT_SIZES.size()> size_mutexes_;
    mutable std::mutex global_mutex_;
    
    // 统计信息
    std::atomic<size_t> total_allocations_{0};
    std::atomic<size_t> total_deallocations_{0};
    std::atomic<size_t> failed_allocations_{0};

    // 第二阶段：智能回收配置
    bool auto_reclaim_enabled_ = true;
    std::chrono::steady_clock::time_point last_reclaim_time_;
    std::atomic<size_t> reclaim_counter_{0};
    
    // ==================== 内部方法 ====================
    
    /**
     * @brief 获取大小索引
     */
    static size_t getSizeIndex(size_t size);
    
    /**
     * @brief 获取实际分配大小
     */
    static size_t getActualSize(size_t size);
    
    /**
     * @brief 查找可用slab
     */
    TXSlab* findAvailableSlab(size_t size_index);
    
    /**
     * @brief 创建新slab
     */
    TXSlab* createNewSlab(size_t size_index);
    
    /**
     * @brief 查找包含指针的slab
     */
    TXSlab* findSlabContaining(void* ptr, size_t& size_index) const;
    
    /**
     * @brief 移除空slab
     */
    size_t removeEmptySlabs(size_t size_index);

    /**
     * @brief 智能回收：移除空slab但保留缓存
     */
    size_t smartRemoveEmptySlabs(size_t size_index);

    /**
     * @brief 检查碎片率并触发合并
     */
    bool shouldTriggerReclaim(size_t size_index) const;

    /**
     * @brief 立即释放空slab（第二阶段优化）
     */
    void immediateReleaseEmptySlabs();
};

/**
 * @brief 线程本地缓存（TLS优化）
 */
class TXSlabTLSCache {
public:
    static constexpr size_t CACHE_SIZE = 16; // 每种大小缓存16个对象
    
    TXSlabTLSCache(TXSlabAllocator& allocator);
    ~TXSlabTLSCache();
    
    /**
     * @brief 从缓存分配
     */
    void* allocate(size_t size);
    
    /**
     * @brief 释放到缓存
     */
    bool deallocate(void* ptr, size_t size);
    
    /**
     * @brief 刷新缓存
     */
    void flush();

private:
    TXSlabAllocator& allocator_;
    
    struct CacheEntry {
        std::array<void*, CACHE_SIZE> objects{};
        size_t count = 0;
    };
    
    std::array<CacheEntry, SlabConfig::OBJECT_SIZES.size()> cache_;
    
    /**
     * @brief 填充缓存
     */
    void fillCache(size_t size_index);
    
    /**
     * @brief 清空缓存
     */
    void drainCache(size_t size_index);
};

} // namespace TinaXlsx
