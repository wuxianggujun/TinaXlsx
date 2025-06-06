//
// @file TXChunkAllocator.hpp
// @brief 分块内存管理器 - 4GB限制和高性能分配
//

#pragma once

#include <memory>
#include <vector>
#include <atomic>
#include <mutex>
#include <array>
#include <cstdint>
#include <chrono>
#include <unordered_map>
#include <unordered_set>
#include <queue>

namespace TinaXlsx {

/**
 * @brief 内存块配置
 */
struct ChunkConfig {
    static constexpr size_t DEFAULT_CHUNK_SIZE = 64 * 1024 * 1024;  // 64MB
    static constexpr size_t MAX_CHUNKS = 64;                        // 最大64块
    static constexpr size_t MAX_TOTAL_MEMORY = 4ULL * 1024 * 1024 * 1024; // 4GB
    static constexpr size_t ALIGNMENT = 32;                         // 32字节对齐
    static constexpr size_t MIN_ALLOCATION = 16;                    // 最小分配16字节

    // 多级块大小策略
    static constexpr size_t SMALL_CHUNK_SIZE = 1024 * 1024;         // 1MB小块
    static constexpr size_t MEDIUM_CHUNK_SIZE = 16 * 1024 * 1024;   // 16MB中块
    static constexpr size_t LARGE_CHUNK_SIZE = 64 * 1024 * 1024;    // 64MB大块

    static constexpr size_t SMALL_ALLOCATION_THRESHOLD = 64 * 1024;  // 64KB以下用小块
    static constexpr size_t MEDIUM_ALLOCATION_THRESHOLD = 4 * 1024 * 1024; // 4MB以下用中块
};

/**
 * @brief 内存块
 */
class TXMemoryChunk {
public:
    explicit TXMemoryChunk(size_t size = ChunkConfig::DEFAULT_CHUNK_SIZE);
    ~TXMemoryChunk();
    
    /**
     * @brief 分配内存
     * @param size 请求的大小
     * @param alignment 对齐要求
     * @return 分配的内存指针，失败返回nullptr
     */
    void* allocate(size_t size, size_t alignment = ChunkConfig::ALIGNMENT);
    
    /**
     * @brief 检查是否可以分配指定大小
     */
    bool canAllocate(size_t size, size_t alignment = ChunkConfig::ALIGNMENT) const;
    
    /**
     * @brief 获取已使用的内存大小
     */
    size_t getUsedSize() const { return used_size_.load(); }
    
    /**
     * @brief 获取总大小
     */
    size_t getTotalSize() const { return total_size_; }
    
    /**
     * @brief 获取剩余大小
     */
    size_t getRemainingSize() const { return total_size_ - used_size_.load(); }
    
    /**
     * @brief 获取使用率
     */
    double getUsageRatio() const { 
        return static_cast<double>(used_size_.load()) / total_size_; 
    }
    
    /**
     * @brief 重置块（清空所有分配）
     */
    void reset();
    
    /**
     * @brief 检查指针是否属于此块
     */
    bool contains(const void* ptr) const;

private:
    std::unique_ptr<char[]> data_;
    size_t total_size_;
    std::atomic<size_t> used_size_{0};
    mutable std::mutex mutex_;
    
    /**
     * @brief 计算对齐后的大小
     */
    static size_t alignSize(size_t size, size_t alignment);
    
    /**
     * @brief 计算对齐后的指针
     */
    static void* alignPointer(void* ptr, size_t alignment);
};

/**
 * @brief 🚀 内存池块信息
 */
struct PoolBlock {
    void* ptr;              // 内存指针
    size_t size;            // 块大小
    bool is_free;           // 是否空闲
    size_t chunk_index;     // 所属chunk索引

    PoolBlock(void* p, size_t s, size_t chunk_idx)
        : ptr(p), size(s), is_free(true), chunk_index(chunk_idx) {}
};

/**
 * @brief 🚀 分块内存分配器 - 支持内存池和单独释放
 */
class TXChunkAllocator {
public:
    /**
     * @brief 分配统计信息
     */
    struct AllocationStats {
        size_t total_allocated = 0;        // 总分配内存
        size_t total_chunks = 0;           // 总块数
        size_t active_chunks = 0;          // 活跃块数
        size_t peak_memory = 0;            // 峰值内存
        size_t allocation_count = 0;       // 分配次数
        size_t failed_allocations = 0;    // 失败分配次数
        double average_chunk_usage = 0.0;  // 平均块使用率
        double memory_efficiency = 0.0;    // 内存效率
        
        std::chrono::steady_clock::time_point start_time;
        std::chrono::steady_clock::time_point last_allocation_time;
    };
    
    TXChunkAllocator();
    ~TXChunkAllocator();
    
    // ==================== 内存分配接口 ====================
    
    /**
     * @brief 分配内存
     * @param size 请求的大小
     * @param alignment 对齐要求
     * @return 分配的内存指针，失败返回nullptr
     */
    void* allocate(size_t size, size_t alignment = ChunkConfig::ALIGNMENT);
    
    /**
     * @brief 分配对齐内存（模板版本）
     */
    template<typename T>
    T* allocate(size_t count = 1);
    
    /**
     * @brief 批量分配
     */
    std::vector<void*> allocateBatch(const std::vector<size_t>& sizes);

    /**
     * @brief 🚀 释放单个内存块 - 支持内存池重用
     * @param ptr 要释放的内存指针
     * @return 是否成功释放
     */
    bool deallocate(void* ptr);

    /**
     * @brief 释放所有内存
     */
    void deallocateAll();
    
    /**
     * @brief 压缩内存（移除空块）
     */
    void compact();
    
    // ==================== 内存监控 ====================
    
    /**
     * @brief 获取总内存使用量
     */
    size_t getTotalMemoryUsage() const;
    
    /**
     * @brief 获取峰值内存使用量
     */
    size_t getPeakMemoryUsage() const { return stats_.peak_memory; }
    
    /**
     * @brief 检查内存限制
     */
    bool checkMemoryLimit(size_t additional_size) const;
    
    /**
     * @brief 获取内存使用率
     */
    double getMemoryUsageRatio() const;
    
    /**
     * @brief 获取分配统计信息
     */
    AllocationStats getStats() const;
    
    /**
     * @brief 重置统计信息
     */
    void resetStats();
    
    // ==================== 配置管理 ====================
    
    /**
     * @brief 设置块大小
     */
    void setChunkSize(size_t size) { chunk_size_ = size; }
    
    /**
     * @brief 获取块大小
     */
    size_t getChunkSize() const { return chunk_size_; }
    
    /**
     * @brief 设置内存限制
     */
    void setMemoryLimit(size_t limit) { memory_limit_ = limit; }
    
    /**
     * @brief 获取内存限制
     */
    size_t getMemoryLimit() const { return memory_limit_; }
    
    /**
     * @brief 启用/禁用自动压缩
     */
    void enableAutoCompact(bool enable) { auto_compact_enabled_ = enable; }
    
    /**
     * @brief 检查自动压缩是否启用
     */
    bool isAutoCompactEnabled() const { return auto_compact_enabled_; }
    
    // ==================== 调试和诊断 ====================
    
    /**
     * @brief 获取块信息
     */
    struct ChunkInfo {
        size_t index;
        size_t total_size;
        size_t used_size;
        double usage_ratio;
        bool is_active;
    };
    
    std::vector<ChunkInfo> getChunkInfos() const;
    
    /**
     * @brief 打印内存使用报告
     */
    std::string generateMemoryReport() const;
    
    /**
     * @brief 验证内存完整性
     */
    bool validateMemoryIntegrity() const;

private:
    // ==================== 内部数据结构 ====================
    
    std::array<std::unique_ptr<TXMemoryChunk>, ChunkConfig::MAX_CHUNKS> chunks_;
    std::atomic<size_t> chunk_count_{0};
    std::atomic<size_t> total_allocated_{0};
    
    // 配置
    size_t chunk_size_ = ChunkConfig::DEFAULT_CHUNK_SIZE;
    size_t memory_limit_ = ChunkConfig::MAX_TOTAL_MEMORY;
    bool auto_compact_enabled_ = true;
    
    // 统计信息
    mutable AllocationStats stats_;
    mutable std::mutex stats_mutex_;
    
    // 线程安全
    mutable std::mutex chunks_mutex_;

    // 🚀 内存池管理
    std::unordered_map<void*, std::unique_ptr<PoolBlock>> allocated_blocks_; // 已分配块映射
    std::unordered_map<size_t, std::queue<std::unique_ptr<PoolBlock>>> free_pools_; // 按大小分类的空闲池
    mutable std::mutex pool_mutex_; // 内存池互斥锁
    
    // ==================== 内部方法 ====================
    
    /**
     * @brief 创建新块
     */
    TXMemoryChunk* createNewChunk(size_t requested_size);

    /**
     * @brief 查找可用块
     */
    TXMemoryChunk* findAvailableChunk(size_t size, size_t alignment);

    /**
     * @brief 选择合适的块大小
     */
    size_t selectOptimalChunkSize(size_t requested_size) const;
    
    /**
     * @brief 更新统计信息
     */
    void updateStats(size_t allocated_size, bool success);
    
    /**
     * @brief 检查是否需要压缩
     */
    bool shouldCompact() const;
    
    /**
     * @brief 计算内存效率
     */
    double calculateMemoryEfficiency() const;

    // 🚀 内存池相关方法

    /**
     * @brief 从内存池获取空闲块
     */
    PoolBlock* getFromPool(size_t size);

    /**
     * @brief 将块返回到内存池
     */
    void returnToPool(std::unique_ptr<PoolBlock> block);

    /**
     * @brief 创建新的池块
     */
    std::unique_ptr<PoolBlock> createPoolBlock(size_t size, size_t chunk_index);

    /**
     * @brief 查找最佳匹配的池块大小
     */
    size_t findBestPoolSize(size_t requested_size) const;

    /**
     * @brief 清理内存池
     */
    void cleanupPools();
};

// ==================== 模板实现 ====================

template<typename T>
T* TXChunkAllocator::allocate(size_t count) {
    size_t size = sizeof(T) * count;
    size_t alignment = std::max(alignof(T), ChunkConfig::ALIGNMENT);
    return static_cast<T*>(allocate(size, alignment));
}

/**
 * @brief RAII内存管理器
 */
template<typename T>
class TXChunkPtr {
public:
    TXChunkPtr(TXChunkAllocator& allocator, size_t count = 1)
        : allocator_(allocator), ptr_(allocator.allocate<T>(count)), count_(count) {}
    
    ~TXChunkPtr() {
        // 注意：TXChunkAllocator不支持单独释放，只能整体释放
    }
    
    T* get() const { return ptr_; }
    T& operator*() const { return *ptr_; }
    T* operator->() const { return ptr_; }
    T& operator[](size_t index) const { return ptr_[index]; }
    
    explicit operator bool() const { return ptr_ != nullptr; }
    
    size_t size() const { return count_; }

private:
    TXChunkAllocator& allocator_;
    T* ptr_;
    size_t count_;
};

} // namespace TinaXlsx
