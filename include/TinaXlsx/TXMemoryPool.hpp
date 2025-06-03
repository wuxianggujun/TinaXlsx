//
// @file TXMemoryPool.hpp
// @brief 高性能内存池管理器 - 解决内存泄漏和提升分配性能
//

#pragma once

#include <memory>
#include <vector>
#include <mutex>
#include <map>
#include <atomic>
#include <cstddef>
#include <type_traits>

namespace TinaXlsx {

/**
 * @brief 🚀 高性能内存池
 * 
 * 特点：
 * - 固定大小块分配，避免碎片
 * - 线程安全
 * - 自动回收，防止内存泄漏
 * - RAII管理
 */
class TXMemoryPool {
public:
    /**
     * @brief 内存池配置
     */
    struct PoolConfig {
        size_t blockSize = 64;           // 块大小（字节）
        size_t blocksPerChunk = 1024;    // 每个chunk的块数
        size_t maxChunks = 100;          // 最大chunk数量
        bool threadSafe = true;          // 是否线程安全
        bool autoGrow = true;            // 是否自动增长
    };
    
    explicit TXMemoryPool(const PoolConfig& config = PoolConfig{});
    ~TXMemoryPool();
    
    // 禁止拷贝和移动
    TXMemoryPool(const TXMemoryPool&) = delete;
    TXMemoryPool& operator=(const TXMemoryPool&) = delete;
    
    /**
     * @brief 分配内存块
     * @param size 请求的大小（必须 <= blockSize）
     * @return 内存指针，失败返回nullptr
     */
    void* allocate(size_t size);
    
    /**
     * @brief 释放内存块
     * @param ptr 要释放的指针
     */
    void deallocate(void* ptr);
    
    /**
     * @brief 类型化分配
     */
    template<typename T>
    T* allocate() {
        static_assert(sizeof(T) <= 64, "Type too large for default pool");
        return static_cast<T*>(allocate(sizeof(T)));
    }
    
    /**
     * @brief 类型化释放
     */
    template<typename T>
    void deallocate(T* ptr) {
        deallocate(static_cast<void*>(ptr));
    }
    
    /**
     * @brief 获取统计信息
     */
    struct PoolStats {
        size_t totalAllocated = 0;      // 总分配字节数
        size_t totalDeallocated = 0;    // 总释放字节数
        size_t currentUsage = 0;        // 当前使用量
        size_t peakUsage = 0;           // 峰值使用量
        size_t totalChunks = 0;         // 总chunk数
        size_t freeBlocks = 0;          // 空闲块数
    };
    
    PoolStats getStats() const;
    
    /**
     * @brief 清空池（释放所有内存）
     */
    void clear();
    
    /**
     * @brief 收缩池（释放未使用的chunk）
     */
    void shrink();

    /**
     * @brief 检查指针是否来自此内存池
     */
    bool isFromPool(void* ptr) const;

private:
    struct Block {
        Block* next = nullptr;
    };
    
    struct Chunk {
        std::unique_ptr<uint8_t[]> memory;
        Block* freeList = nullptr;
        size_t freeCount = 0;
        
        Chunk(size_t blockSize, size_t blockCount);
        ~Chunk() = default;
    };
    
    PoolConfig config_;
    std::vector<std::unique_ptr<Chunk>> chunks_;
    Block* globalFreeList_ = nullptr;
    
    mutable std::mutex mutex_;
    std::atomic<size_t> totalAllocated_{0};
    std::atomic<size_t> totalDeallocated_{0};
    std::atomic<size_t> currentUsage_{0};
    std::atomic<size_t> peakUsage_{0};
    
    Chunk* createChunk();
    Block* allocateFromChunk(Chunk* chunk);
    void deallocateToChunk(void* ptr);
};

/**
 * @brief 🚀 字符串专用内存池
 * 
 * 针对Excel文件中大量字符串优化
 */
class TXStringPool {
public:
    struct StringPoolConfig {
        size_t smallStringSize = 32;     // 小字符串大小
        size_t mediumStringSize = 128;   // 中等字符串大小
        size_t largeStringSize = 512;    // 大字符串大小
        size_t poolsPerSize = 10;        // 每种大小的池数量
    };
    
    explicit TXStringPool(const StringPoolConfig& config = StringPoolConfig{});
    ~TXStringPool();
    
    /**
     * @brief 分配字符串内存
     * @param size 字符串长度
     * @return 内存指针
     */
    char* allocateString(size_t size);
    
    /**
     * @brief 释放字符串内存
     * @param ptr 字符串指针
     * @param size 字符串长度
     */
    void deallocateString(char* ptr, size_t size);
    
    /**
     * @brief 创建字符串（自动管理内存）
     * @param str 源字符串
     * @return 池管理的字符串
     */
    std::string_view createString(std::string_view str);
    
    /**
     * @brief 清空所有字符串
     */
    void clear();
    
    /**
     * @brief 获取统计信息
     */
    struct StringStats {
        size_t totalStrings = 0;
        size_t totalBytes = 0;
        size_t smallStrings = 0;
        size_t mediumStrings = 0;
        size_t largeStrings = 0;
    };
    
    StringStats getStats() const;

private:
    StringPoolConfig config_;
    std::unique_ptr<TXMemoryPool> smallPool_;
    std::unique_ptr<TXMemoryPool> mediumPool_;
    std::unique_ptr<TXMemoryPool> largePool_;
    
    std::vector<std::unique_ptr<char[]>> largeAllocations_;
    mutable std::mutex mutex_;
    
    std::atomic<size_t> totalStrings_{0};
    std::atomic<size_t> totalBytes_{0};
    std::atomic<size_t> smallStrings_{0};
    std::atomic<size_t> mediumStrings_{0};
    std::atomic<size_t> largeStrings_{0};
    
    TXMemoryPool* selectPool(size_t size);
};

/**
 * @brief 🚀 全局内存管理器
 * 
 * 管理所有内存池，提供统一接口
 */
class TXMemoryManager {
public:
    static TXMemoryManager& instance();
    
    /**
     * @brief 获取通用内存池
     */
    TXMemoryPool& getGeneralPool();
    
    /**
     * @brief 获取字符串内存池
     */
    TXStringPool& getStringPool();
    
    /**
     * @brief 获取指定大小的内存池
     */
    TXMemoryPool& getPool(size_t blockSize);
    
    /**
     * @brief 清空所有池
     */
    void clearAll();
    
    /**
     * @brief 收缩所有池
     */
    void shrinkAll();
    
    /**
     * @brief 获取全局统计信息
     */
    struct GlobalStats {
        TXMemoryPool::PoolStats generalPool;
        TXStringPool::StringStats stringPool;
        size_t totalPools = 0;
        size_t totalMemoryUsage = 0;
    };
    
    GlobalStats getGlobalStats() const;

private:
    TXMemoryManager();
    ~TXMemoryManager();
    
    std::unique_ptr<TXMemoryPool> generalPool_;
    std::unique_ptr<TXStringPool> stringPool_;
    std::map<size_t, std::unique_ptr<TXMemoryPool>> sizedPools_;
    mutable std::mutex mutex_;
};

/**
 * @brief RAII内存管理器
 */
template<typename T>
class PoolAllocator {
public:
    using value_type = T;
    
    explicit PoolAllocator(TXMemoryPool* pool = nullptr) 
        : pool_(pool ? pool : &TXMemoryManager::instance().getGeneralPool()) {}
    
    template<typename U>
    PoolAllocator(const PoolAllocator<U>& other) : pool_(other.pool_) {}
    
    T* allocate(size_t n) {
        if (n == 1 && sizeof(T) <= 64) {
            return pool_->template allocate<T>();
        }
        return static_cast<T*>(std::malloc(n * sizeof(T)));
    }
    
    void deallocate(T* ptr, size_t n) {
        if (n == 1 && sizeof(T) <= 64) {
            pool_->deallocate(ptr);
        } else {
            std::free(ptr);
        }
    }
    
    template<typename U>
    bool operator==(const PoolAllocator<U>& other) const {
        return pool_ == other.pool_;
    }
    
    template<typename U>
    bool operator!=(const PoolAllocator<U>& other) const {
        return !(*this == other);
    }

private:
    template<typename U> friend class PoolAllocator;
    TXMemoryPool* pool_;
};

} // namespace TinaXlsx
