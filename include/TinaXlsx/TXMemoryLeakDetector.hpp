//
// @file TXMemoryLeakDetector.hpp
// @brief 内存泄漏检测器 - 自动检测和修复内存泄漏
//

#pragma once

#include <memory>
#include <unordered_map>
#include <mutex>
#include <atomic>
#include <chrono>
#include <functional>

namespace TinaXlsx {

/**
 * @brief 🚀 内存泄漏检测器
 * 
 * 功能：
 * - 跟踪内存分配和释放
 * - 检测潜在的内存泄漏
 * - 自动清理机制
 * - 性能统计
 */
class TXMemoryLeakDetector {
public:
    /**
     * @brief 内存分配信息
     */
    struct AllocationInfo {
        size_t size;
        std::chrono::steady_clock::time_point timestamp;
        const char* file;
        int line;
        const char* function;
    };
    
    /**
     * @brief 检测配置
     */
    struct DetectorConfig {
        bool enableTracking = true;          // 是否启用跟踪
        bool enableAutoCleanup = true;       // 是否启用自动清理
        size_t maxAllocations = 100000;      // 最大跟踪分配数
        std::chrono::seconds cleanupInterval{60}; // 清理间隔
        size_t leakThreshold = 1024 * 1024;  // 泄漏阈值（字节）
    };
    
    static TXMemoryLeakDetector& instance();
    
    /**
     * @brief 设置配置
     */
    void setConfig(const DetectorConfig& config);
    
    /**
     * @brief 记录内存分配
     */
    void recordAllocation(void* ptr, size_t size, const char* file = nullptr, 
                         int line = 0, const char* function = nullptr);
    
    /**
     * @brief 记录内存释放
     */
    void recordDeallocation(void* ptr);
    
    /**
     * @brief 检测内存泄漏
     */
    struct LeakReport {
        size_t totalLeakedBytes = 0;
        size_t leakedAllocations = 0;
        std::vector<std::pair<void*, AllocationInfo>> leaks;
    };
    
    LeakReport detectLeaks() const;
    
    /**
     * @brief 强制清理所有跟踪的内存
     */
    void forceCleanup();
    
    /**
     * @brief 获取内存统计
     */
    struct MemoryStats {
        size_t currentAllocations = 0;
        size_t currentBytes = 0;
        size_t totalAllocations = 0;
        size_t totalDeallocations = 0;
        size_t totalBytes = 0;
        size_t peakAllocations = 0;
        size_t peakBytes = 0;
    };
    
    MemoryStats getStats() const;
    
    /**
     * @brief 重置统计信息
     */
    void reset();
    
    /**
     * @brief 启动自动清理线程
     */
    void startAutoCleanup();
    
    /**
     * @brief 停止自动清理线程
     */
    void stopAutoCleanup();

private:
    TXMemoryLeakDetector() = default;
    ~TXMemoryLeakDetector();
    
    DetectorConfig config_;
    mutable std::mutex mutex_;
    
    std::unordered_map<void*, AllocationInfo> allocations_;
    
    std::atomic<size_t> currentAllocations_{0};
    std::atomic<size_t> currentBytes_{0};
    std::atomic<size_t> totalAllocations_{0};
    std::atomic<size_t> totalDeallocations_{0};
    std::atomic<size_t> totalBytes_{0};
    std::atomic<size_t> peakAllocations_{0};
    std::atomic<size_t> peakBytes_{0};
    
    std::unique_ptr<std::thread> cleanupThread_;
    std::atomic<bool> stopCleanup_{false};
    
    void cleanupOldAllocations();
    void updatePeakStats();
};

/**
 * @brief 🚀 RAII内存跟踪器
 * 
 * 自动跟踪作用域内的内存分配
 */
class TXScopedMemoryTracker {
public:
    explicit TXScopedMemoryTracker(const char* name = "ScopedTracker");
    ~TXScopedMemoryTracker();
    
    /**
     * @brief 获取作用域内的内存统计
     */
    TXMemoryLeakDetector::MemoryStats getScopeStats() const;
    
    /**
     * @brief 检查作用域内是否有泄漏
     */
    bool hasLeaks() const;

private:
    const char* name_;
    TXMemoryLeakDetector::MemoryStats initialStats_;
    std::chrono::steady_clock::time_point startTime_;
};

/**
 * @brief 🚀 智能内存管理器
 * 
 * 结合内存池和泄漏检测的智能管理器
 */
class TXSmartMemoryManager {
public:
    static TXSmartMemoryManager& instance();
    
    /**
     * @brief 智能分配内存
     * @param size 大小
     * @param alignment 对齐要求
     * @return 内存指针
     */
    void* allocate(size_t size, size_t alignment = sizeof(void*));
    
    /**
     * @brief 智能释放内存
     * @param ptr 内存指针
     */
    void deallocate(void* ptr);
    
    /**
     * @brief 类型化分配
     */
    template<typename T, typename... Args>
    T* create(Args&&... args) {
        void* ptr = allocate(sizeof(T), alignof(T));
        if (!ptr) return nullptr;
        
        try {
            return new(ptr) T(std::forward<Args>(args)...);
        } catch (...) {
            deallocate(ptr);
            throw;
        }
    }
    
    /**
     * @brief 类型化释放
     */
    template<typename T>
    void destroy(T* ptr) {
        if (!ptr) return;
        
        ptr->~T();
        deallocate(ptr);
    }
    
    /**
     * @brief 执行内存健康检查
     */
    struct HealthReport {
        bool hasLeaks = false;
        size_t leakedBytes = 0;
        size_t fragmentationLevel = 0;
        double memoryEfficiency = 1.0;
        std::vector<std::string> recommendations;
    };
    
    HealthReport performHealthCheck();
    
    /**
     * @brief 优化内存使用
     */
    void optimize();
    
    /**
     * @brief 紧急清理（释放所有可能的内存）
     */
    void emergencyCleanup();

private:
    TXSmartMemoryManager();
    ~TXSmartMemoryManager();
    
    std::unique_ptr<class TXMemoryPool> generalPool_;
    mutable std::mutex mutex_;
    
    // 智能策略
    bool shouldUsePool(size_t size) const;
    void* allocateFromPool(size_t size);
    void* allocateFromSystem(size_t size);
    void deallocateToPool(void* ptr);
    void deallocateToSystem(void* ptr);
    
    // 统计和优化
    std::atomic<size_t> poolAllocations_{0};
    std::atomic<size_t> systemAllocations_{0};
    std::atomic<size_t> totalFragmentation_{0};
    
    void updateFragmentationStats();
    void suggestOptimizations(HealthReport& report);
};

} // namespace TinaXlsx

// 便利宏
#ifdef TINAXLSX_ENABLE_MEMORY_TRACKING
    #define TX_ALLOC(size) TinaXlsx::TXMemoryLeakDetector::instance().recordAllocation( \
        malloc(size), size, __FILE__, __LINE__, __FUNCTION__); malloc(size)
    #define TX_FREE(ptr) TinaXlsx::TXMemoryLeakDetector::instance().recordDeallocation(ptr); free(ptr)
    #define TX_SCOPE_TRACKER(name) TinaXlsx::TXScopedMemoryTracker _tracker(name)
#else
    #define TX_ALLOC(size) malloc(size)
    #define TX_FREE(ptr) free(ptr)
    #define TX_SCOPE_TRACKER(name)
#endif
