//
// @file TXUnifiedMemoryManager.hpp
// @brief 统一内存管理器 - 简化的双级分配系统
//

#pragma once

#include "TXSlabAllocator.hpp"
#include "TXChunkAllocator.hpp"
#include "TXSmartMemoryManager.hpp"
#include <memory>
#include <atomic>

namespace TinaXlsx {

/**
 * @brief 统一内存管理器 - 简化版
 * 
 * 整合了三个核心组件：
 * - TXSlabAllocator: 小对象分配（<=8KB）
 * - TXChunkAllocator: 大对象分配（>8KB）  
 * - TXSmartMemoryManager: 智能监控和管理
 */
class TXUnifiedMemoryManager {
public:
    /**
     * @brief 统一统计信息
     */
    struct UnifiedStats {
        // Slab分配器统计
        TXSlabAllocator::SlabStats slab_stats;
        
        // Chunk分配器统计
        TXChunkAllocator::AllocationStats chunk_stats;
        
        // 智能管理器统计
        TXSmartMemoryManager::MonitoringStats monitor_stats;
        
        // 综合指标
        size_t total_memory_usage = 0;
        size_t total_used_memory = 0;
        double overall_efficiency = 0.0;
        size_t small_allocations = 0;   // Slab处理
        size_t large_allocations = 0;   // Chunk处理
        
        // 性能指标
        double avg_allocation_time_us = 0.0;
        size_t allocations_per_second = 0;
    };
    
    /**
     * @brief 配置
     */
    struct Config {
        // Slab配置
        bool enable_slab_allocator = true;
        bool enable_auto_reclaim = true;
        
        // Chunk配置  
        size_t chunk_size = 64 * 1024 * 1024;      // 64MB
        size_t memory_limit = 4ULL * 1024 * 1024 * 1024; // 4GB
        
        // 监控配置
        bool enable_monitoring = true;
        size_t warning_threshold_mb = 3072;        // 3GB
        size_t critical_threshold_mb = 3584;       // 3.5GB
        size_t emergency_threshold_mb = 3840;      // 3.75GB
        
        // 分界线配置
        size_t slab_chunk_threshold = 8192;        // 8KB分界线
    };
    
    explicit TXUnifiedMemoryManager(const Config& config = Config{});
    ~TXUnifiedMemoryManager();
    
    // ==================== 统一分配接口 ====================
    
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
     * @brief 全面压缩
     */
    size_t compactAll();
    
    /**
     * @brief 智能清理
     */
    size_t smartCleanup();
    
    /**
     * @brief 清空所有内存
     */
    void clear();
    
    // ==================== 监控和统计 ====================
    
    /**
     * @brief 启动监控
     */
    void startMonitoring();
    
    /**
     * @brief 停止监控
     */
    void stopMonitoring();
    
    /**
     * @brief 获取统一统计信息
     */
    UnifiedStats getUnifiedStats() const;
    
    /**
     * @brief 生成综合报告
     */
    std::string generateComprehensiveReport() const;
    
    /**
     * @brief 获取内存使用情况
     */
    size_t getTotalMemoryUsage() const;
    size_t getUsedMemorySize() const;
    double getOverallEfficiency() const;
    
    // ==================== 配置管理 ====================
    
    /**
     * @brief 更新配置
     */
    void updateConfig(const Config& config);
    
    /**
     * @brief 获取配置
     */
    const Config& getConfig() const { return config_; }
    
    // ==================== 组件访问 ====================
    
    /**
     * @brief 获取Slab分配器（用于特殊需求）
     */
    TXSlabAllocator& getSlabAllocator() { return *slab_allocator_; }
    
    /**
     * @brief 获取Chunk分配器（用于特殊需求）
     */
    TXChunkAllocator& getChunkAllocator() { return *chunk_allocator_; }
    
    /**
     * @brief 获取智能管理器（用于特殊需求）
     */
    TXSmartMemoryManager& getSmartManager() { return *smart_manager_; }

private:
    // ==================== 核心组件 ====================
    
    std::unique_ptr<TXSlabAllocator> slab_allocator_;       // 小对象分配器
    std::unique_ptr<TXChunkAllocator> chunk_allocator_;     // 大对象分配器
    std::unique_ptr<TXSmartMemoryManager> smart_manager_;   // 智能管理器
    
    // ==================== 配置和统计 ====================
    
    Config config_;
    
    // 分配统计
    std::atomic<size_t> small_allocation_count_{0};
    std::atomic<size_t> large_allocation_count_{0};
    std::atomic<size_t> total_allocation_time_us_{0};
    std::atomic<size_t> total_allocations_{0};
    
    // ==================== 内部方法 ====================
    
    /**
     * @brief 选择分配器
     */
    enum class AllocatorType {
        SLAB,   // 小对象（<=8KB）
        CHUNK   // 大对象（>8KB）
    };
    
    AllocatorType selectAllocator(size_t size) const {
        return size <= config_.slab_chunk_threshold ? AllocatorType::SLAB : AllocatorType::CHUNK;
    }
    
    /**
     * @brief 更新统计信息
     */
    void updateStats(AllocatorType type, size_t size, 
                    std::chrono::steady_clock::time_point start_time,
                    std::chrono::steady_clock::time_point end_time);
    
    /**
     * @brief 初始化组件
     */
    void initializeComponents();
};

/**
 * @brief 全局统一内存管理器
 */
class GlobalUnifiedMemoryManager {
public:
    static TXUnifiedMemoryManager& getInstance();
    static void initialize(const TXUnifiedMemoryManager::Config& config = TXUnifiedMemoryManager::Config{});
    static void shutdown();
    
    // 便捷的全局分配接口
    static void* allocate(size_t size) { return getInstance().allocate(size); }
    static bool deallocate(void* ptr) { return getInstance().deallocate(ptr); }
    
    template<typename T>
    static T* allocate(size_t count = 1) { return getInstance().allocate<T>(count); }

private:
    static std::unique_ptr<TXUnifiedMemoryManager> instance_;
    static std::mutex instance_mutex_;
};

} // namespace TinaXlsx
