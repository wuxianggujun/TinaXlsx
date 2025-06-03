//
// @file TXSIMDParallelProcessor.hpp
// @brief SIMD + 并行处理集成模块 - 最高性能的批处理优化
//

#pragma once

#include "TXUltraCompactCell.hpp"
#include "TXBatchCellManager.hpp"
#include "TXSIMDOptimizations.hpp"
#include <vector>
#include <thread>
#include <future>
#include <atomic>
#include <mutex>
#include <condition_variable>
#include <queue>
#include <functional>
#include <chrono>
#include <map>
#include <limits>

namespace TinaXlsx {

/**
 * @brief SIMD + 并行处理配置
 */
struct SIMDParallelConfig {
    size_t thread_count = 0;                           // 0表示自动检测
    size_t min_batch_size = 1000;                      // 最小批处理大小
    size_t max_batch_size = 100000;                    // 最大批处理大小
    size_t simd_batch_size = 32;                       // SIMD批处理大小
    bool enable_simd = true;                           // 启用SIMD
    bool enable_parallel = true;                       // 启用并行
    bool enable_work_stealing = true;                  // 启用工作窃取
    bool enable_numa_aware = false;                    // 启用NUMA感知
};

/**
 * @brief 高性能任务队列
 */
template<typename T>
class TXLockFreeQueue {
public:
    TXLockFreeQueue(size_t capacity = 1024);
    ~TXLockFreeQueue() = default;
    
    bool push(T&& item);
    bool pop(T& item);
    bool empty() const;
    size_t size() const;
    
private:
    std::vector<T> buffer_;
    std::atomic<size_t> head_{0};
    std::atomic<size_t> tail_{0};
    size_t capacity_;
    size_t mask_;
};

/**
 * @brief SIMD + 并行处理器
 */
class TXSIMDParallelProcessor {
public:
    explicit TXSIMDParallelProcessor(const SIMDParallelConfig& config = SIMDParallelConfig{});
    ~TXSIMDParallelProcessor();
    
    // ==================== 高性能批处理API ====================
    
    /**
     * @brief 超高性能批量转换double到UltraCompactCell
     */
    void ultraFastConvertDoublesToCells(const std::vector<double>& input,
                                       std::vector<UltraCompactCell>& output);
    
    /**
     * @brief 超高性能批量转换int64到UltraCompactCell
     */
    void ultraFastConvertInt64sToCells(const std::vector<int64_t>& input,
                                      std::vector<UltraCompactCell>& output);
    
    /**
     * @brief 超高性能批量设置坐标
     */
    void ultraFastSetCoordinates(std::vector<UltraCompactCell>& cells,
                                const std::vector<uint16_t>& rows,
                                const std::vector<uint16_t>& cols);
    
    /**
     * @brief 超高性能批量内存操作
     */
    void ultraFastClearCells(std::vector<UltraCompactCell>& cells);
    void ultraFastCopyCells(const std::vector<UltraCompactCell>& src,
                           std::vector<UltraCompactCell>& dst);
    
    // ==================== 高性能数值计算 ====================
    
    /**
     * @brief 超高性能并行求和
     */
    double ultraFastSumNumbers(const std::vector<UltraCompactCell>& cells);
    
    /**
     * @brief 超高性能并行统计
     */
    struct UltraFastStats {
        double sum = 0.0;
        double mean = 0.0;
        double min = std::numeric_limits<double>::infinity();
        double max = -std::numeric_limits<double>::infinity();
        size_t count = 0;
        double variance = 0.0;
        double std_dev = 0.0;
    };
    
    UltraFastStats ultraFastCalculateStats(const std::vector<UltraCompactCell>& cells);
    
    /**
     * @brief 超高性能并行数值运算
     */
    void ultraFastNumericOperation(const std::vector<UltraCompactCell>& a,
                                  const std::vector<UltraCompactCell>& b,
                                  std::vector<UltraCompactCell>& result,
                                  const std::function<double(double, double)>& operation);
    
    // ==================== 高性能筛选和排序 ====================
    
    /**
     * @brief 超高性能并行筛选
     */
    std::vector<UltraCompactCell> ultraFastFilterCells(
        const std::vector<UltraCompactCell>& cells,
        const std::function<bool(const UltraCompactCell&)>& predicate);
    
    /**
     * @brief 超高性能并行排序
     */
    void ultraFastSortCells(std::vector<UltraCompactCell>& cells,
                           const std::function<bool(const UltraCompactCell&, const UltraCompactCell&)>& comparator);
    
    // ==================== 批处理管理器集成 ====================
    
    /**
     * @brief 超高性能批处理
     */
    size_t ultraFastProcessBatch(TXBatchCellManager& manager,
                                const std::vector<CellData>& cells);
    
    /**
     * @brief 超高性能批量获取
     */
    std::vector<CellData> ultraFastGetBatch(const TXBatchCellManager& manager,
                                           const std::vector<CellRange>& ranges);
    
    // ==================== 性能监控和基准测试 ====================
    
    /**
     * @brief 性能统计
     */
    struct PerformanceMetrics {
        // 基本统计
        size_t total_operations = 0;
        double total_time_ms = 0.0;
        double avg_time_per_operation_ns = 0.0;
        size_t operations_per_second = 0;
        
        // SIMD统计
        size_t simd_operations = 0;
        double simd_speedup = 0.0;
        std::string simd_type;
        
        // 并行统计
        size_t parallel_tasks = 0;
        double parallel_efficiency = 0.0;
        size_t thread_count = 0;
        
        // 内存统计
        size_t memory_bandwidth_mb_s = 0;
        size_t cache_hit_rate_percent = 0;
        
        // 瓶颈分析
        std::string bottleneck_analysis;
        std::vector<std::string> optimization_suggestions;
    };
    
    PerformanceMetrics getPerformanceMetrics() const;
    void resetPerformanceMetrics();
    
    /**
     * @brief 全面性能基准测试
     */
    struct ComprehensiveBenchmarkResult {
        // 不同数据大小的性能
        std::map<size_t, PerformanceMetrics> size_performance;
        
        // 不同线程数的性能
        std::map<size_t, PerformanceMetrics> thread_performance;
        
        // SIMD vs 标量性能对比
        PerformanceMetrics simd_performance;
        PerformanceMetrics scalar_performance;
        
        // 最优配置建议
        SIMDParallelConfig optimal_config;
        std::string performance_summary;
    };
    
    ComprehensiveBenchmarkResult runComprehensiveBenchmark();
    
    // ==================== 配置和控制 ====================
    
    /**
     * @brief 动态调整配置
     */
    void updateConfig(const SIMDParallelConfig& config);
    SIMDParallelConfig getConfig() const { return config_; }
    
    /**
     * @brief 获取系统信息
     */
    struct SystemInfo {
        size_t cpu_cores = 0;
        size_t logical_processors = 0;
        std::string cpu_brand;
        std::vector<std::string> simd_features;
        size_t l1_cache_size = 0;
        size_t l2_cache_size = 0;
        size_t l3_cache_size = 0;
        bool numa_available = false;
        size_t numa_nodes = 0;
    };
    
    static SystemInfo getSystemInfo();
    
    /**
     * @brief 自动优化配置
     */
    SIMDParallelConfig autoOptimizeConfig(size_t typical_data_size = 100000);

private:
    SIMDParallelConfig config_;
    std::vector<std::thread> workers_;
    std::vector<std::unique_ptr<TXLockFreeQueue<std::function<void()>>>> task_queues_;
    std::atomic<bool> stop_flag_{false};
    std::atomic<size_t> next_queue_{0};
    
    // 性能监控
    mutable std::mutex metrics_mutex_;
    mutable PerformanceMetrics metrics_;
    
    // NUMA支持
    std::vector<std::vector<size_t>> numa_topology_;
    
    // ==================== 内部实现 ====================
    
    /**
     * @brief 工作线程
     */
    void workerThread(size_t thread_id);
    
    /**
     * @brief 工作窃取
     */
    bool stealWork(size_t thief_id, std::function<void()>& task);
    
    /**
     * @brief 任务分割策略
     */
    std::vector<std::pair<size_t, size_t>> calculateOptimalSplits(size_t total_size) const;
    
    /**
     * @brief NUMA感知的任务分配
     */
    size_t selectOptimalQueue(size_t preferred_node = 0) const;
    
    /**
     * @brief 性能计数器
     */
    void updateMetrics(const std::string& operation, 
                      std::chrono::steady_clock::time_point start,
                      std::chrono::steady_clock::time_point end,
                      size_t operations_count);
    
    /**
     * @brief 缓存友好的数据分布
     */
    template<typename T>
    std::vector<std::vector<T>> distributeCacheFriendly(const std::vector<T>& data) const;
    
    /**
     * @brief 并行归约模板
     */
    template<typename T, typename BinaryOp>
    T parallelReduce(const std::vector<T>& data, T init, BinaryOp op);
    
    /**
     * @brief 并行映射模板
     */
    template<typename InputT, typename OutputT, typename UnaryOp>
    void parallelMap(const std::vector<InputT>& input,
                    std::vector<OutputT>& output,
                    UnaryOp op);
    
    /**
     * @brief 混合SIMD+并行处理模板
     */
    template<typename T, typename SIMDOp, typename ScalarOp>
    void hybridSIMDParallelProcess(const std::vector<T>& input,
                                  std::vector<T>& output,
                                  SIMDOp simd_op,
                                  ScalarOp scalar_op);
};

// ==================== 模板实现 ====================

template<typename T>
TXLockFreeQueue<T>::TXLockFreeQueue(size_t capacity) 
    : capacity_(capacity), mask_(capacity - 1) {
    // 确保容量是2的幂
    if ((capacity & (capacity - 1)) != 0) {
        throw std::invalid_argument("Capacity must be a power of 2");
    }
    buffer_.resize(capacity);
}

template<typename T>
bool TXLockFreeQueue<T>::push(T&& item) {
    size_t tail = tail_.load(std::memory_order_relaxed);
    size_t next_tail = (tail + 1) & mask_;
    
    if (next_tail == head_.load(std::memory_order_acquire)) {
        return false; // 队列满
    }
    
    buffer_[tail] = std::move(item);
    tail_.store(next_tail, std::memory_order_release);
    return true;
}

template<typename T>
bool TXLockFreeQueue<T>::pop(T& item) {
    size_t head = head_.load(std::memory_order_relaxed);
    
    if (head == tail_.load(std::memory_order_acquire)) {
        return false; // 队列空
    }
    
    item = std::move(buffer_[head]);
    head_.store((head + 1) & mask_, std::memory_order_release);
    return true;
}

template<typename T>
bool TXLockFreeQueue<T>::empty() const {
    return head_.load(std::memory_order_acquire) == tail_.load(std::memory_order_acquire);
}

template<typename T>
size_t TXLockFreeQueue<T>::size() const {
    size_t head = head_.load(std::memory_order_acquire);
    size_t tail = tail_.load(std::memory_order_acquire);
    return (tail - head) & mask_;
}

} // namespace TinaXlsx
