//
// @file TXBatchPipeline.hpp
// @brief 高效批处理流水线 - 4级流水线架构
//

#pragma once

#include "TXUnifiedMemoryManager.hpp"
#include "TXTypes.hpp"
#include "TXResult.hpp"
#include "TXCompactCell.hpp"
#include <vector>
#include <memory>
#include <atomic>
#include <mutex>
#include <condition_variable>
#include <queue>
#include <thread>
#include <future>
#include <functional>
#include <chrono>

namespace TinaXlsx {

/**
 * @brief 批处理数据包
 */
struct TXBatchData {
    size_t batch_id;
    std::vector<TXCompactCell> cells;
    std::vector<std::string> strings;
    std::vector<uint8_t> binary_data;
    size_t estimated_size;
    std::chrono::steady_clock::time_point timestamp;
    
    // 内存管理
    void* memory_context = nullptr;
    size_t memory_size = 0;
    
    TXBatchData(size_t id) : batch_id(id), estimated_size(0), timestamp(std::chrono::steady_clock::now()) {}
    ~TXBatchData() = default;
};

/**
 * @brief 流水线阶段接口
 */
class TXPipelineStage {
public:
    virtual ~TXPipelineStage() = default;
    
    /**
     * @brief 处理批次数据
     */
    virtual TXResult<std::unique_ptr<TXBatchData>> process(std::unique_ptr<TXBatchData> input) = 0;
    
    /**
     * @brief 获取阶段名称
     */
    virtual std::string getStageName() const = 0;
    
    /**
     * @brief 获取处理统计
     */
    struct StageStats {
        size_t processed_batches = 0;
        size_t failed_batches = 0;
        std::chrono::microseconds total_processing_time{0};
        std::chrono::microseconds avg_processing_time{0};
        size_t memory_usage = 0;
        double throughput = 0.0; // 批次/秒
    };
    
    virtual StageStats getStats() const = 0;
    virtual void resetStats() = 0;
};

/**
 * @brief 4级批处理流水线
 */
class TXBatchPipeline {
public:
    /**
     * @brief 流水线配置
     */
    struct PipelineConfig {
        size_t max_concurrent_batches = 16;        // 最大并发批次
        size_t batch_size_threshold = 1000;       // 批次大小阈值
        size_t memory_limit_mb = 512;             // 内存限制512MB
        size_t queue_capacity = 64;              // 队列容量
        bool enable_memory_optimization = true;   // 启用内存优化
        bool enable_async_processing = true;      // 启用异步处理
        bool enable_performance_monitoring = true; // 启用性能监控
        
        // 各阶段配置
        size_t stage1_threads = 2;  // 数据预处理线程
        size_t stage2_threads = 4;  // XML生成线程
        size_t stage3_threads = 2;  // 压缩处理线程
        size_t stage4_threads = 1;  // 输出写入线程
    };
    
    /**
     * @brief 流水线统计
     */
    struct PipelineStats {
        size_t total_batches_processed = 0;
        size_t total_batches_failed = 0;
        size_t batches_in_pipeline = 0;
        
        std::chrono::microseconds total_pipeline_time{0};
        std::chrono::microseconds avg_pipeline_time{0};
        double overall_throughput = 0.0;
        
        // 各阶段统计
        TXPipelineStage::StageStats stage1_stats;
        TXPipelineStage::StageStats stage2_stats;
        TXPipelineStage::StageStats stage3_stats;
        TXPipelineStage::StageStats stage4_stats;
        
        // 内存统计
        size_t peak_memory_usage = 0;
        size_t current_memory_usage = 0;
        double memory_efficiency = 0.0;
        
        // 队列统计
        size_t max_queue_depth = 0;
        double avg_queue_utilization = 0.0;
    };
    
    explicit TXBatchPipeline(const PipelineConfig& config = PipelineConfig{});
    ~TXBatchPipeline();
    
    // ==================== 流水线控制 ====================
    
    /**
     * @brief 启动流水线
     */
    TXResult<void> start();
    
    /**
     * @brief 停止流水线
     */
    TXResult<void> stop();
    
    /**
     * @brief 暂停流水线
     */
    TXResult<void> pause();
    
    /**
     * @brief 恢复流水线
     */
    TXResult<void> resume();
    
    /**
     * @brief 检查流水线状态
     */
    enum class PipelineState {
        STOPPED,
        STARTING,
        RUNNING,
        PAUSED,
        STOPPING,
        ERROR
    };
    
    PipelineState getState() const { return state_.load(); }
    
    // ==================== 数据处理 ====================
    
    /**
     * @brief 提交批次数据
     */
    TXResult<size_t> submitBatch(std::unique_ptr<TXBatchData> batch);
    
    /**
     * @brief 批量提交
     */
    TXResult<std::vector<size_t>> submitBatches(std::vector<std::unique_ptr<TXBatchData>> batches);
    
    /**
     * @brief 等待所有批次完成
     */
    TXResult<void> waitForCompletion(std::chrono::milliseconds timeout = std::chrono::milliseconds(30000));
    
    /**
     * @brief 获取完成的批次
     */
    std::vector<std::unique_ptr<TXBatchData>> getCompletedBatches();
    
    // ==================== 监控和统计 ====================
    
    /**
     * @brief 获取流水线统计
     */
    PipelineStats getStats() const;
    
    /**
     * @brief 重置统计信息
     */
    void resetStats();
    
    /**
     * @brief 生成性能报告
     */
    std::string generatePerformanceReport() const;
    
    /**
     * @brief 获取实时吞吐量
     */
    double getCurrentThroughput() const;
    
    /**
     * @brief 获取内存使用情况
     */
    size_t getCurrentMemoryUsage() const;
    
    // ==================== 配置管理 ====================
    
    /**
     * @brief 更新配置
     */
    TXResult<void> updateConfig(const PipelineConfig& config);
    
    /**
     * @brief 获取配置
     */
    const PipelineConfig& getConfig() const { return config_; }
    
    /**
     * @brief 设置自定义阶段
     */
    TXResult<void> setCustomStage(size_t stage_index, std::unique_ptr<TXPipelineStage> stage);

private:
    // ==================== 内部数据结构 ====================
    
    PipelineConfig config_;
    std::atomic<PipelineState> state_{PipelineState::STOPPED};
    std::atomic<bool> should_stop_{false};

    // 内存管理器
    std::unique_ptr<TXUnifiedMemoryManager> memory_manager_;
    
    // 4级流水线阶段
    std::unique_ptr<TXPipelineStage> stage1_; // 数据预处理
    std::unique_ptr<TXPipelineStage> stage2_; // XML生成
    std::unique_ptr<TXPipelineStage> stage3_; // 压缩处理
    std::unique_ptr<TXPipelineStage> stage4_; // 输出写入
    
    // 阶段间队列
    std::queue<std::unique_ptr<TXBatchData>> stage1_queue_;
    std::queue<std::unique_ptr<TXBatchData>> stage2_queue_;
    std::queue<std::unique_ptr<TXBatchData>> stage3_queue_;
    std::queue<std::unique_ptr<TXBatchData>> stage4_queue_;
    std::queue<std::unique_ptr<TXBatchData>> completed_queue_;
    
    // 队列同步
    mutable std::mutex stage1_mutex_, stage2_mutex_, stage3_mutex_, stage4_mutex_, completed_mutex_;
    std::condition_variable stage1_cv_, stage2_cv_, stage3_cv_, stage4_cv_, completed_cv_;
    
    // 工作线程
    std::vector<std::thread> stage1_threads_;
    std::vector<std::thread> stage2_threads_;
    std::vector<std::thread> stage3_threads_;
    std::vector<std::thread> stage4_threads_;
    
    // 统计信息
    mutable std::mutex stats_mutex_;
    mutable PipelineStats stats_;
    std::atomic<size_t> next_batch_id_{1};
    
    // 性能监控
    std::chrono::steady_clock::time_point start_time_;
    std::atomic<size_t> batches_processed_last_second_{0};
    std::chrono::steady_clock::time_point last_throughput_update_;
    
    // ==================== 内部方法 ====================
    
    /**
     * @brief 初始化默认阶段
     */
    void initializeDefaultStages();
    
    /**
     * @brief 启动工作线程
     */
    void startWorkerThreads();
    
    /**
     * @brief 停止工作线程
     */
    void stopWorkerThreads();
    
    /**
     * @brief 阶段工作线程
     */
    void stageWorker(size_t stage_index, size_t thread_id);
    
    /**
     * @brief 更新统计信息
     */
    void updateStats();
    
    /**
     * @brief 检查内存限制
     */
    bool checkMemoryLimit() const;
    
    /**
     * @brief 内存优化
     */
    void optimizeMemory();
};

} // namespace TinaXlsx
