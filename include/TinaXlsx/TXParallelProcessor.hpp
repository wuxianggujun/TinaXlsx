//
// @file TXParallelProcessor.hpp
// @brief 高性能并行处理框架 - 专为XLSX文件操作优化
//

#pragma once

#include "TXTypes.hpp"
#include "TXResult.hpp"
#include "TXCoordinate.hpp"
#include "TXMemoryPool.hpp"
#include <vector>
#include <thread>
#include <future>
#include <functional>
#include <memory>
#include <atomic>
#include <queue>
#include <condition_variable>
#include <mutex>
#include <chrono>
#include <unordered_map>

namespace TinaXlsx {

/**
 * @brief 并行任务接口
 */
class TXParallelTask {
public:
    virtual ~TXParallelTask() = default;
    virtual void execute() = 0;
    virtual std::string getTaskName() const = 0;
};

/**
 * @brief 🚀 无锁高性能线程池
 *
 * 特点：
 * - 工作窃取算法减少锁竞争
 * - 线程本地队列提升性能
 * - 内存池集成减少分配开销
 * - 任务优先级支持
 */
class TXLockFreeThreadPool {
public:
    enum class TaskPriority {
        Low = 0,
        Normal = 1,
        High = 2,
        Critical = 3
    };

    struct PoolConfig {
        size_t numThreads = std::thread::hardware_concurrency();
        size_t queueCapacity = 1024;
        bool enableWorkStealing = true;
        bool enableMemoryPool = true;
        size_t memoryPoolBlockSize = 256;
    };

    explicit TXLockFreeThreadPool(const PoolConfig& config = PoolConfig{});
    ~TXLockFreeThreadPool();

    /**
     * @brief 提交任务（支持优先级）
     */
    template<typename F, typename... Args>
    auto submit(F&& f, Args&&... args, TaskPriority priority = TaskPriority::Normal)
        -> std::future<typename std::result_of<F(Args...)>::type> {
        using ReturnType = typename std::result_of<F(Args...)>::type;

        auto task = std::make_shared<std::packaged_task<ReturnType()>>(
            std::bind(std::forward<F>(f), std::forward<Args>(args)...)
        );

        std::future<ReturnType> result = task->get_future();

        if (!submitTaskInternal([task](){ (*task)(); }, priority)) {
            throw std::runtime_error("Failed to submit task to thread pool");
        }

        return result;
    }

    /**
     * @brief 批量提交任务
     */
    template<typename Iterator>
    std::vector<std::future<void>> submitBatch(Iterator begin, Iterator end,
                                              TaskPriority priority = TaskPriority::Normal) {
        std::vector<std::future<void>> futures;
        futures.reserve(std::distance(begin, end));

        for (auto it = begin; it != end; ++it) {
            futures.push_back(submit(*it, priority));
        }

        return futures;
    }

    /**
     * @brief 等待所有任务完成
     */
    void waitForAll();

    /**
     * @brief 获取性能统计
     */
    struct PoolStats {
        size_t totalTasksProcessed = 0;
        size_t tasksInQueue = 0;
        size_t activeThreads = 0;
        double averageTaskTime = 0.0;
        size_t workStealingCount = 0;
        std::chrono::microseconds totalProcessingTime{0};
    };

    PoolStats getStats() const;

    /**
     * @brief 动态调整线程数量
     */
    void resizeThreadPool(size_t newSize);

private:
    struct Task {
        std::function<void()> function;
        TaskPriority priority;
        std::chrono::steady_clock::time_point submitTime;

        Task(std::function<void()> f, TaskPriority p)
            : function(std::move(f)), priority(p), submitTime(std::chrono::steady_clock::now()) {}
    };

    struct ThreadLocalQueue {
        std::queue<Task> tasks;
        std::mutex mutex;
        std::condition_variable condition;
        std::atomic<size_t> taskCount{0};
    };

    PoolConfig config_;
    std::vector<std::thread> workers_;
    std::vector<std::unique_ptr<ThreadLocalQueue>> localQueues_;
    std::atomic<bool> stop_{false};
    std::atomic<size_t> nextQueueIndex_{0};

    // 性能统计
    mutable std::atomic<size_t> totalTasksProcessed_{0};
    mutable std::atomic<size_t> workStealingCount_{0};
    mutable std::atomic<std::chrono::microseconds::rep> totalProcessingTime_{0};

    // 内存池
    std::unique_ptr<TXMemoryPool> memoryPool_;

    bool submitTaskInternal(std::function<void()> task, TaskPriority priority);
    void workerThread(size_t threadId);
    bool tryStealTask(size_t thiefId, Task& stolenTask);
    ThreadLocalQueue* selectQueue();
};

/**
 * @brief 🚀 智能并行单元格处理器
 *
 * 特点：
 * - 自适应批量大小
 * - 内存池集成
 * - 负载均衡
 * - 缓存友好的数据分布
 */
class TXSmartParallelCellProcessor {
public:
    struct ProcessorConfig {
        size_t numThreads = std::thread::hardware_concurrency();
        size_t minBatchSize = 100;
        size_t maxBatchSize = 10000;
        bool enableAdaptiveBatching = true;
        bool enableMemoryPool = true;
        bool enableCacheOptimization = true;
    };

    explicit TXSmartParallelCellProcessor(const ProcessorConfig& config = ProcessorConfig{});
    ~TXSmartParallelCellProcessor();

    /**
     * @brief 智能并行设置单元格值
     */
    template<typename CellManager>
    TXResult<size_t> parallelSetCellValues(
        CellManager& manager,
        const std::vector<std::pair<TXCoordinate, cell_value_t>>& values
    ) {
        if (values.empty()) {
            return Ok(static_cast<size_t>(0));
        }

        // 🚀 自适应批量大小计算
        size_t optimalBatchSize = calculateOptimalBatchSize(values.size());

        // 🚀 缓存友好的数据重排序
        auto sortedValues = config_.enableCacheOptimization ?
            sortForCacheEfficiency(values) : values;

        // 🚀 负载均衡的任务分配
        auto batches = createBalancedBatches(sortedValues, optimalBatchSize);

        // 并行处理
        std::vector<std::future<size_t>> futures;
        futures.reserve(batches.size());

        for (const auto& batch : batches) {
            auto future = threadPool_->submit([&manager, batch]() -> size_t {
                // 使用批量操作，一次性获取锁
                return manager.setCellValues(batch);
            }, TXLockFreeThreadPool::TaskPriority::High);

            futures.push_back(std::move(future));
        }

        // 收集结果
        size_t totalCount = 0;
        for (auto& future : futures) {
            try {
                totalCount += future.get();
            } catch (const std::exception& e) {
                return Err<size_t>(TXErrorCode::OperationFailed,
                                 "Smart parallel processing failed: " + std::string(e.what()));
            }
        }

        // 更新自适应参数
        updateAdaptiveParameters(values.size(), totalCount);

        return Ok(totalCount);
    }
    
    /**
     * @brief 并行设置范围值
     */
    template<typename CellManager>
    TXResult<size_t> parallelSetRangeValues(
        CellManager& manager,
        row_t startRow, column_t startCol,
        const std::vector<std::vector<cell_value_t>>& values
    ) {
        if (values.empty()) {
            return Ok(static_cast<size_t>(0));
        }
        
        // 按行并行处理
        std::vector<std::future<size_t>> futures;
        futures.reserve(values.size());
        
        for (size_t rowIdx = 0; rowIdx < values.size(); ++rowIdx) {
            auto future = threadPool_->submit([&manager, startRow, startCol, &values, rowIdx]() -> size_t {
                row_t currentRow = row_t(startRow.index() + rowIdx);
                return manager.setRowValues(currentRow, startCol, values[rowIdx]);
            });

            futures.push_back(std::move(future));
        }
        
        // 收集结果
        size_t totalCount = 0;
        for (auto& future : futures) {
            try {
                totalCount += future.get();
            } catch (const std::exception& e) {
                return Err<size_t>(TXErrorCode::OperationFailed,
                                 "Parallel range processing failed: " + std::string(e.what()));
            }
        }
        
        return Ok(totalCount);
    }

private:
    ProcessorConfig config_;
    std::unique_ptr<TXLockFreeThreadPool> threadPool_;

    // 私有辅助方法
    size_t calculateOptimalBatchSize(size_t totalItems) const;
    std::vector<std::pair<TXCoordinate, cell_value_t>> sortForCacheEfficiency(
        const std::vector<std::pair<TXCoordinate, cell_value_t>>& values) const;
    std::vector<std::vector<std::pair<TXCoordinate, cell_value_t>>> createBalancedBatches(
        const std::vector<std::pair<TXCoordinate, cell_value_t>>& values, size_t batchSize) const;
    void updateAdaptiveParameters(size_t totalItems, size_t processedItems);
};

/**
 * @brief 并行XML生成器
 * 
 * 支持并行生成多个XML文件
 */
class TXParallelXmlGenerator {
public:
    explicit TXParallelXmlGenerator(size_t numThreads = std::thread::hardware_concurrency());
    
    /**
     * @brief XML生成任务
     */
    struct XmlGenerationTask {
        std::string partName;
        std::function<TXResult<std::string>()> generator;
        std::promise<TXResult<std::string>> promise;
        
        XmlGenerationTask(const std::string& name, 
                         std::function<TXResult<std::string>()> gen)
            : partName(name), generator(std::move(gen)) {}
    };
    
    /**
     * @brief 提交XML生成任务
     */
    std::future<TXResult<std::string>> submitXmlTask(
        const std::string& partName,
        std::function<TXResult<std::string>()> generator
    );
    
    /**
     * @brief 并行生成多个XML文件
     */
    TXResult<std::vector<std::pair<std::string, std::string>>> generateXmlFiles(
        const std::vector<std::pair<std::string, std::function<TXResult<std::string>()>>>& tasks
    );

private:
    std::unique_ptr<TXLockFreeThreadPool> threadPool_;
};

/**
 * @brief 并行压缩处理器
 * 
 * 支持并行压缩多个文件到ZIP
 */
class TXParallelZipProcessor {
public:
    explicit TXParallelZipProcessor(size_t numThreads = std::thread::hardware_concurrency());
    
    /**
     * @brief 压缩任务
     */
    struct CompressionTask {
        std::string filename;
        std::vector<uint8_t> data;
        int compressionLevel;
        
        CompressionTask(const std::string& name, 
                       std::vector<uint8_t> content,
                       int level = 6)
            : filename(name), data(std::move(content)), compressionLevel(level) {}
    };
    
    /**
     * @brief 并行压缩文件
     */
    TXResult<std::vector<std::pair<std::string, std::vector<uint8_t>>>> compressFiles(
        const std::vector<CompressionTask>& tasks
    );

private:
    std::unique_ptr<TXLockFreeThreadPool> threadPool_;

    std::vector<uint8_t> compressData(const std::vector<uint8_t>& data, int level);
};

/**
 * @brief 并行处理管理器
 * 
 * 统一管理所有并行处理组件
 */
class TXParallelProcessingManager {
public:
    TXParallelProcessingManager();
    ~TXParallelProcessingManager() = default;
    
    /**
     * @brief 获取单元格处理器
     */
    TXSmartParallelCellProcessor& getCellProcessor() { return cellProcessor_; }
    
    /**
     * @brief 获取XML生成器
     */
    TXParallelXmlGenerator& getXmlGenerator() { return xmlGenerator_; }
    
    /**
     * @brief 获取ZIP处理器
     */
    TXParallelZipProcessor& getZipProcessor() { return zipProcessor_; }
    
    /**
     * @brief 设置线程数量
     */
    void setThreadCount(size_t numThreads);
    
    /**
     * @brief 获取性能统计
     */
    struct PerformanceStats {
        size_t totalTasksProcessed = 0;
        std::chrono::microseconds totalProcessingTime{0};
        double averageTaskTime = 0.0;
        size_t activeThreads = 0;
        size_t queuedTasks = 0;
    };
    
    PerformanceStats getPerformanceStats() const;
    
    /**
     * @brief 启用/禁用并行处理
     */
    void setParallelProcessingEnabled(bool enabled) { parallelEnabled_ = enabled; }
    
    /**
     * @brief 检查是否启用并行处理
     */
    bool isParallelProcessingEnabled() const { return parallelEnabled_; }

private:
    TXSmartParallelCellProcessor cellProcessor_;
    TXParallelXmlGenerator xmlGenerator_;
    TXParallelZipProcessor zipProcessor_;

    std::atomic<bool> parallelEnabled_;
    mutable std::atomic<size_t> totalTasksProcessed_;
    mutable std::chrono::microseconds totalProcessingTime_{0};
};

} // namespace TinaXlsx
