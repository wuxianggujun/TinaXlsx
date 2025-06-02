//
// @file TXParallelProcessor.hpp
// @brief 并行处理框架 - 提升大数据量处理性能
//

#pragma once

#include "TXTypes.hpp"
#include "TXResult.hpp"
#include <vector>
#include <thread>
#include <future>
#include <functional>
#include <memory>
#include <atomic>

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
 * @brief 线程池
 * 
 * 高性能线程池实现，支持任务队列和工作窃取
 */
class TXThreadPool {
public:
    explicit TXThreadPool(size_t numThreads = std::thread::hardware_concurrency());
    ~TXThreadPool();
    
    /**
     * @brief 提交任务
     */
    template<typename F, typename... Args>
    auto submit(F&& f, Args&&... args) -> std::future<typename std::result_of<F(Args...)>::type> {
        using ReturnType = typename std::result_of<F(Args...)>::type;
        
        auto task = std::make_shared<std::packaged_task<ReturnType()>>(
            std::bind(std::forward<F>(f), std::forward<Args>(args)...)
        );
        
        std::future<ReturnType> result = task->get_future();
        
        {
            std::unique_lock<std::mutex> lock(queueMutex_);
            if (stop_) {
                throw std::runtime_error("ThreadPool is stopped");
            }
            
            tasks_.emplace([task](){ (*task)(); });
        }
        
        condition_.notify_one();
        return result;
    }
    
    /**
     * @brief 等待所有任务完成
     */
    void waitForAll();
    
    /**
     * @brief 获取线程数量
     */
    size_t getThreadCount() const { return workers_.size(); }
    
    /**
     * @brief 获取队列中的任务数量
     */
    size_t getQueueSize() const;

private:
    std::vector<std::thread> workers_;
    std::queue<std::function<void()>> tasks_;
    std::mutex queueMutex_;
    std::condition_variable condition_;
    std::atomic<bool> stop_;
    
    void workerThread();
};

/**
 * @brief 并行单元格处理器
 * 
 * 专门用于并行处理大量单元格操作
 */
class TXParallelCellProcessor {
public:
    explicit TXParallelCellProcessor(size_t numThreads = std::thread::hardware_concurrency());
    
    /**
     * @brief 并行设置单元格值
     */
    template<typename CellManager>
    TXResult<size_t> parallelSetCellValues(
        CellManager& manager,
        const std::vector<std::pair<TXCoordinate, cell_value_t>>& values,
        size_t batchSize = 1000
    ) {
        if (values.empty()) {
            return Ok(static_cast<size_t>(0));
        }
        
        // 分批处理
        std::vector<std::future<size_t>> futures;
        futures.reserve((values.size() + batchSize - 1) / batchSize);
        
        for (size_t i = 0; i < values.size(); i += batchSize) {
            size_t endIdx = std::min(i + batchSize, values.size());
            
            auto future = threadPool_.submit([&manager, &values, i, endIdx]() -> size_t {
                size_t count = 0;
                for (size_t j = i; j < endIdx; ++j) {
                    if (manager.setCellValue(values[j].first, values[j].second)) {
                        ++count;
                    }
                }
                return count;
            });
            
            futures.push_back(std::move(future));
        }
        
        // 收集结果
        size_t totalCount = 0;
        for (auto& future : futures) {
            try {
                totalCount += future.get();
            } catch (const std::exception& e) {
                return Err<size_t>(TXErrorCode::ProcessingError, 
                                 "Parallel processing failed: " + std::string(e.what()));
            }
        }
        
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
            auto future = threadPool_.submit([&manager, startRow, startCol, &values, rowIdx]() -> size_t {
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
                return Err<size_t>(TXErrorCode::ProcessingError, 
                                 "Parallel range processing failed: " + std::string(e.what()));
            }
        }
        
        return Ok(totalCount);
    }

private:
    TXThreadPool threadPool_;
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
    TXThreadPool threadPool_;
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
    TXThreadPool threadPool_;
    
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
    TXParallelCellProcessor& getCellProcessor() { return cellProcessor_; }
    
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
    TXParallelCellProcessor cellProcessor_;
    TXParallelXmlGenerator xmlGenerator_;
    TXParallelZipProcessor zipProcessor_;
    
    std::atomic<bool> parallelEnabled_;
    mutable std::atomic<size_t> totalTasksProcessed_;
    mutable std::chrono::microseconds totalProcessingTime_{0};
};

} // namespace TinaXlsx
