//
// @file TXAdvancedParallelFramework.hpp
// @brief 高级并行处理框架 - 专为XLSX大文件操作设计
//

#pragma once

#include "TXTypes.hpp"
#include "TXResult.hpp"
#include "TXUnifiedMemoryManager.hpp"
#include "TXParallelProcessor.hpp"
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
#include <algorithm>

namespace TinaXlsx {

// 前向声明
struct TXWorkbookContext;
class TXZipArchiveWriter;

/**
 * @brief 🚀 XLSX专用并行任务调度器
 * 
 * 特点：
 * - 任务依赖管理
 * - 资源感知调度
 * - 动态负载均衡
 * - 内存压力感知
 */
class TXXlsxTaskScheduler {
public:
    enum class TaskType {
        CellProcessing,     // 单元格处理
        XmlGeneration,      // XML生成
        Compression,        // 压缩
        IO,                 // 文件IO
        StringProcessing    // 字符串处理
    };
    
    struct TaskMetrics {
        TaskType type;
        size_t estimatedMemory;
        std::chrono::microseconds estimatedTime;
        std::vector<size_t> dependencies;
        
        TaskMetrics(TaskType t, size_t mem = 0, std::chrono::microseconds time = {})
            : type(t), estimatedMemory(mem), estimatedTime(time) {}
    };
    
    struct SchedulerConfig {
        size_t maxConcurrentTasks = std::thread::hardware_concurrency();
        size_t memoryThreshold = 512 * 1024 * 1024; // 512MB
        bool enableDependencyTracking = true;
        bool enableResourceMonitoring = true;
        bool enableAdaptiveScheduling = true;
    };
    
    explicit TXXlsxTaskScheduler(const SchedulerConfig& config = SchedulerConfig{});
    ~TXXlsxTaskScheduler();
    
    /**
     * @brief 提交任务到调度器
     */
    template<typename F, typename... Args>
    std::future<typename std::result_of<F(Args...)>::type> 
    scheduleTask(const TaskMetrics& metrics, F&& f, Args&&... args) {
        using ReturnType = typename std::result_of<F(Args...)>::type;
        
        auto task = std::make_shared<std::packaged_task<ReturnType()>>(
            std::bind(std::forward<F>(f), std::forward<Args>(args)...)
        );
        
        std::future<ReturnType> result = task->get_future();
        
        scheduleTaskInternal(metrics, [task](){ (*task)(); });
        
        return result;
    }
    
    /**
     * @brief 批量调度任务
     */
    std::vector<std::future<void>> scheduleBatch(
        const std::vector<std::pair<TaskMetrics, std::function<void()>>>& tasks);
    
    /**
     * @brief 等待所有任务完成
     */
    void waitForAll();
    
    /**
     * @brief 获取调度器统计信息
     */
    struct SchedulerStats {
        size_t totalTasksScheduled = 0;
        size_t tasksCompleted = 0;
        size_t tasksInQueue = 0;
        size_t currentMemoryUsage = 0;
        double averageTaskTime = 0.0;
        std::unordered_map<TaskType, size_t> taskTypeDistribution;
    };
    
    SchedulerStats getStats() const;

private:
    struct ScheduledTask {
        size_t taskId;
        TaskMetrics metrics;
        std::function<void()> function;
        std::chrono::steady_clock::time_point submitTime;
        std::vector<size_t> waitingFor;
        
        ScheduledTask(size_t id, TaskMetrics m, std::function<void()> f)
            : taskId(id), metrics(std::move(m)), function(std::move(f))
            , submitTime(std::chrono::steady_clock::now()) {}
    };
    
    SchedulerConfig config_;
    std::unique_ptr<TXLockFreeThreadPool> threadPool_;
    
    std::queue<std::unique_ptr<ScheduledTask>> taskQueue_;
    std::unordered_map<size_t, std::unique_ptr<ScheduledTask>> dependencyMap_;
    mutable std::mutex queueMutex_;
    std::condition_variable queueCondition_;
    std::atomic<bool> stop_{false};
    std::atomic<size_t> nextTaskId_{1};
    
    // 资源监控
    std::atomic<size_t> currentMemoryUsage_{0};
    std::atomic<size_t> activeTasks_{0};
    
    // 统计信息
    mutable std::atomic<size_t> totalTasksScheduled_{0};
    mutable std::atomic<size_t> tasksCompleted_{0};
    
    void scheduleTaskInternal(const TaskMetrics& metrics, std::function<void()> task);
    void schedulerThread();
    bool canExecuteTask(const ScheduledTask& task) const;
    void executeTask(std::unique_ptr<ScheduledTask> task);
    void updateResourceUsage(const TaskMetrics& metrics, bool starting);
};

/**
 * @brief 🚀 并行XLSX读取器
 * 
 * 支持大文件的并行读取和解析
 */
class TXParallelXlsxReader {
public:
    struct ReaderConfig {
        size_t numReaderThreads = 2;
        size_t numParserThreads = std::thread::hardware_concurrency() - 2;
        size_t bufferSize = 1024 * 1024; // 1MB
        bool enableStreamingParse = true;
        bool enableMemoryMapping = true;
    };
    
    explicit TXParallelXlsxReader(const ReaderConfig& config = ReaderConfig{});
    ~TXParallelXlsxReader();
    
    /**
     * @brief 并行读取XLSX文件
     */
    TXResult<std::unique_ptr<TXWorkbookContext>> readFile(const std::string& filename);
    
    /**
     * @brief 并行读取工作表
     */
    TXResult<void> readWorksheetParallel(const std::string& xmlData, 
                                        TXWorkbookContext& context, 
                                        size_t sheetIndex);

private:
    ReaderConfig config_;
    std::unique_ptr<TXXlsxTaskScheduler> scheduler_;
    std::unique_ptr<TXUnifiedMemoryManager> memoryManager_;
    
    TXResult<std::vector<std::string>> extractXmlFiles(const std::string& filename);
    TXResult<void> parseSharedStrings(const std::string& xmlData, TXWorkbookContext& context);
    TXResult<void> parseStyles(const std::string& xmlData, TXWorkbookContext& context);
};

/**
 * @brief 🚀 并行XLSX写入器
 * 
 * 高性能的并行写入实现
 */
class TXParallelXlsxWriter {
public:
    struct WriterConfig {
        size_t numWriterThreads = std::thread::hardware_concurrency();
        size_t compressionLevel = 6;
        bool enableParallelCompression = true;
        bool enableStreamingWrite = true;
        size_t bufferSize = 2 * 1024 * 1024; // 2MB
    };
    
    explicit TXParallelXlsxWriter(const WriterConfig& config = WriterConfig{});
    ~TXParallelXlsxWriter();
    
    /**
     * @brief 并行写入XLSX文件
     */
    TXResult<void> writeFile(const std::string& filename, const TXWorkbookContext& context);
    
    /**
     * @brief 获取写入统计信息
     */
    struct WriterStats {
        size_t totalBytesWritten = 0;
        std::chrono::milliseconds totalTime{0};
        std::chrono::milliseconds xmlGenerationTime{0};
        std::chrono::milliseconds compressionTime{0};
        std::chrono::milliseconds ioTime{0};
        double compressionRatio = 0.0;
    };
    
    WriterStats getStats() const;

private:
    WriterConfig config_;
    std::unique_ptr<TXXlsxTaskScheduler> scheduler_;
    std::unique_ptr<TXUnifiedMemoryManager> memoryManager_;
    WriterStats stats_;
    
    TXResult<std::vector<std::pair<std::string, std::vector<uint8_t>>>> 
    generateXmlFilesParallel(const TXWorkbookContext& context);
    
    TXResult<void> compressAndWriteParallel(
        const std::string& filename,
        const std::vector<std::pair<std::string, std::vector<uint8_t>>>& files);
};

} // namespace TinaXlsx
