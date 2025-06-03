#include "TinaXlsx/TXAdvancedParallelFramework.hpp"
#include <algorithm>

namespace TinaXlsx {

TXXlsxTaskScheduler::TXXlsxTaskScheduler(const SchedulerConfig& config)
    : config_(config)
    , stop_(false)
    , nextTaskId_(1)
    , currentMemoryUsage_(0)
    , activeTasks_(0)
    , totalTasksScheduled_(0)
    , tasksCompleted_(0) {
    
    // 创建线程池
    TXLockFreeThreadPool::PoolConfig poolConfig;
    poolConfig.numThreads = config_.maxConcurrentTasks;
    threadPool_ = std::make_unique<TXLockFreeThreadPool>(poolConfig);
    
    // 启动调度器线程
    std::thread schedulerThread([this] { schedulerThread(); });
    schedulerThread.detach();
}

TXXlsxTaskScheduler::~TXXlsxTaskScheduler() {
    stop_.store(true);
    queueCondition_.notify_all();
}

std::vector<std::future<void>> TXXlsxTaskScheduler::scheduleBatch(
    const std::vector<std::pair<TaskMetrics, std::function<void()>>>& tasks) {
    
    std::vector<std::future<void>> futures;
    futures.reserve(tasks.size());
    
    for (const auto& [metrics, task] : tasks) {
        auto future = scheduleTask(metrics, task);
        futures.push_back(std::move(future));
    }
    
    return futures;
}

void TXXlsxTaskScheduler::waitForAll() {
    if (threadPool_) {
        threadPool_->waitForAll();
    }
    
    // 等待所有任务完成
    std::unique_lock<std::mutex> lock(queueMutex_);
    queueCondition_.wait(lock, [this] {
        return taskQueue_.empty() && dependencyMap_.empty();
    });
}

TXXlsxTaskScheduler::SchedulerStats TXXlsxTaskScheduler::getStats() const {
    SchedulerStats stats;
    stats.totalTasksScheduled = totalTasksScheduled_.load();
    stats.tasksCompleted = tasksCompleted_.load();
    stats.currentMemoryUsage = currentMemoryUsage_.load();
    stats.averageTaskTime = 0.0; // 简化实现
    
    {
        std::lock_guard<std::mutex> lock(queueMutex_);
        stats.tasksInQueue = taskQueue_.size();
    }
    
    // 统计任务类型分布
    // 这里简化实现，实际应该维护详细统计
    
    return stats;
}

void TXXlsxTaskScheduler::scheduleTaskInternal(const TaskMetrics& metrics, std::function<void()> task) {
    auto taskId = nextTaskId_.fetch_add(1);
    auto scheduledTask = std::make_unique<ScheduledTask>(taskId, metrics, std::move(task));
    
    {
        std::lock_guard<std::mutex> lock(queueMutex_);
        
        if (config_.enableDependencyTracking && !metrics.dependencies.empty()) {
            // 有依赖的任务放入依赖映射
            dependencyMap_[taskId] = std::move(scheduledTask);
        } else {
            // 无依赖的任务直接入队
            taskQueue_.push(std::move(scheduledTask));
        }
    }
    
    totalTasksScheduled_.fetch_add(1);
    queueCondition_.notify_one();
}

void TXXlsxTaskScheduler::schedulerThread() {
    while (!stop_.load()) {
        std::unique_lock<std::mutex> lock(queueMutex_);
        
        // 等待任务或停止信号
        queueCondition_.wait(lock, [this] {
            return !taskQueue_.empty() || !dependencyMap_.empty() || stop_.load();
        });
        
        if (stop_.load()) {
            break;
        }
        
        // 处理可执行的任务
        while (!taskQueue_.empty()) {
            auto task = std::move(taskQueue_.front());
            taskQueue_.pop();
            
            if (canExecuteTask(*task)) {
                lock.unlock();
                executeTask(std::move(task));
                lock.lock();
            } else {
                // 资源不足，重新入队
                taskQueue_.push(std::move(task));
                break;
            }
        }
        
        // 检查依赖任务
        auto it = dependencyMap_.begin();
        while (it != dependencyMap_.end()) {
            if (canExecuteTask(*it->second)) {
                taskQueue_.push(std::move(it->second));
                it = dependencyMap_.erase(it);
            } else {
                ++it;
            }
        }
    }
}

bool TXXlsxTaskScheduler::canExecuteTask(const ScheduledTask& task) const {
    if (config_.enableResourceMonitoring) {
        // 检查内存使用
        if (currentMemoryUsage_.load() + task.metrics.estimatedMemory > config_.memoryThreshold) {
            return false;
        }
    }
    
    // 检查并发任务数
    if (activeTasks_.load() >= config_.maxConcurrentTasks) {
        return false;
    }
    
    return true;
}

void TXXlsxTaskScheduler::executeTask(std::unique_ptr<ScheduledTask> task) {
    updateResourceUsage(task->metrics, true);
    activeTasks_.fetch_add(1);
    
    // 提交到线程池执行
    if (threadPool_) {
        threadPool_->submit([this, taskPtr = task.release()]() {
            std::unique_ptr<ScheduledTask> taskGuard(taskPtr);
            
            try {
                taskGuard->function();
            } catch (...) {
                // 记录异常但不中断
            }
            
            updateResourceUsage(taskGuard->metrics, false);
            activeTasks_.fetch_sub(1);
            tasksCompleted_.fetch_add(1);
        });
    }
}

void TXXlsxTaskScheduler::updateResourceUsage(const TaskMetrics& metrics, bool starting) {
    if (starting) {
        currentMemoryUsage_.fetch_add(metrics.estimatedMemory);
    } else {
        currentMemoryUsage_.fetch_sub(metrics.estimatedMemory);
    }
}

} // namespace TinaXlsx
