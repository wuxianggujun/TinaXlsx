#include "TinaXlsx/TXParallelProcessor.hpp"
#include <iostream>

namespace TinaXlsx {

// TXLockFreeThreadPool 实现（TXParallelProcessor.hpp 中的版本）
TXLockFreeThreadPool::TXLockFreeThreadPool(const PoolConfig& config)
    : config_(config) {
    
    // 创建线程本地队列
    localQueues_.reserve(config_.numThreads);
    for (size_t i = 0; i < config_.numThreads; ++i) {
        localQueues_.emplace_back(std::make_unique<ThreadLocalQueue>());
    }
    
    // 创建工作线程
    workers_.reserve(config_.numThreads);
    for (size_t i = 0; i < config_.numThreads; ++i) {
        workers_.emplace_back([this, i] { workerThread(i); });
    }
}

TXLockFreeThreadPool::~TXLockFreeThreadPool() {
    stop_.store(true);
    
    // 唤醒所有线程
    for (auto& queue : localQueues_) {
        queue->condition.notify_all();
    }
    
    // 等待所有线程完成
    for (auto& worker : workers_) {
        if (worker.joinable()) {
            worker.join();
        }
    }
}

bool TXLockFreeThreadPool::submitTaskInternal(std::function<void()> task, TaskPriority priority) {
    if (stop_.load()) {
        return false;
    }
    
    // 选择队列
    auto* queue = selectQueue();
    if (!queue) {
        return false;
    }
    
    // 添加任务到队列
    {
        std::lock_guard<std::mutex> lock(queue->mutex);
        queue->tasks.emplace(std::move(task), priority);
        queue->taskCount.fetch_add(1);
    }
    
    // 通知工作线程
    queue->condition.notify_one();
    totalTasksProcessed_.fetch_add(1);
    
    return true;
}

void TXLockFreeThreadPool::waitForAll() {
    // 简单实现：等待所有队列为空
    bool allEmpty = false;
    while (!allEmpty) {
        allEmpty = true;
        for (const auto& queue : localQueues_) {
            if (queue->taskCount.load() > 0) {
                allEmpty = false;
                break;
            }
        }
        if (!allEmpty) {
            std::this_thread::sleep_for(std::chrono::milliseconds(1));
        }
    }
}

TXLockFreeThreadPool::PoolStats TXLockFreeThreadPool::getStats() const {
    PoolStats stats;
    stats.totalTasksProcessed = totalTasksProcessed_.load();
    stats.workStealingCount = workStealingCount_.load();
    stats.activeThreads = config_.numThreads;
    
    // 计算队列中的任务数
    for (const auto& queue : localQueues_) {
        stats.tasksInQueue += queue->taskCount.load();
    }
    
    // 计算平均任务时间
    if (stats.totalTasksProcessed > 0) {
        stats.averageTaskTime = static_cast<double>(totalProcessingTime_.load()) / stats.totalTasksProcessed;
    }
    
    return stats;
}

void TXLockFreeThreadPool::resizeThreadPool(size_t newSize) {
    // 简单实现：不支持动态调整
    // 在实际应用中可以实现更复杂的逻辑
}

void TXLockFreeThreadPool::workerThread(size_t threadId) {
    auto* myQueue = localQueues_[threadId].get();

    while (!stop_.load()) {
        std::function<void()> taskFunction;
        bool gotTask = false;

        // 尝试从自己的队列获取任务
        {
            std::unique_lock<std::mutex> lock(myQueue->mutex);
            if (!myQueue->tasks.empty()) {
                taskFunction = std::move(myQueue->tasks.front().function);
                myQueue->tasks.pop();
                myQueue->taskCount.fetch_sub(1);
                gotTask = true;
            }
        }

        // 如果没有任务，尝试工作窃取
        if (!gotTask && config_.enableWorkStealing) {
            Task stolenTask([](){}, TaskPriority::Normal);
            if (tryStealTask(threadId, stolenTask)) {
                taskFunction = std::move(stolenTask.function);
                gotTask = true;
                workStealingCount_.fetch_add(1);
            }
        }

        if (gotTask) {
            // 执行任务
            auto startTime = std::chrono::steady_clock::now();
            try {
                if (taskFunction) {
                    taskFunction();
                }
            } catch (...) {
                // 忽略异常
            }
            auto endTime = std::chrono::steady_clock::now();

            auto duration = std::chrono::duration_cast<std::chrono::microseconds>(endTime - startTime);
            totalProcessingTime_.fetch_add(duration.count());
        } else {
            // 没有任务，等待
            std::unique_lock<std::mutex> lock(myQueue->mutex);
            myQueue->condition.wait_for(lock, std::chrono::milliseconds(10));
        }
    }
}

bool TXLockFreeThreadPool::tryStealTask(size_t thiefId, Task& stolenTask) {
    // 尝试从其他线程的队列偷取任务
    for (size_t i = 0; i < localQueues_.size(); ++i) {
        if (i == thiefId) continue;
        
        auto* victimQueue = localQueues_[i].get();
        std::unique_lock<std::mutex> lock(victimQueue->mutex, std::try_to_lock);
        
        if (lock.owns_lock() && !victimQueue->tasks.empty()) {
            stolenTask = std::move(victimQueue->tasks.front());
            victimQueue->tasks.pop();
            victimQueue->taskCount.fetch_sub(1);
            return true;
        }
    }
    return false;
}

TXLockFreeThreadPool::ThreadLocalQueue* TXLockFreeThreadPool::selectQueue() {
    // 简单的轮询选择
    size_t index = nextQueueIndex_.fetch_add(1) % localQueues_.size();
    return localQueues_[index].get();
}

} // namespace TinaXlsx
