//
// @file TXAsyncProcessingFramework.cpp
// @brief 异步处理框架实现
//

#include "TinaXlsx/TXAsyncProcessingFramework.hpp"
#include <algorithm>
#include <sstream>
#include <iomanip>

namespace TinaXlsx {

// ==================== TXAsyncProcessingFramework 实现 ====================

TXAsyncProcessingFramework::TXAsyncProcessingFramework(TXUnifiedMemoryManager& memory_manager,
                                                       const FrameworkConfig& config)
    : memory_manager_(memory_manager), config_(config) {
    
    // 初始化无锁队列
    normal_queue_ = std::make_unique<TXLockFreeQueue<std::unique_ptr<TXAsyncTask>>>(config_.task_queue_capacity);
    high_priority_queue_ = std::make_unique<TXLockFreeQueue<std::unique_ptr<TXAsyncTask>>>(config_.high_priority_queue_capacity);
}

TXAsyncProcessingFramework::~TXAsyncProcessingFramework() {
    if (state_.load() == FrameworkState::RUNNING) {
        stop();
    }
}

TXResult<void> TXAsyncProcessingFramework::start() {
    FrameworkState expected = FrameworkState::STOPPED;
    if (!state_.compare_exchange_strong(expected, FrameworkState::STARTING)) {
        return Err(TXErrorCode::Unknown, "Framework is not in STOPPED state");
    }

    try {
        should_stop_.store(false);

        // 确定工作线程数
        size_t thread_count = config_.worker_thread_count;
        if (thread_count == 0) {
            thread_count = getOptimalThreadCount();
        }

        // 启动工作线程
        worker_threads_.reserve(thread_count);
        for (size_t i = 0; i < thread_count; ++i) {
            worker_threads_.emplace_back(&TXAsyncProcessingFramework::workerThreadFunction, this, i);
        }

        state_.store(FrameworkState::RUNNING);

        return Ok();

    } catch (const std::exception& e) {
        state_.store(FrameworkState::STOPPED);
        return Err(TXErrorCode::Unknown, "Failed to start framework: " + std::string(e.what()));
    }
}

TXResult<void> TXAsyncProcessingFramework::stop() {
    FrameworkState expected = FrameworkState::RUNNING;
    if (!state_.compare_exchange_strong(expected, FrameworkState::STOPPING)) {
        expected = FrameworkState::PAUSED;
        if (!state_.compare_exchange_strong(expected, FrameworkState::STOPPING)) {
            return Err(TXErrorCode::Unknown, "Framework is not in RUNNING or PAUSED state");
        }
    }

    try {
        // 设置停止标志
        should_stop_.store(true);

        // 等待所有工作线程结束
        for (auto& thread : worker_threads_) {
            if (thread.joinable()) {
                thread.join();
            }
        }
        worker_threads_.clear();

        state_.store(FrameworkState::STOPPED);
        should_stop_.store(false);

        return Ok();

    } catch (const std::exception& e) {
        return Err(TXErrorCode::Unknown, "Failed to stop framework: " + std::string(e.what()));
    }
}

TXResult<void> TXAsyncProcessingFramework::pause() {
    FrameworkState expected = FrameworkState::RUNNING;
    if (!state_.compare_exchange_strong(expected, FrameworkState::PAUSED)) {
        return Err(TXErrorCode::Unknown, "Framework is not in RUNNING state");
    }

    return Ok();
}

TXResult<void> TXAsyncProcessingFramework::resume() {
    FrameworkState expected = FrameworkState::PAUSED;
    if (!state_.compare_exchange_strong(expected, FrameworkState::RUNNING)) {
        return Err(TXErrorCode::Unknown, "Framework is not in PAUSED state");
    }

    return Ok();
}

TXResult<void> TXAsyncProcessingFramework::waitForAll(std::chrono::milliseconds timeout) {
    auto start_time = std::chrono::steady_clock::now();

    while (true) {
        // 检查队列是否为空
        if (normal_queue_->empty() && high_priority_queue_->empty()) {
            // 等待一小段时间确保所有任务都完成
            std::this_thread::sleep_for(std::chrono::milliseconds(10));
            if (normal_queue_->empty() && high_priority_queue_->empty()) {
                break;
            }
        }

        // 检查超时
        auto current_time = std::chrono::steady_clock::now();
        if (std::chrono::duration_cast<std::chrono::milliseconds>(current_time - start_time) >= timeout) {
            return Err(TXErrorCode::Unknown, "Timeout waiting for all tasks to complete");
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }

    return Ok();
}

TXResult<void> TXAsyncProcessingFramework::waitForCount(size_t count, std::chrono::milliseconds timeout) {
    auto start_time = std::chrono::steady_clock::now();
    size_t initial_completed = 0;

    {
        std::lock_guard<std::mutex> lock(stats_mutex_);
        initial_completed = stats_.total_tasks_completed;
    }

    while (true) {
        size_t current_completed = 0;
        {
            std::lock_guard<std::mutex> lock(stats_mutex_);
            current_completed = stats_.total_tasks_completed;
        }

        if (current_completed - initial_completed >= count) {
            break;
        }

        // 检查超时
        auto current_time = std::chrono::steady_clock::now();
        if (std::chrono::duration_cast<std::chrono::milliseconds>(current_time - start_time) >= timeout) {
            return Err(TXErrorCode::Unknown, "Timeout waiting for " + std::to_string(count) + " tasks to complete");
        }

        std::this_thread::sleep_for(std::chrono::milliseconds(1));
    }

    return Ok();
}

size_t TXAsyncProcessingFramework::clearQueue() {
    size_t cleared = 0;
    
    // 清空普通优先级队列
    std::unique_ptr<TXAsyncTask> task;
    while (normal_queue_->dequeue(task)) {
        cleared++;
    }
    
    // 清空高优先级队列
    while (high_priority_queue_->dequeue(task)) {
        cleared++;
    }
    
    return cleared;
}

TXAsyncProcessingFramework::FrameworkStats TXAsyncProcessingFramework::getStats() const {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    
    FrameworkStats current_stats = stats_;
    
    // 计算派生统计
    if (current_stats.total_tasks_completed > 0) {
        current_stats.avg_execution_time = std::chrono::microseconds(
            current_stats.total_execution_time.count() / current_stats.total_tasks_completed);
        
        auto total_seconds = current_stats.total_execution_time.count() / 1000000.0;
        if (total_seconds > 0) {
            current_stats.tasks_per_second = current_stats.total_tasks_completed / total_seconds;
        }
    }
    
    // 计算队列中的任务数
    current_stats.tasks_in_queue = normal_queue_->size() + high_priority_queue_->size();
    
    // 计算活跃和空闲线程数
    current_stats.active_worker_threads = worker_threads_.size();
    current_stats.idle_worker_threads = 0; // 简化实现
    
    // 获取内存使用情况
    auto memory_stats = memory_manager_.getUnifiedStats();
    current_stats.memory_usage = memory_stats.total_memory_usage;
    if (memory_stats.total_memory_usage > current_stats.peak_memory_usage) {
        current_stats.peak_memory_usage = memory_stats.total_memory_usage;
    }
    
    return current_stats;
}

void TXAsyncProcessingFramework::resetStats() {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    stats_ = FrameworkStats{};
}

std::string TXAsyncProcessingFramework::generatePerformanceReport() const {
    auto stats = getStats();
    
    std::ostringstream report;
    report << "=== TXAsyncProcessingFramework 性能报告 ===\n";
    
    report << "\n📊 任务统计:\n";
    report << "  总提交任务: " << stats.total_tasks_submitted << "\n";
    report << "  总完成任务: " << stats.total_tasks_completed << "\n";
    report << "  总失败任务: " << stats.total_tasks_failed << "\n";
    report << "  队列中任务: " << stats.tasks_in_queue << "\n";
    
    report << "\n⚡ 性能指标:\n";
    report << "  平均执行时间: " << stats.avg_execution_time.count() << " μs\n";
    report << "  任务处理速率: " << std::fixed << std::setprecision(2) << stats.tasks_per_second << " 任务/秒\n";
    
    report << "\n🧵 线程统计:\n";
    report << "  活跃工作线程: " << stats.active_worker_threads << "\n";
    report << "  空闲工作线程: " << stats.idle_worker_threads << "\n";
    
    report << "\n💾 内存统计:\n";
    report << "  当前内存使用: " << (stats.memory_usage / 1024.0 / 1024.0) << " MB\n";
    report << "  峰值内存使用: " << (stats.peak_memory_usage / 1024.0 / 1024.0) << " MB\n";
    
    report << "\n🔄 工作窃取统计:\n";
    report << "  工作窃取事件: " << stats.work_stealing_events << "\n";
    report << "  优先级提升: " << stats.priority_promotions << "\n";
    
    return report.str();
}

double TXAsyncProcessingFramework::getCurrentTaskRate() const {
    auto stats = getStats();
    return stats.tasks_per_second;
}

TXResult<void> TXAsyncProcessingFramework::updateConfig(const FrameworkConfig& config) {
    if (state_.load() == FrameworkState::RUNNING) {
        return Err(TXErrorCode::Unknown, "Cannot update config while framework is running");
    }

    config_ = config;

    // 重新创建队列如果容量改变
    if (normal_queue_->size() != config_.task_queue_capacity) {
        normal_queue_ = std::make_unique<TXLockFreeQueue<std::unique_ptr<TXAsyncTask>>>(config_.task_queue_capacity);
    }

    if (high_priority_queue_->size() != config_.high_priority_queue_capacity) {
        high_priority_queue_ = std::make_unique<TXLockFreeQueue<std::unique_ptr<TXAsyncTask>>>(config_.high_priority_queue_capacity);
    }

    return Ok();
}

TXResult<void> TXAsyncProcessingFramework::adjustWorkerThreads(size_t new_count) {
    if (state_.load() == FrameworkState::RUNNING) {
        return Err(TXErrorCode::Unknown, "Cannot adjust threads while framework is running");
    }

    config_.worker_thread_count = new_count;
    return Ok();
}

bool TXAsyncProcessingFramework::enqueueTask(std::unique_ptr<TXAsyncTask> task) {
    if (state_.load() != FrameworkState::RUNNING) {
        return false;
    }
    
    // 检查内存限制
    if (config_.enable_memory_management) {
        auto memory_stats = memory_manager_.getUnifiedStats();
        size_t memory_limit = config_.memory_limit_mb * 1024 * 1024;
        if (memory_stats.total_memory_usage > memory_limit * config_.memory_pressure_threshold) {
            // 内存压力过大，拒绝新任务
            return false;
        }
    }
    
    bool enqueued = false;
    
    // 根据优先级选择队列
    if (config_.enable_priority_scheduling && task->getPriority() > 0) {
        enqueued = high_priority_queue_->enqueue(std::move(task));
    } else {
        enqueued = normal_queue_->enqueue(std::move(task));
    }
    
    if (enqueued) {
        std::lock_guard<std::mutex> lock(stats_mutex_);
        stats_.total_tasks_submitted++;
    }
    
    return enqueued;
}

void TXAsyncProcessingFramework::workerThreadFunction(size_t thread_id) {
    while (!should_stop_.load()) {
        // 检查是否暂停
        if (state_.load() == FrameworkState::PAUSED) {
            std::this_thread::sleep_for(config_.worker_idle_timeout);
            continue;
        }
        
        // 尝试获取任务
        std::unique_ptr<TXAsyncTask> task = dequeueTask();
        
        if (!task) {
            // 没有任务，短暂休眠
            std::this_thread::sleep_for(config_.worker_idle_timeout);
            continue;
        }
        
        // 执行任务
        auto start_time = std::chrono::steady_clock::now();
        
        try {
            task->execute();
            
            auto end_time = std::chrono::steady_clock::now();
            auto execution_time = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
            
            // 更新统计
            updateStats(*task, execution_time);
            
        } catch (const std::exception&) {
            // 任务执行失败
            std::lock_guard<std::mutex> lock(stats_mutex_);
            stats_.total_tasks_failed++;
        }
    }
}

std::unique_ptr<TXAsyncTask> TXAsyncProcessingFramework::dequeueTask() {
    std::unique_ptr<TXAsyncTask> task;
    
    // 优先从高优先级队列获取任务
    if (config_.enable_priority_scheduling && high_priority_queue_->dequeue(task)) {
        return task;
    }
    
    // 从普通队列获取任务
    if (normal_queue_->dequeue(task)) {
        return task;
    }
    
    // 工作窃取（简化实现）
    if (config_.enable_work_stealing) {
        // 这里可以实现更复杂的工作窃取逻辑
        std::lock_guard<std::mutex> lock(stats_mutex_);
        stats_.work_stealing_events++;
    }
    
    return nullptr;
}

void TXAsyncProcessingFramework::updateStats(const TXAsyncTask& task, std::chrono::microseconds execution_time) {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    
    stats_.total_tasks_completed++;
    stats_.total_execution_time += execution_time;
    
    // 更新内存统计
    if (config_.enable_memory_management) {
        auto memory_stats = memory_manager_.getUnifiedStats();
        stats_.memory_usage = memory_stats.total_memory_usage;
        if (memory_stats.total_memory_usage > stats_.peak_memory_usage) {
            stats_.peak_memory_usage = memory_stats.total_memory_usage;
        }
    }
    
    // 通知等待的线程
    completion_cv_.notify_all();
}

size_t TXAsyncProcessingFramework::getOptimalThreadCount() const {
    // 获取硬件并发数
    size_t hardware_threads = std::thread::hardware_concurrency();
    
    if (hardware_threads == 0) {
        hardware_threads = 4; // 默认值
    }
    
    // 根据任务类型调整
    // 对于I/O密集型任务，可以使用更多线程
    // 对于CPU密集型任务，使用硬件线程数
    return std::min(hardware_threads * 2, static_cast<size_t>(16)); // 最多16个线程
}

} // namespace TinaXlsx
