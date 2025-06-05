//
// @file TXAsyncProcessingFramework.hpp
// @brief 异步处理框架 - 高性能异步任务处理
//

#pragma once

#include "TXUnifiedMemoryManager.hpp"
#include "TXResult.hpp"
#include <functional>
#include <future>
#include <queue>
#include <thread>
#include <atomic>
#include <mutex>
#include <condition_variable>
#include <memory>
#include <vector>
#include <chrono>
#include <type_traits>

#ifdef _WIN32
#include <malloc.h>
#endif

namespace TinaXlsx {

/**
 * @brief 异步任务接口
 */
class TXAsyncTask {
public:
    virtual ~TXAsyncTask() = default;
    virtual void execute() = 0;
    virtual std::string getTaskName() const = 0;
    virtual size_t getEstimatedMemoryUsage() const { return 0; }
    virtual int getPriority() const { return 0; } // 0=normal, >0=high, <0=low
};

/**
 * @brief 函数式异步任务
 */
template<typename Func>
class TXFunctionTask : public TXAsyncTask {
public:
    explicit TXFunctionTask(Func&& func, const std::string& name = "FunctionTask")
        : func_(std::forward<Func>(func)), name_(name) {}

    void execute() override {
        if constexpr (std::is_void_v<std::invoke_result_t<Func>>) {
            func_();
        } else {
            result_ = func_();
        }
    }

    std::string getTaskName() const override { return name_; }

    // 获取结果（仅对非void返回类型有效）
    template<typename T = std::invoke_result_t<Func>>
    std::enable_if_t<!std::is_void_v<T>, T> getResult() const {
        return result_;
    }

private:
    Func func_;
    std::string name_;
    std::conditional_t<std::is_void_v<std::invoke_result_t<Func>>, int, std::invoke_result_t<Func>> result_;
};

/**
 * @brief 带结果接口的任务基类
 */
template<typename T>
class TXResultTask : public TXAsyncTask {
public:
    virtual ~TXResultTask() = default;
    virtual T getResult() const = 0;
};

/**
 * @brief 返回类型特化的函数式异步任务
 */
template<typename ReturnType, typename Func>
class TXTypedFunctionTask : public TXResultTask<ReturnType> {
public:
    explicit TXTypedFunctionTask(Func&& func, const std::string& name = "TypedFunctionTask")
        : func_(std::forward<Func>(func)), name_(name) {}

    void execute() override {
        if constexpr (std::is_void_v<ReturnType>) {
            func_();
        } else {
            result_ = func_();
        }
    }

    std::string getTaskName() const override { return name_; }

    // 获取结果（仅对非void返回类型有效）
    ReturnType getResult() const override {
        if constexpr (std::is_void_v<ReturnType>) {
            return;
        } else {
            return result_;
        }
    }

private:
    Func func_;
    std::string name_;
    std::conditional_t<std::is_void_v<ReturnType>, int, ReturnType> result_;
};

/**
 * @brief 无锁队列（用于高性能任务调度）
 */
template<typename T>
class TXLockFreeQueue {
public:
    explicit TXLockFreeQueue(size_t capacity = 1024) : capacity_(capacity) {
        // Windows兼容的内存分配
#ifdef _WIN32
        buffer_ = static_cast<Node*>(_aligned_malloc(sizeof(Node) * capacity_, alignof(Node)));
#else
        buffer_ = static_cast<Node*>(std::aligned_alloc(alignof(Node), sizeof(Node) * capacity_));
#endif
        for (size_t i = 0; i < capacity_; ++i) {
            new (&buffer_[i]) Node();
        }
    }
    
    ~TXLockFreeQueue() {
        for (size_t i = 0; i < capacity_; ++i) {
            buffer_[i].~Node();
        }
#ifdef _WIN32
        _aligned_free(buffer_);
#else
        std::free(buffer_);
#endif
    }
    
    bool enqueue(T&& item) {
        size_t tail = tail_.load(std::memory_order_relaxed);
        size_t next_tail = (tail + 1) % capacity_;
        
        if (next_tail == head_.load(std::memory_order_acquire)) {
            return false; // Queue is full
        }
        
        buffer_[tail].data = std::move(item);
        tail_.store(next_tail, std::memory_order_release);
        return true;
    }
    
    bool dequeue(T& item) {
        size_t head = head_.load(std::memory_order_relaxed);
        
        if (head == tail_.load(std::memory_order_acquire)) {
            return false; // Queue is empty
        }
        
        item = std::move(buffer_[head].data);
        head_.store((head + 1) % capacity_, std::memory_order_release);
        return true;
    }
    
    size_t size() const {
        size_t tail = tail_.load(std::memory_order_acquire);
        size_t head = head_.load(std::memory_order_acquire);
        return (tail >= head) ? (tail - head) : (capacity_ - head + tail);
    }
    
    bool empty() const {
        return head_.load(std::memory_order_acquire) == tail_.load(std::memory_order_acquire);
    }

private:
    struct Node {
        T data;
    };
    
    Node* buffer_;
    size_t capacity_;
    std::atomic<size_t> head_{0};
    std::atomic<size_t> tail_{0};
};

/**
 * @brief 异步处理框架
 */
class TXAsyncProcessingFramework {
public:
    /**
     * @brief 框架配置
     */
    struct FrameworkConfig {
        size_t worker_thread_count = 0;          // 工作线程数(0=auto)
        size_t task_queue_capacity = 10000;      // 任务队列容量
        size_t high_priority_queue_capacity = 1000; // 高优先级队列容量
        
        bool enable_work_stealing = true;        // 启用工作窃取
        bool enable_priority_scheduling = true;  // 启用优先级调度
        bool enable_memory_management = true;    // 启用内存管理
        bool enable_performance_monitoring = true; // 启用性能监控
        
        std::chrono::milliseconds worker_idle_timeout{100}; // 工作线程空闲超时
        std::chrono::milliseconds shutdown_timeout{5000};   // 关闭超时
        
        size_t memory_limit_mb = 1024;           // 内存限制1GB
        double memory_pressure_threshold = 0.8;  // 内存压力阈值
    };
    
    /**
     * @brief 框架统计
     */
    struct FrameworkStats {
        size_t total_tasks_submitted = 0;        // 总提交任务数
        size_t total_tasks_completed = 0;        // 总完成任务数
        size_t total_tasks_failed = 0;           // 总失败任务数
        size_t tasks_in_queue = 0;               // 队列中任务数
        
        std::chrono::microseconds total_execution_time{0}; // 总执行时间
        std::chrono::microseconds avg_execution_time{0};   // 平均执行时间
        double tasks_per_second = 0.0;           // 任务处理速率
        
        size_t active_worker_threads = 0;        // 活跃工作线程数
        size_t idle_worker_threads = 0;          // 空闲工作线程数
        
        size_t memory_usage = 0;                 // 内存使用量
        size_t peak_memory_usage = 0;            // 峰值内存使用
        
        size_t work_stealing_events = 0;         // 工作窃取事件数
        size_t priority_promotions = 0;          // 优先级提升数
    };
    
    explicit TXAsyncProcessingFramework(TXUnifiedMemoryManager& memory_manager,
                                       const FrameworkConfig& config = FrameworkConfig{});
    ~TXAsyncProcessingFramework();
    
    // ==================== 框架控制 ====================
    
    /**
     * @brief 启动框架
     */
    TXResult<void> start();
    
    /**
     * @brief 停止框架
     */
    TXResult<void> stop();
    
    /**
     * @brief 暂停处理
     */
    TXResult<void> pause();
    
    /**
     * @brief 恢复处理
     */
    TXResult<void> resume();
    
    /**
     * @brief 获取框架状态
     */
    enum class FrameworkState {
        STOPPED,
        STARTING,
        RUNNING,
        PAUSED,
        STOPPING
    };
    
    FrameworkState getState() const { return state_.load(); }
    
    // ==================== 任务提交 ====================
    
    /**
     * @brief 提交任务
     */
    template<typename TaskType>
    TXResult<std::future<void>> submitTask(std::unique_ptr<TaskType> task) {
        static_assert(std::is_base_of_v<TXAsyncTask, TaskType>, "TaskType must inherit from TXAsyncTask");

        if (state_.load() != FrameworkState::RUNNING) {
            return Err<std::future<void>>(TXErrorCode::Unknown, "Framework is not running");
        }

        auto promise = std::make_shared<std::promise<void>>();
        auto future = promise->get_future();

        auto wrapper = std::make_unique<TaskWrapper>(std::move(task), promise);

        if (!enqueueTask(std::move(wrapper))) {
            return Err<std::future<void>>(TXErrorCode::Unknown, "Failed to enqueue task");
        }

        return Ok(std::move(future));
    }
    
    /**
     * @brief 提交函数任务
     */
    template<typename Func>
    TXResult<std::future<std::invoke_result_t<Func>>> submitFunction(Func&& func, const std::string& name = "Function") {
        using ReturnType = std::invoke_result_t<Func>;

        if (state_.load() != FrameworkState::RUNNING) {
            return Err<std::future<ReturnType>>(TXErrorCode::Unknown, "Framework is not running");
        }

        auto promise = std::make_shared<std::promise<ReturnType>>();
        auto future = promise->get_future();

        auto task = std::make_unique<TXTypedFunctionTask<ReturnType, Func>>(std::forward<Func>(func), name);
        auto wrapper = std::make_unique<FunctionTaskWrapper<ReturnType>>(task.release(), promise);

        if (!enqueueTask(std::move(wrapper))) {
            return Err<std::future<ReturnType>>(TXErrorCode::Unknown, "Failed to enqueue task");
        }

        return Ok(std::move(future));
    }
    
    /**
     * @brief 批量提交任务
     */
    template<typename TaskType>
    TXResult<std::vector<std::future<void>>> submitTasks(std::vector<std::unique_ptr<TaskType>> tasks) {
        std::vector<std::future<void>> futures;
        futures.reserve(tasks.size());

        for (auto& task : tasks) {
            auto result = submitTask(std::move(task));
            if (!result.isOk()) {
                return Err<std::vector<std::future<void>>>(result.error());
            }
            futures.push_back(result.value());
        }

        return Ok(std::move(futures));
    }
    
    // ==================== 等待和同步 ====================
    
    /**
     * @brief 等待所有任务完成
     */
    TXResult<void> waitForAll(std::chrono::milliseconds timeout = std::chrono::milliseconds(30000));
    
    /**
     * @brief 等待指定数量的任务完成
     */
    TXResult<void> waitForCount(size_t count, std::chrono::milliseconds timeout = std::chrono::milliseconds(30000));
    
    /**
     * @brief 清空任务队列
     */
    size_t clearQueue();
    
    // ==================== 统计和监控 ====================
    
    /**
     * @brief 获取统计信息
     */
    FrameworkStats getStats() const;
    
    /**
     * @brief 重置统计信息
     */
    void resetStats();
    
    /**
     * @brief 生成性能报告
     */
    std::string generatePerformanceReport() const;
    
    /**
     * @brief 获取当前任务处理速率
     */
    double getCurrentTaskRate() const;
    
    // ==================== 配置管理 ====================
    
    /**
     * @brief 更新配置
     */
    TXResult<void> updateConfig(const FrameworkConfig& config);
    
    /**
     * @brief 获取配置
     */
    const FrameworkConfig& getConfig() const { return config_; }
    
    /**
     * @brief 动态调整工作线程数
     */
    TXResult<void> adjustWorkerThreads(size_t new_count);

private:
    // ==================== 内部数据结构 ====================
    
    class TaskWrapper : public TXAsyncTask {
    public:
        TaskWrapper(std::unique_ptr<TXAsyncTask> task, std::shared_ptr<std::promise<void>> promise)
            : task_(std::move(task)), promise_(promise) {}
        
        void execute() override {
            try {
                task_->execute();
                promise_->set_value();
            } catch (...) {
                promise_->set_exception(std::current_exception());
            }
        }
        
        std::string getTaskName() const override { return task_->getTaskName(); }
        size_t getEstimatedMemoryUsage() const override { return task_->getEstimatedMemoryUsage(); }
        int getPriority() const override { return task_->getPriority(); }
        
    private:
        std::unique_ptr<TXAsyncTask> task_;
        std::shared_ptr<std::promise<void>> promise_;
    };
    
    template<typename T>
    class FunctionTaskWrapper : public TXAsyncTask {
    public:
        FunctionTaskWrapper(TXResultTask<T>* task, std::shared_ptr<std::promise<T>> promise)
            : task_(task), promise_(promise) {}

        ~FunctionTaskWrapper() {
            delete task_;
        }

        void execute() override {
            try {
                task_->execute();
                if constexpr (std::is_void_v<T>) {
                    promise_->set_value();
                } else {
                    promise_->set_value(task_->getResult());
                }
            } catch (...) {
                promise_->set_exception(std::current_exception());
            }
        }

        std::string getTaskName() const override { return task_->getTaskName(); }
        size_t getEstimatedMemoryUsage() const override { return task_->getEstimatedMemoryUsage(); }
        int getPriority() const override { return task_->getPriority(); }

    private:
        TXResultTask<T>* task_;
        std::shared_ptr<std::promise<T>> promise_;
    };
    
    TXUnifiedMemoryManager& memory_manager_;
    FrameworkConfig config_;
    std::atomic<FrameworkState> state_{FrameworkState::STOPPED};
    
    // 任务队列
    std::unique_ptr<TXLockFreeQueue<std::unique_ptr<TXAsyncTask>>> normal_queue_;
    std::unique_ptr<TXLockFreeQueue<std::unique_ptr<TXAsyncTask>>> high_priority_queue_;
    
    // 工作线程
    std::vector<std::thread> worker_threads_;
    std::atomic<bool> should_stop_{false};
    
    // 统计信息
    mutable FrameworkStats stats_;
    mutable std::mutex stats_mutex_;
    
    // 同步原语
    std::condition_variable completion_cv_;
    std::mutex completion_mutex_;
    
    // ==================== 内部方法 ====================
    
    bool enqueueTask(std::unique_ptr<TXAsyncTask> task);
    void workerThreadFunction(size_t thread_id);
    std::unique_ptr<TXAsyncTask> dequeueTask();
    void updateStats(const TXAsyncTask& task, std::chrono::microseconds execution_time);
    size_t getOptimalThreadCount() const;
};

} // namespace TinaXlsx
