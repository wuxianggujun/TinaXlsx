//
// @file TXSmartMemoryManager.hpp
// @brief 智能内存管理器 - 监控、预警和自动清理
//

#pragma once

#include "TXChunkAllocator.hpp"
#include <memory>
#include <thread>
#include <atomic>
#include <condition_variable>
#include <functional>
#include <queue>
#include <chrono>

namespace TinaXlsx {

/**
 * @brief 内存监控配置
 */
struct MemoryMonitorConfig {
    size_t warning_threshold_mb = 3072;        // 3GB警告阈值
    size_t critical_threshold_mb = 3584;       // 3.5GB严重阈值
    size_t emergency_threshold_mb = 3840;      // 3.75GB紧急阈值
    
    std::chrono::milliseconds monitor_interval{1000};  // 监控间隔1秒
    std::chrono::milliseconds cleanup_interval{5000};  // 清理间隔5秒
    
    bool enable_auto_cleanup = true;           // 启用自动清理
    bool enable_memory_warnings = true;       // 启用内存警告
    bool enable_emergency_cleanup = true;     // 启用紧急清理
    
    double cleanup_target_ratio = 0.7;        // 清理目标比例（70%）
    size_t min_cleanup_size_mb = 100;         // 最小清理大小100MB
};

/**
 * @brief 内存事件类型
 */
enum class MemoryEventType {
    ALLOCATION,         // 分配事件
    DEALLOCATION,       // 释放事件
    WARNING,            // 警告事件
    CRITICAL,           // 严重事件
    EMERGENCY,          // 紧急事件
    CLEANUP_START,      // 清理开始
    CLEANUP_END,        // 清理结束
    LIMIT_EXCEEDED      // 超出限制
};

/**
 * @brief 内存事件
 */
struct MemoryEvent {
    MemoryEventType type;
    size_t memory_usage_mb;
    size_t memory_limit_mb;
    double usage_ratio;
    std::chrono::steady_clock::time_point timestamp;
    std::string message;
    
    MemoryEvent(MemoryEventType t, size_t usage, size_t limit, const std::string& msg = "")
        : type(t), memory_usage_mb(usage), memory_limit_mb(limit), 
          usage_ratio(static_cast<double>(usage) / limit),
          timestamp(std::chrono::steady_clock::now()), message(msg) {}
};

/**
 * @brief 内存清理策略
 */
class MemoryCleanupStrategy {
public:
    virtual ~MemoryCleanupStrategy() = default;
    
    /**
     * @brief 执行清理
     * @param allocator 内存分配器
     * @param target_reduction_mb 目标减少的内存量（MB）
     * @return 实际清理的内存量（MB）
     */
    virtual size_t cleanup(TXChunkAllocator& allocator, size_t target_reduction_mb) = 0;
    
    /**
     * @brief 获取策略名称
     */
    virtual std::string getName() const = 0;
    
    /**
     * @brief 估算可清理的内存量
     */
    virtual size_t estimateCleanupSize(const TXChunkAllocator& allocator) const = 0;
};

/**
 * @brief 压缩清理策略
 */
class CompactCleanupStrategy : public MemoryCleanupStrategy {
public:
    size_t cleanup(TXChunkAllocator& allocator, size_t target_reduction_mb) override;
    std::string getName() const override { return "Compact"; }
    size_t estimateCleanupSize(const TXChunkAllocator& allocator) const override;
};

/**
 * @brief 全量清理策略
 */
class FullCleanupStrategy : public MemoryCleanupStrategy {
public:
    size_t cleanup(TXChunkAllocator& allocator, size_t target_reduction_mb) override;
    std::string getName() const override { return "Full"; }
    size_t estimateCleanupSize(const TXChunkAllocator& allocator) const override;
};

/**
 * @brief 智能内存管理器
 */
class TXSmartMemoryManager {
public:
    /**
     * @brief 内存监控统计
     */
    struct MonitoringStats {
        size_t total_events = 0;
        size_t warning_events = 0;
        size_t critical_events = 0;
        size_t emergency_events = 0;
        size_t cleanup_events = 0;
        size_t total_cleanup_mb = 0;
        
        std::chrono::steady_clock::time_point start_time;
        std::chrono::steady_clock::time_point last_event_time;
        
        double avg_memory_usage = 0.0;
        size_t peak_memory_usage = 0;
        size_t current_memory_usage = 0;
    };
    
    explicit TXSmartMemoryManager(TXChunkAllocator& allocator,
                                 const MemoryMonitorConfig& config = MemoryMonitorConfig{});
    ~TXSmartMemoryManager();
    
    // ==================== 监控控制 ====================
    
    /**
     * @brief 启动内存监控
     */
    void startMonitoring();
    
    /**
     * @brief 停止内存监控
     */
    void stopMonitoring();
    
    /**
     * @brief 检查是否正在监控
     */
    bool isMonitoring() const { return monitoring_active_.load(); }
    
    // ==================== 事件处理 ====================
    
    /**
     * @brief 设置事件回调
     */
    using EventCallback = std::function<void(const MemoryEvent&)>;
    void setEventCallback(EventCallback callback) { event_callback_ = std::move(callback); }
    
    /**
     * @brief 手动触发内存检查
     */
    void checkMemoryStatus();
    
    /**
     * @brief 手动触发清理
     */
    size_t triggerCleanup(bool force = false);
    
    // ==================== 清理策略管理 ====================
    
    /**
     * @brief 添加清理策略
     */
    void addCleanupStrategy(std::unique_ptr<MemoryCleanupStrategy> strategy);
    
    /**
     * @brief 移除所有清理策略
     */
    void clearCleanupStrategies();
    
    /**
     * @brief 获取清理策略数量
     */
    size_t getCleanupStrategyCount() const { return cleanup_strategies_.size(); }
    
    // ==================== 配置管理 ====================
    
    /**
     * @brief 更新配置
     */
    void updateConfig(const MemoryMonitorConfig& config);
    
    /**
     * @brief 获取配置
     */
    const MemoryMonitorConfig& getConfig() const { return config_; }
    
    // ==================== 统计信息 ====================
    
    /**
     * @brief 获取监控统计
     */
    MonitoringStats getStats() const;
    
    /**
     * @brief 重置统计信息
     */
    void resetStats();
    
    /**
     * @brief 获取最近的事件
     */
    std::vector<MemoryEvent> getRecentEvents(size_t count = 100) const;
    
    /**
     * @brief 生成监控报告
     */
    std::string generateMonitoringReport() const;
    
    // ==================== 内存预测 ====================
    
    /**
     * @brief 预测内存使用趋势
     */
    struct MemoryTrend {
        double growth_rate_mb_per_sec = 0.0;
        std::chrono::seconds time_to_warning{0};
        std::chrono::seconds time_to_critical{0};
        std::chrono::seconds time_to_limit{0};
        bool is_growing = false;
    };
    
    MemoryTrend predictMemoryTrend() const;
    
    /**
     * @brief 检查是否需要预防性清理
     */
    bool shouldPreventiveCleanup() const;

private:
    // ==================== 内部数据 ====================
    
    TXChunkAllocator& allocator_;
    MemoryMonitorConfig config_;
    
    // 监控线程
    std::unique_ptr<std::thread> monitor_thread_;
    std::atomic<bool> monitoring_active_{false};
    std::atomic<bool> stop_requested_{false};
    
    // 清理策略
    std::vector<std::unique_ptr<MemoryCleanupStrategy>> cleanup_strategies_;
    
    // 事件处理
    EventCallback event_callback_;
    mutable std::mutex events_mutex_;
    std::queue<MemoryEvent> recent_events_;
    static constexpr size_t MAX_RECENT_EVENTS = 1000;
    
    // 统计信息
    mutable MonitoringStats stats_;
    mutable std::mutex stats_mutex_;
    
    // 内存使用历史（用于趋势分析）
    mutable std::mutex history_mutex_;
    std::queue<std::pair<std::chrono::steady_clock::time_point, size_t>> memory_history_;
    static constexpr size_t MAX_HISTORY_POINTS = 300; // 5分钟历史（1秒间隔）
    
    // ==================== 内部方法 ====================
    
    /**
     * @brief 监控线程主循环
     */
    void monitoringLoop();
    
    /**
     * @brief 处理内存事件
     */
    void handleMemoryEvent(const MemoryEvent& event);
    
    /**
     * @brief 执行清理策略
     */
    size_t executeCleanupStrategies(size_t target_reduction_mb);
    
    /**
     * @brief 更新统计信息
     */
    void updateStats(const MemoryEvent& event);
    
    /**
     * @brief 添加内存历史记录
     */
    void addMemoryHistoryPoint(size_t memory_usage);
    
    /**
     * @brief 计算内存增长率
     */
    double calculateGrowthRate() const;
    
    /**
     * @brief 检查阈值
     */
    MemoryEventType checkThresholds(size_t memory_usage_mb) const;
    
    /**
     * @brief 格式化内存大小
     */
    static std::string formatMemorySize(size_t size_mb);
    
    /**
     * @brief 格式化时间间隔
     */
    static std::string formatDuration(std::chrono::seconds duration);
};

/**
 * @brief 全局内存管理器实例
 */
class GlobalMemoryManager {
public:
    static TXSmartMemoryManager& getInstance();
    static void initialize(TXChunkAllocator& allocator, 
                          const MemoryMonitorConfig& config = MemoryMonitorConfig{});
    static void shutdown();
    
private:
    static std::unique_ptr<TXSmartMemoryManager> instance_;
    static std::mutex instance_mutex_;
};

} // namespace TinaXlsx
