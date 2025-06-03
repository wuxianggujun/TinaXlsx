//
// @file TXSmartMemoryManager.cpp
// @brief 智能内存管理器实现
//

#include "TinaXlsx/TXSmartMemoryManager.hpp"
#include <algorithm>
#include <sstream>
#include <iomanip>
#include <iostream>

namespace TinaXlsx {

// ==================== 清理策略实现 ====================

size_t CompactCleanupStrategy::cleanup(TXChunkAllocator& allocator, size_t target_reduction_mb) {
    size_t before_usage = allocator.getTotalMemoryUsage();
    allocator.compact();
    size_t after_usage = allocator.getTotalMemoryUsage();
    
    return (before_usage - after_usage) / (1024 * 1024);
}

size_t CompactCleanupStrategy::estimateCleanupSize(const TXChunkAllocator& allocator) const {
    auto chunk_infos = allocator.getChunkInfos();
    size_t fragmented_space = 0;
    
    for (const auto& info : chunk_infos) {
        if (!info.is_active) {
            fragmented_space += info.total_size;
        }
    }
    
    return fragmented_space / (1024 * 1024);
}

size_t FullCleanupStrategy::cleanup(TXChunkAllocator& allocator, size_t target_reduction_mb) {
    size_t before_usage = allocator.getTotalMemoryUsage();
    allocator.deallocateAll();
    size_t after_usage = allocator.getTotalMemoryUsage();
    
    return (before_usage - after_usage) / (1024 * 1024);
}

size_t FullCleanupStrategy::estimateCleanupSize(const TXChunkAllocator& allocator) const {
    return allocator.getTotalMemoryUsage() / (1024 * 1024);
}

// ==================== TXSmartMemoryManager 实现 ====================

TXSmartMemoryManager::TXSmartMemoryManager(TXChunkAllocator& allocator,
                                          const MemoryMonitorConfig& config)
    : allocator_(allocator), config_(config) {
    
    stats_.start_time = std::chrono::steady_clock::now();
    
    // 添加默认清理策略
    addCleanupStrategy(std::make_unique<CompactCleanupStrategy>());
    addCleanupStrategy(std::make_unique<FullCleanupStrategy>());
}

TXSmartMemoryManager::~TXSmartMemoryManager() {
    stopMonitoring();
}

void TXSmartMemoryManager::startMonitoring() {
    if (monitoring_active_.load()) {
        return; // 已经在监控
    }
    
    stop_requested_.store(false);
    monitoring_active_.store(true);
    
    monitor_thread_ = std::make_unique<std::thread>(&TXSmartMemoryManager::monitoringLoop, this);
    
    std::cout << "内存监控已启动" << std::endl;
}

void TXSmartMemoryManager::stopMonitoring() {
    if (!monitoring_active_.load()) {
        return; // 没有在监控
    }
    
    stop_requested_.store(true);
    monitoring_active_.store(false);
    
    if (monitor_thread_ && monitor_thread_->joinable()) {
        monitor_thread_->join();
    }
    
    monitor_thread_.reset();
    
    std::cout << "内存监控已停止" << std::endl;
}

void TXSmartMemoryManager::checkMemoryStatus() {
    size_t memory_usage = allocator_.getTotalMemoryUsage();
    size_t memory_usage_mb = memory_usage / (1024 * 1024);
    size_t memory_limit_mb = allocator_.getMemoryLimit() / (1024 * 1024);
    
    // 添加历史记录
    addMemoryHistoryPoint(memory_usage_mb);
    
    // 检查阈值
    MemoryEventType event_type = checkThresholds(memory_usage_mb);
    
    if (event_type != MemoryEventType::ALLOCATION) {
        std::string message = "内存使用: " + formatMemorySize(memory_usage_mb) + 
                             "/" + formatMemorySize(memory_limit_mb);
        
        MemoryEvent event(event_type, memory_usage_mb, memory_limit_mb, message);
        handleMemoryEvent(event);
        
        // 根据事件类型执行相应操作
        if (config_.enable_auto_cleanup) {
            if (event_type == MemoryEventType::EMERGENCY && config_.enable_emergency_cleanup) {
                triggerCleanup(true);
            } else if (event_type == MemoryEventType::CRITICAL || shouldPreventiveCleanup()) {
                triggerCleanup(false);
            }
        }
    }
}

size_t TXSmartMemoryManager::triggerCleanup(bool force) {
    size_t memory_usage_mb = allocator_.getTotalMemoryUsage() / (1024 * 1024);
    size_t memory_limit_mb = allocator_.getMemoryLimit() / (1024 * 1024);
    
    // 计算目标清理量
    size_t target_reduction_mb;
    if (force) {
        // 紧急清理：清理到安全水平
        target_reduction_mb = memory_usage_mb - static_cast<size_t>(memory_limit_mb * config_.cleanup_target_ratio);
    } else {
        // 常规清理：清理最小量
        target_reduction_mb = std::max(config_.min_cleanup_size_mb, 
                                      memory_usage_mb - config_.warning_threshold_mb);
    }
    
    if (target_reduction_mb < config_.min_cleanup_size_mb && !force) {
        return 0; // 不需要清理
    }
    
    // 发送清理开始事件
    MemoryEvent start_event(MemoryEventType::CLEANUP_START, memory_usage_mb, memory_limit_mb,
                           "开始清理，目标: " + formatMemorySize(target_reduction_mb));
    handleMemoryEvent(start_event);
    
    // 执行清理
    size_t actual_cleaned = executeCleanupStrategies(target_reduction_mb);
    
    // 发送清理结束事件
    size_t new_memory_usage_mb = allocator_.getTotalMemoryUsage() / (1024 * 1024);
    MemoryEvent end_event(MemoryEventType::CLEANUP_END, new_memory_usage_mb, memory_limit_mb,
                         "清理完成，释放: " + formatMemorySize(actual_cleaned));
    handleMemoryEvent(end_event);
    
    return actual_cleaned;
}

void TXSmartMemoryManager::addCleanupStrategy(std::unique_ptr<MemoryCleanupStrategy> strategy) {
    cleanup_strategies_.push_back(std::move(strategy));
}

void TXSmartMemoryManager::clearCleanupStrategies() {
    cleanup_strategies_.clear();
}

void TXSmartMemoryManager::updateConfig(const MemoryMonitorConfig& config) {
    config_ = config;
}

TXSmartMemoryManager::MonitoringStats TXSmartMemoryManager::getStats() const {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    
    MonitoringStats current_stats = stats_;
    current_stats.current_memory_usage = allocator_.getTotalMemoryUsage() / (1024 * 1024);
    
    return current_stats;
}

void TXSmartMemoryManager::resetStats() {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    stats_ = MonitoringStats{};
    stats_.start_time = std::chrono::steady_clock::now();
}

std::vector<MemoryEvent> TXSmartMemoryManager::getRecentEvents(size_t count) const {
    std::lock_guard<std::mutex> lock(events_mutex_);
    
    std::vector<MemoryEvent> events;
    auto temp_queue = recent_events_;
    
    while (!temp_queue.empty() && events.size() < count) {
        events.push_back(temp_queue.front());
        temp_queue.pop();
    }
    
    std::reverse(events.begin(), events.end()); // 最新的在前
    return events;
}

std::string TXSmartMemoryManager::generateMonitoringReport() const {
    std::ostringstream report;
    auto stats = getStats();
    auto trend = predictMemoryTrend();
    
    report << "=== TXSmartMemoryManager 监控报告 ===\n";
    report << "监控时间: " << formatDuration(
        std::chrono::duration_cast<std::chrono::seconds>(
            std::chrono::steady_clock::now() - stats.start_time)) << "\n";
    report << "当前内存使用: " << formatMemorySize(stats.current_memory_usage) << "\n";
    report << "峰值内存使用: " << formatMemorySize(stats.peak_memory_usage) << "\n";
    report << "平均内存使用: " << formatMemorySize(static_cast<size_t>(stats.avg_memory_usage)) << "\n";
    
    report << "\n事件统计:\n";
    report << "  总事件数: " << stats.total_events << "\n";
    report << "  警告事件: " << stats.warning_events << "\n";
    report << "  严重事件: " << stats.critical_events << "\n";
    report << "  紧急事件: " << stats.emergency_events << "\n";
    report << "  清理事件: " << stats.cleanup_events << "\n";
    report << "  总清理量: " << formatMemorySize(stats.total_cleanup_mb) << "\n";
    
    report << "\n内存趋势:\n";
    report << "  增长率: " << std::fixed << std::setprecision(2) 
           << trend.growth_rate_mb_per_sec << " MB/秒\n";
    report << "  是否增长: " << (trend.is_growing ? "是" : "否") << "\n";
    
    if (trend.is_growing) {
        report << "  到达警告阈值: " << formatDuration(trend.time_to_warning) << "\n";
        report << "  到达严重阈值: " << formatDuration(trend.time_to_critical) << "\n";
        report << "  到达内存限制: " << formatDuration(trend.time_to_limit) << "\n";
    }
    
    return report.str();
}

// ==================== 内存预测和趋势分析 ====================

TXSmartMemoryManager::MemoryTrend TXSmartMemoryManager::predictMemoryTrend() const {
    MemoryTrend trend;

    std::lock_guard<std::mutex> lock(history_mutex_);

    if (memory_history_.size() < 2) {
        return trend; // 数据不足
    }

    // 计算增长率
    trend.growth_rate_mb_per_sec = calculateGrowthRate();
    trend.is_growing = trend.growth_rate_mb_per_sec > 0.1; // 0.1 MB/s 以上认为是增长

    if (trend.is_growing) {
        size_t current_usage = memory_history_.back().second;

        // 计算到达各阈值的时间
        if (current_usage < config_.warning_threshold_mb) {
            double seconds_to_warning = (config_.warning_threshold_mb - current_usage) / trend.growth_rate_mb_per_sec;
            trend.time_to_warning = std::chrono::seconds(static_cast<long>(seconds_to_warning));
        }

        if (current_usage < config_.critical_threshold_mb) {
            double seconds_to_critical = (config_.critical_threshold_mb - current_usage) / trend.growth_rate_mb_per_sec;
            trend.time_to_critical = std::chrono::seconds(static_cast<long>(seconds_to_critical));
        }

        size_t memory_limit_mb = allocator_.getMemoryLimit() / (1024 * 1024);
        if (current_usage < memory_limit_mb) {
            double seconds_to_limit = (memory_limit_mb - current_usage) / trend.growth_rate_mb_per_sec;
            trend.time_to_limit = std::chrono::seconds(static_cast<long>(seconds_to_limit));
        }
    }

    return trend;
}

bool TXSmartMemoryManager::shouldPreventiveCleanup() const {
    auto trend = predictMemoryTrend();

    // 如果预测在5分钟内达到严重阈值，执行预防性清理
    return trend.is_growing && trend.time_to_critical.count() > 0 &&
           trend.time_to_critical < std::chrono::minutes(5);
}

// ==================== 私有方法实现 ====================

void TXSmartMemoryManager::monitoringLoop() {
    while (!stop_requested_.load()) {
        try {
            checkMemoryStatus();

            // 等待下一次检查
            std::this_thread::sleep_for(config_.monitor_interval);

        } catch (const std::exception& e) {
            std::cerr << "内存监控异常: " << e.what() << std::endl;
            std::this_thread::sleep_for(std::chrono::seconds(1));
        }
    }
}

void TXSmartMemoryManager::handleMemoryEvent(const MemoryEvent& event) {
    // 更新统计信息
    updateStats(event);

    // 添加到最近事件队列
    {
        std::lock_guard<std::mutex> lock(events_mutex_);
        recent_events_.push(event);

        // 保持队列大小
        while (recent_events_.size() > MAX_RECENT_EVENTS) {
            recent_events_.pop();
        }
    }

    // 调用事件回调
    if (event_callback_) {
        try {
            event_callback_(event);
        } catch (const std::exception& e) {
            std::cerr << "事件回调异常: " << e.what() << std::endl;
        }
    }

    // 输出警告信息
    if (config_.enable_memory_warnings) {
        switch (event.type) {
            case MemoryEventType::WARNING:
                std::cout << "⚠️  内存警告: " << event.message << std::endl;
                break;
            case MemoryEventType::CRITICAL:
                std::cout << "🔥 内存严重: " << event.message << std::endl;
                break;
            case MemoryEventType::EMERGENCY:
                std::cout << "🚨 内存紧急: " << event.message << std::endl;
                break;
            case MemoryEventType::CLEANUP_START:
                std::cout << "🧹 " << event.message << std::endl;
                break;
            case MemoryEventType::CLEANUP_END:
                std::cout << "✅ " << event.message << std::endl;
                break;
            default:
                break;
        }
    }
}

size_t TXSmartMemoryManager::executeCleanupStrategies(size_t target_reduction_mb) {
    size_t total_cleaned = 0;

    for (auto& strategy : cleanup_strategies_) {
        if (total_cleaned >= target_reduction_mb) {
            break; // 已达到目标
        }

        try {
            size_t remaining_target = target_reduction_mb - total_cleaned;
            size_t cleaned = strategy->cleanup(allocator_, remaining_target);
            total_cleaned += cleaned;

            if (cleaned > 0) {
                std::cout << "清理策略 " << strategy->getName()
                         << " 释放了 " << formatMemorySize(cleaned) << std::endl;
            }

        } catch (const std::exception& e) {
            std::cerr << "清理策略 " << strategy->getName()
                     << " 执行失败: " << e.what() << std::endl;
        }
    }

    return total_cleaned;
}

void TXSmartMemoryManager::updateStats(const MemoryEvent& event) {
    std::lock_guard<std::mutex> lock(stats_mutex_);

    stats_.total_events++;
    stats_.last_event_time = event.timestamp;
    stats_.current_memory_usage = event.memory_usage_mb;
    stats_.peak_memory_usage = std::max(stats_.peak_memory_usage, event.memory_usage_mb);

    // 更新平均内存使用
    if (stats_.total_events == 1) {
        stats_.avg_memory_usage = event.memory_usage_mb;
    } else {
        stats_.avg_memory_usage = (stats_.avg_memory_usage * (stats_.total_events - 1) +
                                  event.memory_usage_mb) / stats_.total_events;
    }

    // 更新事件计数
    switch (event.type) {
        case MemoryEventType::WARNING:
            stats_.warning_events++;
            break;
        case MemoryEventType::CRITICAL:
            stats_.critical_events++;
            break;
        case MemoryEventType::EMERGENCY:
            stats_.emergency_events++;
            break;
        case MemoryEventType::CLEANUP_END:
            stats_.cleanup_events++;
            // 从消息中提取清理量（简化实现）
            break;
        default:
            break;
    }
}

void TXSmartMemoryManager::addMemoryHistoryPoint(size_t memory_usage) {
    std::lock_guard<std::mutex> lock(history_mutex_);

    auto now = std::chrono::steady_clock::now();
    memory_history_.push({now, memory_usage});

    // 保持历史记录大小
    while (memory_history_.size() > MAX_HISTORY_POINTS) {
        memory_history_.pop();
    }
}

double TXSmartMemoryManager::calculateGrowthRate() const {
    if (memory_history_.size() < 2) {
        return 0.0;
    }

    // 使用最近10个点计算线性回归
    size_t points_to_use = std::min(memory_history_.size(), size_t(10));

    auto temp_queue = memory_history_;
    std::vector<std::pair<double, double>> points;

    // 提取最近的点 - 修复：检查队列是否为空
    while (temp_queue.size() > memory_history_.size() - points_to_use && !temp_queue.empty()) {
        temp_queue.pop();
    }

    if (temp_queue.empty()) {
        return 0.0; // 安全检查
    }

    auto start_time = temp_queue.front().first;
    while (!temp_queue.empty()) {
        auto point = temp_queue.front();
        temp_queue.pop();

        double seconds = std::chrono::duration<double>(point.first - start_time).count();
        double memory_mb = static_cast<double>(point.second);
        points.push_back({seconds, memory_mb});
    }

    // 简单线性回归计算斜率
    if (points.size() < 2) return 0.0;

    double sum_x = 0, sum_y = 0, sum_xy = 0, sum_x2 = 0;
    for (const auto& point : points) {
        sum_x += point.first;
        sum_y += point.second;
        sum_xy += point.first * point.second;
        sum_x2 += point.first * point.first;
    }

    double n = static_cast<double>(points.size());
    double denominator = n * sum_x2 - sum_x * sum_x;

    if (std::abs(denominator) < 1e-10) return 0.0;

    return (n * sum_xy - sum_x * sum_y) / denominator;
}

MemoryEventType TXSmartMemoryManager::checkThresholds(size_t memory_usage_mb) const {
    if (memory_usage_mb >= config_.emergency_threshold_mb) {
        return MemoryEventType::EMERGENCY;
    } else if (memory_usage_mb >= config_.critical_threshold_mb) {
        return MemoryEventType::CRITICAL;
    } else if (memory_usage_mb >= config_.warning_threshold_mb) {
        return MemoryEventType::WARNING;
    }

    return MemoryEventType::ALLOCATION;
}

std::string TXSmartMemoryManager::formatMemorySize(size_t size_mb) {
    if (size_mb >= 1024) {
        return std::to_string(size_mb / 1024) + "." +
               std::to_string((size_mb % 1024) / 102) + " GB";
    } else {
        return std::to_string(size_mb) + " MB";
    }
}

std::string TXSmartMemoryManager::formatDuration(std::chrono::seconds duration) {
    auto hours = std::chrono::duration_cast<std::chrono::hours>(duration);
    auto minutes = std::chrono::duration_cast<std::chrono::minutes>(duration - hours);
    auto seconds = duration - hours - minutes;

    std::ostringstream oss;
    if (hours.count() > 0) {
        oss << hours.count() << "h ";
    }
    if (minutes.count() > 0) {
        oss << minutes.count() << "m ";
    }
    oss << seconds.count() << "s";

    return oss.str();
}

// ==================== 全局内存管理器 ====================

std::unique_ptr<TXSmartMemoryManager> GlobalMemoryManager::instance_;
std::mutex GlobalMemoryManager::instance_mutex_;

TXSmartMemoryManager& GlobalMemoryManager::getInstance() {
    std::lock_guard<std::mutex> lock(instance_mutex_);
    if (!instance_) {
        throw std::runtime_error("GlobalMemoryManager not initialized");
    }
    return *instance_;
}

void GlobalMemoryManager::initialize(TXChunkAllocator& allocator,
                                    const MemoryMonitorConfig& config) {
    std::lock_guard<std::mutex> lock(instance_mutex_);
    instance_ = std::make_unique<TXSmartMemoryManager>(allocator, config);
}

void GlobalMemoryManager::shutdown() {
    std::lock_guard<std::mutex> lock(instance_mutex_);
    instance_.reset();
}

} // namespace TinaXlsx
