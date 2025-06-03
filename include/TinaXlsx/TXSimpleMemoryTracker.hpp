//
// @file TXSimpleMemoryTracker.hpp
// @brief 简化的内存跟踪器 - 轻量级内存监控
//

#pragma once

#include <chrono>
#include <atomic>
#include <string>

namespace TinaXlsx {

/**
 * @brief 🚀 简化的内存跟踪器
 * 
 * 轻量级内存监控，专注于：
 * - 系统内存使用监控
 * - 作用域内存跟踪
 * - 简单的泄漏检测
 */
class TXSimpleMemoryTracker {
public:
    /**
     * @brief 内存快照
     */
    struct MemorySnapshot {
        size_t systemMemory = 0;        // 系统内存使用
        size_t workingSet = 0;          // 工作集大小
        size_t virtualMemory = 0;       // 虚拟内存使用
        std::chrono::steady_clock::time_point timestamp;
        
        MemorySnapshot() : timestamp(std::chrono::steady_clock::now()) {}
    };
    
    /**
     * @brief 获取当前内存快照
     */
    static MemorySnapshot getCurrentSnapshot();
    
    /**
     * @brief 计算内存差异
     */
    static MemorySnapshot calculateDifference(const MemorySnapshot& end, const MemorySnapshot& start);
    
    /**
     * @brief 格式化内存大小
     */
    static std::string formatMemorySize(size_t bytes);
    
    /**
     * @brief 检查是否可能存在内存泄漏
     * @param growth 内存增长量
     * @param baseline 基线内存
     * @param threshold 泄漏阈值（百分比，默认10%）
     */
    static bool isPossibleLeak(size_t growth, size_t baseline, double threshold = 0.1);

private:
    static size_t getSystemMemoryUsage();
    static size_t getWorkingSetSize();
    static size_t getVirtualMemoryUsage();
};

/**
 * @brief 🚀 RAII作用域内存监控器
 * 
 * 自动监控作用域内的内存变化
 */
class TXScopeMemoryMonitor {
public:
    explicit TXScopeMemoryMonitor(const std::string& name = "ScopeMonitor");
    ~TXScopeMemoryMonitor();
    
    /**
     * @brief 获取当前内存变化
     */
    TXSimpleMemoryTracker::MemorySnapshot getCurrentDifference() const;
    
    /**
     * @brief 检查是否有内存增长
     */
    bool hasMemoryGrowth() const;
    
    /**
     * @brief 获取内存增长量
     */
    size_t getMemoryGrowth() const;
    
    /**
     * @brief 设置是否在析构时输出报告
     */
    void setReportOnDestroy(bool enable) { reportOnDestroy_ = enable; }

private:
    std::string name_;
    TXSimpleMemoryTracker::MemorySnapshot startSnapshot_;
    bool reportOnDestroy_;
};

/**
 * @brief 🚀 内存使用统计器
 * 
 * 收集和分析内存使用模式
 */
class TXMemoryUsageStats {
public:
    static TXMemoryUsageStats& instance();
    
    /**
     * @brief 记录内存使用点
     */
    void recordUsage(const std::string& operation, size_t memoryUsage);
    
    /**
     * @brief 开始操作监控
     */
    void startOperation(const std::string& operation);
    
    /**
     * @brief 结束操作监控
     */
    void endOperation(const std::string& operation);
    
    /**
     * @brief 获取操作统计
     */
    struct OperationStats {
        std::string name;
        size_t count = 0;
        size_t totalMemoryUsed = 0;
        size_t averageMemoryUsed = 0;
        size_t peakMemoryUsed = 0;
        double averageTimeMs = 0.0;
    };
    
    std::vector<OperationStats> getOperationStats() const;
    
    /**
     * @brief 重置统计信息
     */
    void reset();
    
    /**
     * @brief 生成内存使用报告
     */
    std::string generateReport() const;

private:
    TXMemoryUsageStats() = default;
    
    struct OperationRecord {
        std::string name;
        size_t memoryUsage;
        std::chrono::steady_clock::time_point startTime;
        std::chrono::steady_clock::time_point endTime;
    };
    
    std::vector<OperationRecord> records_;
    std::map<std::string, std::chrono::steady_clock::time_point> activeOperations_;
    mutable std::mutex mutex_;
};

/**
 * @brief 🚀 自动内存监控器
 * 
 * 自动监控TinaXlsx操作的内存使用
 */
class TXAutoMemoryMonitor {
public:
    /**
     * @brief 监控配置
     */
    struct MonitorConfig {
        bool enableAutoMonitoring = true;      // 启用自动监控
        bool enableOperationTracking = true;   // 启用操作跟踪
        bool enableLeakDetection = true;       // 启用泄漏检测
        double leakThreshold = 0.1;            // 泄漏阈值（10%）
        size_t maxRecords = 1000;              // 最大记录数
    };
    
    static TXAutoMemoryMonitor& instance();
    
    /**
     * @brief 设置监控配置
     */
    void setConfig(const MonitorConfig& config) { config_ = config; }
    
    /**
     * @brief 开始监控工作簿操作
     */
    void startWorkbookOperation(const std::string& operation);
    
    /**
     * @brief 结束监控工作簿操作
     */
    void endWorkbookOperation(const std::string& operation);
    
    /**
     * @brief 检查内存健康状况
     */
    struct MemoryHealth {
        bool isHealthy = true;
        std::vector<std::string> warnings;
        std::vector<std::string> recommendations;
        TXSimpleMemoryTracker::MemorySnapshot currentSnapshot;
    };
    
    MemoryHealth checkMemoryHealth() const;
    
    /**
     * @brief 生成内存健康报告
     */
    std::string generateHealthReport() const;

private:
    TXAutoMemoryMonitor() = default;
    
    MonitorConfig config_;
    TXSimpleMemoryTracker::MemorySnapshot baselineSnapshot_;
    std::map<std::string, TXSimpleMemoryTracker::MemorySnapshot> operationSnapshots_;
    mutable std::mutex mutex_;
};

} // namespace TinaXlsx

// 便利宏
#define TX_MEMORY_SCOPE_MONITOR(name) TinaXlsx::TXScopeMemoryMonitor _mem_monitor(name)
#define TX_MEMORY_OPERATION_START(op) TinaXlsx::TXAutoMemoryMonitor::instance().startWorkbookOperation(op)
#define TX_MEMORY_OPERATION_END(op) TinaXlsx::TXAutoMemoryMonitor::instance().endWorkbookOperation(op)
