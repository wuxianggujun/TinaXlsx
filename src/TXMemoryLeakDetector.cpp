//
// @file TXMemoryLeakDetector.cpp
// @brief 内存泄漏检测器实现
//

#include "TinaXlsx/TXMemoryLeakDetector.hpp"
#include "TinaXlsx/TXMemoryPool.hpp"
#include <algorithm>
#include <thread>

namespace TinaXlsx {

// ==================== TXMemoryLeakDetector 实现 ====================

TXMemoryLeakDetector& TXMemoryLeakDetector::instance() {
    static TXMemoryLeakDetector instance;
    return instance;
}

TXMemoryLeakDetector::~TXMemoryLeakDetector() {
    stopAutoCleanup();
}

void TXMemoryLeakDetector::setConfig(const DetectorConfig& config) {
    std::lock_guard<std::mutex> lock(mutex_);
    config_ = config;
}

void TXMemoryLeakDetector::recordAllocation(void* ptr, size_t size, const char* file, 
                                          int line, const char* function) {
    if (!config_.enableTracking || !ptr) return;
    
    std::lock_guard<std::mutex> lock(mutex_);
    
    // 检查是否超过最大跟踪数量
    if (allocations_.size() >= config_.maxAllocations) {
        cleanupOldAllocations();
    }
    
    AllocationInfo info;
    info.size = size;
    info.timestamp = std::chrono::steady_clock::now();
    info.file = file;
    info.line = line;
    info.function = function;
    
    allocations_[ptr] = info;
    
    // 更新统计信息
    totalAllocations_++;
    totalBytes_ += size;
    currentAllocations_++;
    currentBytes_ += size;
    
    updatePeakStats();
}

void TXMemoryLeakDetector::recordDeallocation(void* ptr) {
    if (!config_.enableTracking || !ptr) return;
    
    std::lock_guard<std::mutex> lock(mutex_);
    
    auto it = allocations_.find(ptr);
    if (it != allocations_.end()) {
        size_t size = it->second.size;
        allocations_.erase(it);
        
        // 更新统计信息
        totalDeallocations_++;
        currentAllocations_--;
        currentBytes_ -= size;
    }
}

TXMemoryLeakDetector::LeakReport TXMemoryLeakDetector::detectLeaks() const {
    std::lock_guard<std::mutex> lock(mutex_);
    
    LeakReport report;
    auto now = std::chrono::steady_clock::now();
    
    for (const auto& [ptr, info] : allocations_) {
        // 检查是否为长期未释放的内存
        auto duration = std::chrono::duration_cast<std::chrono::seconds>(now - info.timestamp);
        if (duration.count() > 60) { // 超过60秒未释放
            report.leaks.emplace_back(ptr, info);
            report.totalLeakedBytes += info.size;
            report.leakedAllocations++;
        }
    }
    
    return report;
}

void TXMemoryLeakDetector::forceCleanup() {
    std::lock_guard<std::mutex> lock(mutex_);
    allocations_.clear();
    
    currentAllocations_ = 0;
    currentBytes_ = 0;
}

TXMemoryLeakDetector::MemoryStats TXMemoryLeakDetector::getStats() const {
    std::lock_guard<std::mutex> lock(mutex_);
    
    MemoryStats stats;
    stats.currentAllocations = currentAllocations_.load();
    stats.currentBytes = currentBytes_.load();
    stats.totalAllocations = totalAllocations_.load();
    stats.totalDeallocations = totalDeallocations_.load();
    stats.totalBytes = totalBytes_.load();
    stats.peakAllocations = peakAllocations_.load();
    stats.peakBytes = peakBytes_.load();
    
    return stats;
}

void TXMemoryLeakDetector::reset() {
    std::lock_guard<std::mutex> lock(mutex_);
    
    allocations_.clear();
    
    currentAllocations_ = 0;
    currentBytes_ = 0;
    totalAllocations_ = 0;
    totalDeallocations_ = 0;
    totalBytes_ = 0;
    peakAllocations_ = 0;
    peakBytes_ = 0;
}

void TXMemoryLeakDetector::startAutoCleanup() {
    if (cleanupThread_) return;
    
    stopCleanup_ = false;
    cleanupThread_ = std::make_unique<std::thread>([this]() {
        while (!stopCleanup_) {
            std::this_thread::sleep_for(config_.cleanupInterval);
            
            if (config_.enableAutoCleanup) {
                cleanupOldAllocations();
            }
        }
    });
}

void TXMemoryLeakDetector::stopAutoCleanup() {
    if (!cleanupThread_) return;
    
    stopCleanup_ = true;
    if (cleanupThread_->joinable()) {
        cleanupThread_->join();
    }
    cleanupThread_.reset();
}

void TXMemoryLeakDetector::cleanupOldAllocations() {
    auto now = std::chrono::steady_clock::now();
    
    auto it = allocations_.begin();
    while (it != allocations_.end()) {
        auto duration = std::chrono::duration_cast<std::chrono::minutes>(now - it->second.timestamp);
        if (duration.count() > 10) { // 清理10分钟前的分配
            currentAllocations_--;
            currentBytes_ -= it->second.size;
            it = allocations_.erase(it);
        } else {
            ++it;
        }
    }
}

void TXMemoryLeakDetector::updatePeakStats() {
    size_t current = currentAllocations_.load();
    size_t peak = peakAllocations_.load();
    while (current > peak && !peakAllocations_.compare_exchange_weak(peak, current)) {
        peak = peakAllocations_.load();
    }
    
    size_t currentBytes = currentBytes_.load();
    size_t peakBytes = peakBytes_.load();
    while (currentBytes > peakBytes && !peakBytes_.compare_exchange_weak(peakBytes, currentBytes)) {
        peakBytes = peakBytes_.load();
    }
}

// ==================== TXScopedMemoryTracker 实现 ====================

TXScopedMemoryTracker::TXScopedMemoryTracker(const char* name) 
    : name_(name), startTime_(std::chrono::steady_clock::now()) {
    initialStats_ = TXMemoryLeakDetector::instance().getStats();
}

TXScopedMemoryTracker::~TXScopedMemoryTracker() {
    auto endStats = TXMemoryLeakDetector::instance().getStats();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(
        std::chrono::steady_clock::now() - startTime_);
    
    // 可以在这里输出作用域内存统计
    // printf("[%s] 内存变化: %zu -> %zu bytes, 耗时: %lldms\n", 
    //        name_, initialStats_.currentBytes, endStats.currentBytes, duration.count());
}

TXMemoryLeakDetector::MemoryStats TXScopedMemoryTracker::getScopeStats() const {
    auto currentStats = TXMemoryLeakDetector::instance().getStats();
    
    TXMemoryLeakDetector::MemoryStats scopeStats;
    scopeStats.currentAllocations = currentStats.currentAllocations - initialStats_.currentAllocations;
    scopeStats.currentBytes = currentStats.currentBytes - initialStats_.currentBytes;
    scopeStats.totalAllocations = currentStats.totalAllocations - initialStats_.totalAllocations;
    scopeStats.totalDeallocations = currentStats.totalDeallocations - initialStats_.totalDeallocations;
    
    return scopeStats;
}

bool TXScopedMemoryTracker::hasLeaks() const {
    auto scopeStats = getScopeStats();
    return scopeStats.currentBytes > 0; // 作用域结束时还有未释放的内存
}

// ==================== TXSmartMemoryManager 实现 ====================

TXSmartMemoryManager& TXSmartMemoryManager::instance() {
    static TXSmartMemoryManager instance;
    return instance;
}

TXSmartMemoryManager::TXSmartMemoryManager() {
    generalPool_ = std::make_unique<TXMemoryPool>();
    stringPool_ = std::make_unique<TXStringPool>();
}

TXSmartMemoryManager::~TXSmartMemoryManager() {
    emergencyCleanup();
}

void* TXSmartMemoryManager::allocate(size_t size, size_t alignment) {
    void* ptr = nullptr;
    
    if (shouldUsePool(size)) {
        ptr = allocateFromPool(size);
        poolAllocations_++;
    } else {
        ptr = allocateFromSystem(size);
        systemAllocations_++;
    }
    
    // 记录分配
    TXMemoryLeakDetector::instance().recordAllocation(ptr, size);
    
    return ptr;
}

void TXSmartMemoryManager::deallocate(void* ptr) {
    if (!ptr) return;
    
    // 记录释放
    TXMemoryLeakDetector::instance().recordDeallocation(ptr);
    
    if (generalPool_->isFromPool(ptr)) {
        deallocateToPool(ptr);
    } else {
        deallocateToSystem(ptr);
    }
}

TXSmartMemoryManager::HealthReport TXSmartMemoryManager::performHealthCheck() {
    HealthReport report;
    
    // 检查内存泄漏
    auto leakReport = TXMemoryLeakDetector::instance().detectLeaks();
    report.hasLeaks = leakReport.leakedAllocations > 0;
    report.leakedBytes = leakReport.totalLeakedBytes;
    
    // 计算内存效率
    auto stats = TXMemoryLeakDetector::instance().getStats();
    if (stats.totalAllocations > 0) {
        report.memoryEfficiency = static_cast<double>(stats.totalDeallocations) / stats.totalAllocations;
    }
    
    // 生成建议
    suggestOptimizations(report);
    
    return report;
}

void TXSmartMemoryManager::optimize() {
    std::lock_guard<std::mutex> lock(mutex_);
    
    // 收缩内存池
    if (generalPool_) {
        generalPool_->shrink();
    }
    
    if (stringPool_) {
        stringPool_->clear();
    }
    
    // 清理旧的分配记录
    TXMemoryLeakDetector::instance().forceCleanup();
}

void TXSmartMemoryManager::emergencyCleanup() {
    std::lock_guard<std::mutex> lock(mutex_);
    
    if (generalPool_) {
        generalPool_->clear();
    }
    
    if (stringPool_) {
        stringPool_->clear();
    }
    
    TXMemoryLeakDetector::instance().forceCleanup();
}

bool TXSmartMemoryManager::shouldUsePool(size_t size) const {
    return size <= 64; // 小于64字节使用内存池
}

void* TXSmartMemoryManager::allocateFromPool(size_t size) {
    return generalPool_->allocate(size);
}

void* TXSmartMemoryManager::allocateFromSystem(size_t size) {
    return std::malloc(size);
}

void TXSmartMemoryManager::deallocateToPool(void* ptr) {
    generalPool_->deallocate(ptr);
}

void TXSmartMemoryManager::deallocateToSystem(void* ptr) {
    std::free(ptr);
}

void TXSmartMemoryManager::updateFragmentationStats() {
    // TODO: 实现碎片统计
}

void TXSmartMemoryManager::suggestOptimizations(HealthReport& report) {
    if (report.hasLeaks) {
        report.recommendations.push_back("检测到内存泄漏，建议检查对象生命周期管理");
    }
    
    if (report.memoryEfficiency < 0.9) {
        report.recommendations.push_back("内存释放效率较低，建议优化内存管理策略");
    }
    
    auto poolRatio = static_cast<double>(poolAllocations_) / (poolAllocations_ + systemAllocations_);
    if (poolRatio < 0.5) {
        report.recommendations.push_back("内存池使用率较低，建议调整池大小配置");
    }
}

} // namespace TinaXlsx
