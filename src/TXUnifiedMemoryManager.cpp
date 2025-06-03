//
// @file TXUnifiedMemoryManager.cpp
// @brief 统一内存管理器实现
//

#include "TinaXlsx/TXUnifiedMemoryManager.hpp"
#include <chrono>
#include <sstream>
#include <iomanip>

namespace TinaXlsx {

// ==================== TXUnifiedMemoryManager 实现 ====================

TXUnifiedMemoryManager::TXUnifiedMemoryManager(const Config& config) : config_(config) {
    initializeComponents();
}

TXUnifiedMemoryManager::~TXUnifiedMemoryManager() {
    if (smart_manager_) {
        smart_manager_->stopMonitoring();
    }
}

void TXUnifiedMemoryManager::initializeComponents() {
    // 初始化Slab分配器
    if (config_.enable_slab_allocator) {
        slab_allocator_ = std::make_unique<TXSlabAllocator>();
        slab_allocator_->enableAutoReclaim(config_.enable_auto_reclaim);
    }
    
    // 初始化Chunk分配器
    chunk_allocator_ = std::make_unique<TXChunkAllocator>();
    chunk_allocator_->setChunkSize(config_.chunk_size);
    chunk_allocator_->setMemoryLimit(config_.memory_limit);
    
    // 初始化智能管理器
    if (config_.enable_monitoring) {
        MemoryMonitorConfig monitor_config;
        monitor_config.warning_threshold_mb = config_.warning_threshold_mb;
        monitor_config.critical_threshold_mb = config_.critical_threshold_mb;
        monitor_config.emergency_threshold_mb = config_.emergency_threshold_mb;
        monitor_config.enable_auto_cleanup = true;
        monitor_config.enable_memory_warnings = true;
        
        smart_manager_ = std::make_unique<TXSmartMemoryManager>(*chunk_allocator_, monitor_config);
    }
}

void* TXUnifiedMemoryManager::allocate(size_t size) {
    auto start_time = std::chrono::steady_clock::now();
    void* ptr = nullptr;
    
    AllocatorType allocator_type = selectAllocator(size);
    
    if (allocator_type == AllocatorType::SLAB && slab_allocator_) {
        ptr = slab_allocator_->allocate(size);
        if (ptr) {
            small_allocation_count_.fetch_add(1);
        }
    }
    
    if (!ptr) {
        // 回退到Chunk分配器
        ptr = chunk_allocator_->allocate(size);
        if (ptr) {
            large_allocation_count_.fetch_add(1);
            allocator_type = AllocatorType::CHUNK;
        }
    }
    
    auto end_time = std::chrono::steady_clock::now();
    
    if (ptr) {
        updateStats(allocator_type, size, start_time, end_time);
    }
    
    return ptr;
}

bool TXUnifiedMemoryManager::deallocate(void* ptr) {
    if (!ptr) return false;

    // 先尝试Slab分配器
    if (slab_allocator_ && slab_allocator_->deallocate(ptr)) {
        return true;
    }

    // 注意：TXChunkAllocator不支持单独释放，只能整体释放
    // 这里我们只能返回false，表示无法释放单个Chunk分配的内存
    return false;
}

std::vector<void*> TXUnifiedMemoryManager::allocateBatch(const std::vector<size_t>& sizes) {
    std::vector<void*> results;
    results.reserve(sizes.size());
    
    for (size_t size : sizes) {
        results.push_back(allocate(size));
    }
    
    return results;
}

size_t TXUnifiedMemoryManager::compactAll() {
    size_t total_freed = 0;

    if (slab_allocator_) {
        total_freed += slab_allocator_->smartCompact();
    }

    // TXChunkAllocator::compact()返回void，我们无法获取释放的内存量
    chunk_allocator_->compact();

    return total_freed;
}

size_t TXUnifiedMemoryManager::smartCleanup() {
    size_t total_cleaned = 0;
    
    if (slab_allocator_) {
        total_cleaned += slab_allocator_->checkAndReclaim();
    }
    
    if (smart_manager_) {
        total_cleaned += smart_manager_->triggerCleanup(false);
    }
    
    return total_cleaned;
}

void TXUnifiedMemoryManager::clear() {
    if (slab_allocator_) {
        slab_allocator_->clear();
    }
    
    chunk_allocator_->deallocateAll();
    
    // 重置统计
    small_allocation_count_.store(0);
    large_allocation_count_.store(0);
    total_allocation_time_us_.store(0);
    total_allocations_.store(0);
}

void TXUnifiedMemoryManager::startMonitoring() {
    if (smart_manager_) {
        smart_manager_->startMonitoring();
    }
}

void TXUnifiedMemoryManager::stopMonitoring() {
    if (smart_manager_) {
        smart_manager_->stopMonitoring();
    }
}

TXUnifiedMemoryManager::UnifiedStats TXUnifiedMemoryManager::getUnifiedStats() const {
    UnifiedStats stats;
    
    // 获取各组件统计
    if (slab_allocator_) {
        stats.slab_stats = slab_allocator_->getStats();
    }
    
    stats.chunk_stats = chunk_allocator_->getStats();
    
    if (smart_manager_) {
        stats.monitor_stats = smart_manager_->getStats();
    }
    
    // 计算综合指标
    stats.total_memory_usage = getTotalMemoryUsage();
    stats.total_used_memory = getUsedMemorySize();
    stats.overall_efficiency = getOverallEfficiency();
    stats.small_allocations = small_allocation_count_.load();
    stats.large_allocations = large_allocation_count_.load();
    
    // 计算性能指标
    size_t total_allocs = total_allocations_.load();
    if (total_allocs > 0) {
        stats.avg_allocation_time_us = static_cast<double>(total_allocation_time_us_.load()) / total_allocs;
        stats.allocations_per_second = total_allocs * 1000000.0 / total_allocation_time_us_.load();
    }
    
    return stats;
}

std::string TXUnifiedMemoryManager::generateComprehensiveReport() const {
    std::ostringstream report;
    auto stats = getUnifiedStats();
    
    report << "=== TXUnifiedMemoryManager 综合报告 ===\n";
    
    // 总体概况
    report << "\n📊 总体概况:\n";
    report << "  总内存使用: " << (stats.total_memory_usage / 1024.0 / 1024.0) << " MB\n";
    report << "  实际使用: " << (stats.total_used_memory / 1024.0 / 1024.0) << " MB\n";
    report << "  整体效率: " << std::fixed << std::setprecision(2) << (stats.overall_efficiency * 100) << "%\n";
    report << "  小对象分配: " << stats.small_allocations << " 次\n";
    report << "  大对象分配: " << stats.large_allocations << " 次\n";
    
    // 性能指标
    report << "\n⚡ 性能指标:\n";
    report << "  平均分配时间: " << std::fixed << std::setprecision(2) << stats.avg_allocation_time_us << " μs\n";
    report << "  分配速率: " << std::fixed << std::setprecision(0) << stats.allocations_per_second << " 次/秒\n";
    
    // Slab分配器详情
    if (slab_allocator_) {
        report << "\n🔹 Slab分配器 (小对象 ≤8KB):\n";
        report << "  总内存: " << (stats.slab_stats.total_memory / 1024.0) << " KB\n";
        report << "  使用内存: " << (stats.slab_stats.used_memory / 1024.0) << " KB\n";
        report << "  内存效率: " << std::fixed << std::setprecision(2) << (stats.slab_stats.memory_efficiency * 100) << "%\n";
        report << "  碎片率: " << std::fixed << std::setprecision(2) << (stats.slab_stats.fragmentation_ratio * 100) << "%\n";
        report << "  活跃slab: " << stats.slab_stats.active_slabs << "/" << stats.slab_stats.total_slabs << "\n";
    }
    
    // Chunk分配器详情
    report << "\n🔸 Chunk分配器 (大对象 >8KB):\n";
    report << "  总内存: " << (stats.chunk_stats.total_allocated / 1024.0 / 1024.0) << " MB\n";
    report << "  分配次数: " << stats.chunk_stats.allocation_count << "\n";
    report << "  失败次数: " << stats.chunk_stats.failed_allocations << "\n";
    report << "  活跃块数: " << stats.chunk_stats.active_chunks << "\n";
    
    // 智能监控详情
    if (smart_manager_) {
        report << "\n🧠 智能监控:\n";
        report << "  当前内存: " << stats.monitor_stats.current_memory_usage << " MB\n";
        report << "  峰值内存: " << stats.monitor_stats.peak_memory_usage << " MB\n";
        report << "  警告事件: " << stats.monitor_stats.warning_events << " 次\n";
        report << "  严重事件: " << stats.monitor_stats.critical_events << " 次\n";
        report << "  清理事件: " << stats.monitor_stats.cleanup_events << " 次\n";
    }
    
    return report.str();
}

size_t TXUnifiedMemoryManager::getTotalMemoryUsage() const {
    size_t total = 0;
    
    if (slab_allocator_) {
        total += slab_allocator_->getTotalMemoryUsage();
    }
    
    total += chunk_allocator_->getTotalMemoryUsage();
    
    return total;
}

size_t TXUnifiedMemoryManager::getUsedMemorySize() const {
    size_t total = 0;
    
    if (slab_allocator_) {
        total += slab_allocator_->getUsedMemorySize();
    }
    
    // Chunk分配器的使用量等于分配量
    total += chunk_allocator_->getTotalMemoryUsage();
    
    return total;
}

double TXUnifiedMemoryManager::getOverallEfficiency() const {
    size_t total_memory = getTotalMemoryUsage();
    size_t used_memory = getUsedMemorySize();
    
    if (total_memory == 0) return 0.0;
    
    return static_cast<double>(used_memory) / total_memory;
}

void TXUnifiedMemoryManager::updateConfig(const Config& config) {
    config_ = config;
    
    // 重新初始化组件
    if (smart_manager_) {
        smart_manager_->stopMonitoring();
    }
    
    initializeComponents();
}

void TXUnifiedMemoryManager::updateStats(AllocatorType type, size_t size, 
                                        std::chrono::steady_clock::time_point start_time,
                                        std::chrono::steady_clock::time_point end_time) {
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    
    total_allocation_time_us_.fetch_add(duration.count());
    total_allocations_.fetch_add(1);
}

// ==================== GlobalUnifiedMemoryManager 实现 ====================

std::unique_ptr<TXUnifiedMemoryManager> GlobalUnifiedMemoryManager::instance_;
std::mutex GlobalUnifiedMemoryManager::instance_mutex_;

TXUnifiedMemoryManager& GlobalUnifiedMemoryManager::getInstance() {
    std::lock_guard<std::mutex> lock(instance_mutex_);
    if (!instance_) {
        throw std::runtime_error("GlobalUnifiedMemoryManager not initialized");
    }
    return *instance_;
}

void GlobalUnifiedMemoryManager::initialize(const TXUnifiedMemoryManager::Config& config) {
    std::lock_guard<std::mutex> lock(instance_mutex_);
    instance_ = std::make_unique<TXUnifiedMemoryManager>(config);
}

void GlobalUnifiedMemoryManager::shutdown() {
    std::lock_guard<std::mutex> lock(instance_mutex_);
    instance_.reset();
}

} // namespace TinaXlsx
