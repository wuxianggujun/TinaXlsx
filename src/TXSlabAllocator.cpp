//
// @file TXSlabAllocator.cpp
// @brief Slab分配器实现
//

#include "TinaXlsx/TXSlabAllocator.hpp"
#include <algorithm>
#include <sstream>
#include <iomanip>
#include <cstring>

namespace TinaXlsx {

// ==================== TXSlab 实现 ====================

TXSlab::TXSlab(size_t object_size) : object_size_(object_size) {
    // 第三阶段：使用优化配置表选择slab大小
    size_t slab_size = SlabConfig::getSlabSize(object_size);

    // 计算可容纳的对象数量
    max_objects_ = slab_size / object_size_;
    free_count_.store(max_objects_);

    // 分配slab内存
    data_ = std::make_unique<char[]>(slab_size);

    // 初始化空闲链表
    initializeFreeList();
}

TXSlab::~TXSlab() = default;

void TXSlab::initializeFreeList() {
    char* ptr = data_.get();
    FreeNode* prev = nullptr;
    
    // 构建空闲对象链表
    for (size_t i = 0; i < max_objects_; ++i) {
        FreeNode* node = reinterpret_cast<FreeNode*>(ptr + i * object_size_);
        if (prev) {
            prev->next = node;
        } else {
            free_list_.store(node);
        }
        node->next = nullptr;
        prev = node;
    }
}

void* TXSlab::allocate() {
    std::lock_guard<std::mutex> lock(mutex_);
    
    FreeNode* node = free_list_.load();
    if (!node) {
        return nullptr; // slab已满
    }
    
    // 从空闲链表移除
    free_list_.store(node->next);
    free_count_.fetch_sub(1);
    
    return static_cast<void*>(node);
}

bool TXSlab::deallocate(void* ptr) {
    if (!contains(ptr)) {
        return false;
    }
    
    std::lock_guard<std::mutex> lock(mutex_);
    
    // 添加到空闲链表头部
    FreeNode* node = static_cast<FreeNode*>(ptr);
    node->next = free_list_.load();
    free_list_.store(node);
    free_count_.fetch_add(1);
    
    return true;
}

bool TXSlab::contains(const void* ptr) const {
    const char* char_ptr = static_cast<const char*>(ptr);
    const char* start = data_.get();
    size_t slab_size = SlabConfig::getSlabSize(object_size_);
    const char* end = start + slab_size;

    if (char_ptr < start || char_ptr >= end) {
        return false;
    }

    // 检查对齐
    size_t offset = char_ptr - start;
    return (offset % object_size_) == 0 && (offset / object_size_) < max_objects_;
}

// ==================== TXSlabAllocator 实现 ====================

TXSlabAllocator::TXSlabAllocator() {
    last_reclaim_time_ = std::chrono::steady_clock::now();
}

TXSlabAllocator::~TXSlabAllocator() {
    clear();
}

void* TXSlabAllocator::allocate(size_t size) {
    if (!canHandle(size)) {
        failed_allocations_.fetch_add(1);
        return nullptr;
    }
    
    size_t size_index = getSizeIndex(size);
    TXSlab* slab = findAvailableSlab(size_index);
    
    if (!slab) {
        slab = createNewSlab(size_index);
        if (!slab) {
            failed_allocations_.fetch_add(1);
            return nullptr;
        }
    }
    
    void* ptr = slab->allocate();
    if (ptr) {
        total_allocations_.fetch_add(1);

        // 第二阶段：每100次分配检查一次自动回收
        if (auto_reclaim_enabled_ && (total_allocations_.load() % 100 == 0)) {
            checkAndReclaim();
        }
    } else {
        failed_allocations_.fetch_add(1);
    }

    return ptr;
}

bool TXSlabAllocator::deallocate(void* ptr) {
    if (!ptr) return false;
    
    size_t size_index;
    TXSlab* slab = findSlabContaining(ptr, size_index);
    
    if (slab && slab->deallocate(ptr)) {
        total_deallocations_.fetch_add(1);
        return true;
    }
    
    return false;
}

std::vector<void*> TXSlabAllocator::allocateBatch(const std::vector<size_t>& sizes) {
    std::vector<void*> results;
    results.reserve(sizes.size());
    
    for (size_t size : sizes) {
        results.push_back(allocate(size));
    }
    
    return results;
}

size_t TXSlabAllocator::compact() {
    size_t total_freed_memory = 0;

    for (size_t i = 0; i < SlabConfig::OBJECT_SIZES.size(); ++i) {
        size_t removed_count = removeEmptySlabs(i);
        size_t object_size = SlabConfig::OBJECT_SIZES[i];
        size_t slab_size = SlabConfig::getSlabSize(object_size);
        total_freed_memory += removed_count * slab_size;
    }

    return total_freed_memory;
}

// ==================== 第二阶段：智能回收实现 ====================

size_t TXSlabAllocator::smartCompact() {
    size_t total_freed_memory = 0;

    for (size_t i = 0; i < SlabConfig::OBJECT_SIZES.size(); ++i) {
        // 智能回收：保留缓存slab
        size_t removed_count = smartRemoveEmptySlabs(i);
        size_t object_size = SlabConfig::OBJECT_SIZES[i];
        size_t slab_size = SlabConfig::getSlabSize(object_size);
        total_freed_memory += removed_count * slab_size;
    }

    return total_freed_memory;
}

size_t TXSlabAllocator::checkAndReclaim() {
    auto now = std::chrono::steady_clock::now();
    auto time_since_last = std::chrono::duration_cast<std::chrono::seconds>(now - last_reclaim_time_);

    // 至少间隔5秒才执行回收
    if (time_since_last.count() < 5) {
        return 0;
    }

    size_t total_reclaimed = 0;

    // 检查每种大小是否需要回收
    for (size_t i = 0; i < SlabConfig::OBJECT_SIZES.size(); ++i) {
        if (shouldTriggerReclaim(i)) {
            size_t reclaimed = smartRemoveEmptySlabs(i);
            size_t object_size = SlabConfig::OBJECT_SIZES[i];
            size_t slab_size = SlabConfig::getSlabSize(object_size);
            total_reclaimed += reclaimed * slab_size;
        }
    }

    if (total_reclaimed > 0) {
        last_reclaim_time_ = now;
        reclaim_counter_.fetch_add(1);
    }

    return total_reclaimed;
}

void TXSlabAllocator::clear() {
    std::lock_guard<std::mutex> lock(global_mutex_);
    
    for (auto& slab_list : slabs_) {
        slab_list.clear();
    }
    
    total_allocations_.store(0);
    total_deallocations_.store(0);
    failed_allocations_.store(0);
}

size_t TXSlabAllocator::getTotalMemoryUsage() const {
    size_t total = 0;

    for (size_t i = 0; i < SlabConfig::OBJECT_SIZES.size(); ++i) {
        // 第三阶段：使用优化配置表
        size_t slab_size = SlabConfig::getSlabSizeByIndex(i);
        total += slabs_[i].size() * slab_size;
    }

    return total;
}

size_t TXSlabAllocator::getUsedMemorySize() const {
    size_t used = 0;
    
    for (size_t i = 0; i < SlabConfig::OBJECT_SIZES.size(); ++i) {
        std::lock_guard<std::mutex> lock(size_mutexes_[i]);
        
        for (const auto& slab : slabs_[i]) {
            size_t allocated_objects = slab->getMaxObjects() - slab->getFreeCount();
            used += allocated_objects * slab->getObjectSize();
        }
    }
    
    return used;
}

TXSlabAllocator::SlabStats TXSlabAllocator::getStats() const {
    SlabStats stats;
    
    for (size_t i = 0; i < SlabConfig::OBJECT_SIZES.size(); ++i) {
        std::lock_guard<std::mutex> lock(size_mutexes_[i]);
        
        stats.slabs_per_size[i] = slabs_[i].size();
        stats.total_slabs += slabs_[i].size();
        
        for (const auto& slab : slabs_[i]) {
            if (!slab->isEmpty()) {
                stats.active_slabs++;
            }
            
            size_t allocated = slab->getMaxObjects() - slab->getFreeCount();
            stats.objects_per_size[i] += allocated;
            stats.allocated_objects += allocated;
            stats.total_objects += slab->getMaxObjects();
        }
        
        // 计算每种大小的效率
        if (stats.slabs_per_size[i] > 0) {
            size_t object_size = SlabConfig::OBJECT_SIZES[i];
            size_t slab_size = SlabConfig::getSlabSize(object_size);
            size_t total_capacity = stats.slabs_per_size[i] * slab_size;
            size_t used_space = stats.objects_per_size[i] * object_size;
            stats.efficiency_per_size[i] = static_cast<double>(used_space) / total_capacity;
        }
    }
    
    stats.total_memory = getTotalMemoryUsage();
    stats.used_memory = getUsedMemorySize();
    
    if (stats.total_memory > 0) {
        stats.memory_efficiency = static_cast<double>(stats.used_memory) / stats.total_memory;
        stats.fragmentation_ratio = 1.0 - stats.memory_efficiency;
    }
    
    return stats;
}

std::string TXSlabAllocator::generateReport() const {
    std::ostringstream report;
    auto stats = getStats();
    
    report << "=== TXSlabAllocator 详细报告 ===\n";
    report << "总内存使用: " << (stats.total_memory / 1024.0) << " KB\n";
    report << "实际使用: " << (stats.used_memory / 1024.0) << " KB\n";
    report << "内存效率: " << std::fixed << std::setprecision(2) << (stats.memory_efficiency * 100) << "%\n";
    report << "碎片率: " << std::fixed << std::setprecision(2) << (stats.fragmentation_ratio * 100) << "%\n";
    report << "总slab数: " << stats.total_slabs << "\n";
    report << "活跃slab数: " << stats.active_slabs << "\n";
    report << "分配对象数: " << stats.allocated_objects << "/" << stats.total_objects << "\n";
    
    report << "\n各大小分级详情:\n";
    for (size_t i = 0; i < SlabConfig::OBJECT_SIZES.size(); ++i) {
        if (stats.slabs_per_size[i] > 0) {
            report << "  " << SlabConfig::OBJECT_SIZES[i] << "字节: "
                   << stats.slabs_per_size[i] << "个slab, "
                   << stats.objects_per_size[i] << "个对象, "
                   << "效率" << std::fixed << std::setprecision(1) 
                   << (stats.efficiency_per_size[i] * 100) << "%\n";
        }
    }
    
    report << "\n分配统计:\n";
    report << "  总分配: " << total_allocations_.load() << "\n";
    report << "  总释放: " << total_deallocations_.load() << "\n";
    report << "  失败分配: " << failed_allocations_.load() << "\n";
    
    return report.str();
}

double TXSlabAllocator::getFragmentationRatio() const {
    return getStats().fragmentation_ratio;
}

// ==================== 私有方法实现 ====================

size_t TXSlabAllocator::getSizeIndex(size_t size) {
    for (size_t i = 0; i < SlabConfig::OBJECT_SIZES.size(); ++i) {
        if (size <= SlabConfig::OBJECT_SIZES[i]) {
            return i;
        }
    }
    return SlabConfig::OBJECT_SIZES.size() - 1; // 最大大小
}

size_t TXSlabAllocator::getActualSize(size_t size) {
    size_t index = getSizeIndex(size);
    return SlabConfig::OBJECT_SIZES[index];
}

TXSlab* TXSlabAllocator::findAvailableSlab(size_t size_index) {
    std::lock_guard<std::mutex> lock(size_mutexes_[size_index]);
    
    for (auto& slab : slabs_[size_index]) {
        if (slab->canAllocate()) {
            return slab.get();
        }
    }
    
    return nullptr;
}

TXSlab* TXSlabAllocator::createNewSlab(size_t size_index) {
    std::lock_guard<std::mutex> lock(size_mutexes_[size_index]);
    
    if (slabs_[size_index].size() >= SlabConfig::MAX_SLABS_PER_SIZE) {
        return nullptr; // 达到最大slab数限制
    }
    
    try {
        auto slab = std::make_unique<TXSlab>(SlabConfig::OBJECT_SIZES[size_index]);
        TXSlab* ptr = slab.get();
        slabs_[size_index].push_back(std::move(slab));
        return ptr;
    } catch (const std::bad_alloc&) {
        return nullptr;
    }
}

TXSlab* TXSlabAllocator::findSlabContaining(void* ptr, size_t& size_index) const {
    for (size_t i = 0; i < SlabConfig::OBJECT_SIZES.size(); ++i) {
        std::lock_guard<std::mutex> lock(size_mutexes_[i]);
        
        for (const auto& slab : slabs_[i]) {
            if (slab->contains(ptr)) {
                size_index = i;
                return slab.get();
            }
        }
    }
    
    return nullptr;
}

size_t TXSlabAllocator::removeEmptySlabs(size_t size_index) {
    std::lock_guard<std::mutex> lock(size_mutexes_[size_index]);
    
    size_t removed_count = 0;
    auto& slab_list = slabs_[size_index];
    
    auto new_end = std::remove_if(slab_list.begin(), slab_list.end(),
        [&removed_count](const std::unique_ptr<TXSlab>& slab) {
            if (slab->isEmpty()) {
                removed_count++;
                return true;
            }
            return false;
        });
    
    slab_list.erase(new_end, slab_list.end());
    return removed_count;
}

// ==================== 智能回收辅助方法 ====================

size_t TXSlabAllocator::smartRemoveEmptySlabs(size_t size_index) {
    std::lock_guard<std::mutex> lock(size_mutexes_[size_index]);

    size_t removed_count = 0;
    auto& slab_list = slabs_[size_index];

    // 计算要保留的缓存slab数量
    size_t cache_count = std::min(SlabConfig::CACHE_SLABS_COUNT, slab_list.size());
    size_t empty_count = 0;

    // 统计空slab数量
    for (const auto& slab : slab_list) {
        if (slab->isEmpty()) {
            empty_count++;
        }
    }

    // 如果空slab数量超过缓存数量，移除多余的
    if (empty_count > cache_count) {
        size_t to_remove = empty_count - cache_count;

        auto new_end = std::remove_if(slab_list.begin(), slab_list.end(),
            [&removed_count, &to_remove](const std::unique_ptr<TXSlab>& slab) {
                if (to_remove > 0 && slab->isEmpty()) {
                    removed_count++;
                    to_remove--;
                    return true;
                }
                return false;
            });

        slab_list.erase(new_end, slab_list.end());
    }

    return removed_count;
}

bool TXSlabAllocator::shouldTriggerReclaim(size_t size_index) const {
    std::lock_guard<std::mutex> lock(size_mutexes_[size_index]);

    const auto& slab_list = slabs_[size_index];
    if (slab_list.empty()) return false;

    size_t total_capacity = 0;
    size_t used_capacity = 0;

    for (const auto& slab : slab_list) {
        total_capacity += slab->getMaxObjects();
        used_capacity += (slab->getMaxObjects() - slab->getFreeCount());
    }

    if (total_capacity == 0) return false;

    double usage_ratio = static_cast<double>(used_capacity) / total_capacity;
    double fragmentation_ratio = 1.0 - usage_ratio;

    // 碎片率超过30%时触发回收
    return fragmentation_ratio > SlabConfig::FRAGMENTATION_THRESHOLD;
}

void TXSlabAllocator::immediateReleaseEmptySlabs() {
    for (size_t i = 0; i < SlabConfig::OBJECT_SIZES.size(); ++i) {
        removeEmptySlabs(i);
    }
}

} // namespace TinaXlsx
