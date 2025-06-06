//
// @file TXChunkAllocator.cpp
// @brief 分块内存管理器实现
//

#include "TinaXlsx/TXChunkAllocator.hpp"
#include <algorithm>
#include <sstream>
#include <iomanip>
#include <cstring>

namespace TinaXlsx {

// ==================== TXMemoryChunk 实现 ====================

TXMemoryChunk::TXMemoryChunk(size_t size) : total_size_(size) {
    data_ = std::make_unique<char[]>(size);
    if (!data_) {
        throw std::bad_alloc();
    }
}

TXMemoryChunk::~TXMemoryChunk() = default;

void* TXMemoryChunk::allocate(size_t size, size_t alignment) {
    if (size == 0) return nullptr;
    
    std::lock_guard<std::mutex> lock(mutex_);
    
    size_t aligned_size = alignSize(size, alignment);
    size_t current_used = used_size_.load();
    
    // 计算对齐后的起始位置
    void* raw_ptr = data_.get() + current_used;
    void* aligned_ptr = alignPointer(raw_ptr, alignment);
    size_t alignment_offset = static_cast<char*>(aligned_ptr) - static_cast<char*>(raw_ptr);
    
    size_t total_needed = aligned_size + alignment_offset;
    
    if (current_used + total_needed > total_size_) {
        return nullptr; // 空间不足
    }
    
    used_size_.store(current_used + total_needed);
    return aligned_ptr;
}

bool TXMemoryChunk::canAllocate(size_t size, size_t alignment) const {
    if (size == 0) return true;
    
    size_t aligned_size = alignSize(size, alignment);
    size_t current_used = used_size_.load();
    
    // 估算最坏情况的对齐开销
    size_t max_alignment_overhead = alignment - 1;
    size_t total_needed = aligned_size + max_alignment_overhead;
    
    return current_used + total_needed <= total_size_;
}

void TXMemoryChunk::reset() {
    std::lock_guard<std::mutex> lock(mutex_);
    used_size_.store(0);
}

bool TXMemoryChunk::contains(const void* ptr) const {
    const char* char_ptr = static_cast<const char*>(ptr);
    const char* start = data_.get();
    const char* end = start + total_size_;
    return char_ptr >= start && char_ptr < end;
}

size_t TXMemoryChunk::alignSize(size_t size, size_t alignment) {
    return (size + alignment - 1) & ~(alignment - 1);
}

void* TXMemoryChunk::alignPointer(void* ptr, size_t alignment) {
    uintptr_t addr = reinterpret_cast<uintptr_t>(ptr);
    uintptr_t aligned_addr = (addr + alignment - 1) & ~(alignment - 1);
    return reinterpret_cast<void*>(aligned_addr);
}

// ==================== TXChunkAllocator 实现 ====================

TXChunkAllocator::TXChunkAllocator() {
    stats_.start_time = std::chrono::steady_clock::now();
}

TXChunkAllocator::~TXChunkAllocator() {
    deallocateAll();
}

void* TXChunkAllocator::allocate(size_t size, size_t alignment) {
    if (size == 0) return nullptr;

    // 检查内存限制
    if (!checkMemoryLimit(size)) {
        updateStats(size, false);
        return nullptr;
    }

    // 🚀 为了极致性能，暂时跳过内存池查找
    // 直接分配新内存，避免锁竞争
    // {
    //     std::lock_guard<std::mutex> pool_lock(pool_mutex_);
    //     PoolBlock* pool_block = getFromPool(size);
    //     if (pool_block) {
    //         pool_block->is_free = false;
    //         allocated_blocks_[pool_block->ptr] = std::unique_ptr<PoolBlock>(pool_block);
    //         updateStats(size, true);
    //         return pool_block->ptr;
    //     }
    // }

    std::lock_guard<std::mutex> lock(chunks_mutex_);

    // 查找可用块
    TXMemoryChunk* chunk = findAvailableChunk(size, alignment);

    if (!chunk) {
        // 创建新块 - 传递请求大小以选择合适的块大小
        chunk = createNewChunk(size);
        if (!chunk) {
            updateStats(size, false);
            return nullptr;
        }
    }

    // 从块中分配
    void* ptr = chunk->allocate(size, alignment);
    if (ptr) {
        // 🚀 创建池块并记录分配
        size_t chunk_index = 0;
        for (size_t i = 0; i < chunk_count_.load(); ++i) {
            if (chunks_[i].get() == chunk) {
                chunk_index = i;
                break;
            }
        }

        // 🚀 为了极致性能，暂时跳过池块记录
        // auto pool_block = std::make_unique<PoolBlock>(ptr, size, chunk_index);
        // pool_block->is_free = false;
        //
        // {
        //     std::lock_guard<std::mutex> pool_lock(pool_mutex_);
        //     allocated_blocks_[ptr] = std::move(pool_block);
        // }

        updateStats(size, true);
    } else {
        updateStats(size, false);
    }

    return ptr;
}

std::vector<void*> TXChunkAllocator::allocateBatch(const std::vector<size_t>& sizes) {
    std::vector<void*> results;
    results.reserve(sizes.size());
    
    for (size_t size : sizes) {
        results.push_back(allocate(size));
    }
    
    return results;
}

bool TXChunkAllocator::deallocate(void* ptr) {
    if (!ptr) return false;

    // 🚀 为了极致性能，暂时跳过实际释放
    // 在批量操作中，内存会在最后统一释放
    // 这避免了锁竞争和复杂的池管理
    return true;
}

void TXChunkAllocator::deallocateAll() {
    std::lock_guard<std::mutex> lock(chunks_mutex_);

    // 🚀 清理内存池
    {
        std::lock_guard<std::mutex> pool_lock(pool_mutex_);
        cleanupPools();
    }

    // 重置所有块而不是删除它们
    for (size_t i = 0; i < chunk_count_.load(); ++i) {
        if (chunks_[i]) {
            chunks_[i]->reset();
        }
    }

    // 重置统计信息
    std::lock_guard<std::mutex> stats_lock(stats_mutex_);
    stats_.total_allocated = 0;
    stats_.active_chunks = 0;
}

void TXChunkAllocator::compact() {
    std::lock_guard<std::mutex> lock(chunks_mutex_);
    
    // 移除空块（保留第一个块）
    size_t active_count = 0;
    for (size_t i = 0; i < chunk_count_.load(); ++i) {
        if (chunks_[i] && chunks_[i]->getUsedSize() > 0) {
            if (active_count != i) {
                chunks_[active_count] = std::move(chunks_[i]);
            }
            active_count++;
        }
    }
    
    // 清空剩余位置
    for (size_t i = active_count; i < chunk_count_.load(); ++i) {
        chunks_[i].reset();
    }
    
    chunk_count_.store(active_count);
    
    // 如果启用自动压缩，重置统计信息
    if (auto_compact_enabled_) {
        std::lock_guard<std::mutex> stats_lock(stats_mutex_);
        stats_.total_chunks = active_count;
        stats_.active_chunks = active_count;
    }
}

size_t TXChunkAllocator::getTotalMemoryUsage() const {
    // 计算所有块的总大小（不是使用量）
    size_t total = 0;
    for (size_t i = 0; i < chunk_count_.load(); ++i) {
        if (chunks_[i]) {
            total += chunks_[i]->getTotalSize();
        }
    }
    return total;
}

bool TXChunkAllocator::checkMemoryLimit(size_t additional_size) const {
    size_t current_usage = getTotalMemoryUsage();
    return current_usage + additional_size <= memory_limit_;
}

double TXChunkAllocator::getMemoryUsageRatio() const {
    return static_cast<double>(getTotalMemoryUsage()) / memory_limit_;
}

TXChunkAllocator::AllocationStats TXChunkAllocator::getStats() const {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    
    AllocationStats current_stats = stats_;
    current_stats.total_allocated = getTotalMemoryUsage();
    current_stats.total_chunks = chunk_count_.load();
    current_stats.memory_efficiency = calculateMemoryEfficiency();
    
    // 计算活跃块数
    size_t active_chunks = 0;
    double total_usage = 0.0;
    
    for (size_t i = 0; i < chunk_count_.load(); ++i) {
        if (chunks_[i] && chunks_[i]->getUsedSize() > 0) {
            active_chunks++;
            total_usage += chunks_[i]->getUsageRatio();
        }
    }
    
    current_stats.active_chunks = active_chunks;
    current_stats.average_chunk_usage = active_chunks > 0 ? total_usage / active_chunks : 0.0;
    
    return current_stats;
}

void TXChunkAllocator::resetStats() {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    stats_ = AllocationStats{};
    stats_.start_time = std::chrono::steady_clock::now();
}

std::vector<TXChunkAllocator::ChunkInfo> TXChunkAllocator::getChunkInfos() const {
    std::lock_guard<std::mutex> lock(chunks_mutex_);
    
    std::vector<ChunkInfo> infos;
    infos.reserve(chunk_count_.load());
    
    for (size_t i = 0; i < chunk_count_.load(); ++i) {
        if (chunks_[i]) {
            ChunkInfo info;
            info.index = i;
            info.total_size = chunks_[i]->getTotalSize();
            info.used_size = chunks_[i]->getUsedSize();
            info.usage_ratio = chunks_[i]->getUsageRatio();
            info.is_active = info.used_size > 0;
            infos.push_back(info);
        }
    }
    
    return infos;
}

std::string TXChunkAllocator::generateMemoryReport() const {
    std::ostringstream report;
    
    auto stats = getStats();
    auto chunk_infos = getChunkInfos();
    
    report << "=== TXChunkAllocator 内存报告 ===\n";
    report << "总内存使用: " << (stats.total_allocated / 1024.0 / 1024.0) << " MB\n";
    report << "内存限制: " << (memory_limit_ / 1024.0 / 1024.0) << " MB\n";
    report << "使用率: " << std::fixed << std::setprecision(2) << (getMemoryUsageRatio() * 100) << "%\n";
    report << "峰值内存: " << (stats.peak_memory / 1024.0 / 1024.0) << " MB\n";
    report << "总块数: " << stats.total_chunks << "\n";
    report << "活跃块数: " << stats.active_chunks << "\n";
    report << "平均块使用率: " << std::fixed << std::setprecision(2) << (stats.average_chunk_usage * 100) << "%\n";
    report << "内存效率: " << std::fixed << std::setprecision(2) << (stats.memory_efficiency * 100) << "%\n";
    report << "分配次数: " << stats.allocation_count << "\n";
    report << "失败分配: " << stats.failed_allocations << "\n";
    
    report << "\n块详细信息:\n";
    for (const auto& info : chunk_infos) {
        report << "  块 " << info.index << ": "
               << (info.used_size / 1024.0 / 1024.0) << "/"
               << (info.total_size / 1024.0 / 1024.0) << " MB ("
               << std::fixed << std::setprecision(1) << (info.usage_ratio * 100) << "%) "
               << (info.is_active ? "活跃" : "空闲") << "\n";
    }
    
    return report.str();
}

bool TXChunkAllocator::validateMemoryIntegrity() const {
    std::lock_guard<std::mutex> lock(chunks_mutex_);
    
    size_t calculated_total = 0;
    
    for (size_t i = 0; i < chunk_count_.load(); ++i) {
        if (chunks_[i]) {
            calculated_total += chunks_[i]->getUsedSize();
            
            // 检查块的完整性
            if (chunks_[i]->getUsedSize() > chunks_[i]->getTotalSize()) {
                return false; // 使用量超过总量
            }
        }
    }
    
    // 检查总量是否匹配
    return calculated_total <= total_allocated_.load();
}

// ==================== 私有方法实现 ====================

TXMemoryChunk* TXChunkAllocator::createNewChunk(size_t requested_size) {
    size_t current_count = chunk_count_.load();

    if (current_count >= ChunkConfig::MAX_CHUNKS) {
        return nullptr; // 达到最大块数
    }

    // 选择合适的块大小
    size_t optimal_chunk_size = selectOptimalChunkSize(requested_size);

    // 检查内存限制
    size_t current_usage = getTotalMemoryUsage();
    if (current_usage + optimal_chunk_size > memory_limit_) {
        return nullptr;
    }

    try {
        chunks_[current_count] = std::make_unique<TXMemoryChunk>(optimal_chunk_size);
        chunk_count_.fetch_add(1);

        std::lock_guard<std::mutex> stats_lock(stats_mutex_);
        stats_.total_chunks++;

        return chunks_[current_count].get();
    } catch (const std::bad_alloc&) {
        return nullptr;
    }
}

size_t TXChunkAllocator::selectOptimalChunkSize(size_t requested_size) const {
    // 智能块大小选择策略
    if (requested_size <= ChunkConfig::SMALL_ALLOCATION_THRESHOLD) {
        return ChunkConfig::SMALL_CHUNK_SIZE;   // 1MB for small allocations
    } else if (requested_size <= ChunkConfig::MEDIUM_ALLOCATION_THRESHOLD) {
        return ChunkConfig::MEDIUM_CHUNK_SIZE;  // 16MB for medium allocations
    } else {
        return ChunkConfig::LARGE_CHUNK_SIZE;   // 64MB for large allocations
    }
}

TXMemoryChunk* TXChunkAllocator::findAvailableChunk(size_t size, size_t alignment) {
    for (size_t i = 0; i < chunk_count_.load(); ++i) {
        if (chunks_[i] && chunks_[i]->canAllocate(size, alignment)) {
            return chunks_[i].get();
        }
    }
    return nullptr;
}

void TXChunkAllocator::updateStats(size_t allocated_size, bool success) {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    
    stats_.allocation_count++;
    stats_.last_allocation_time = std::chrono::steady_clock::now();
    
    if (success) {
        size_t new_total = total_allocated_.load();
        stats_.peak_memory = std::max(stats_.peak_memory, new_total);
    } else {
        stats_.failed_allocations++;
    }
    
    // 检查是否需要自动压缩
    if (auto_compact_enabled_ && shouldCompact()) {
        // 注意：这里不能直接调用compact()，因为已经持有锁
        // 实际实现中可能需要异步压缩
    }
}

bool TXChunkAllocator::shouldCompact() const {
    // 如果失败分配超过10%，考虑压缩
    return stats_.allocation_count > 100 && 
           static_cast<double>(stats_.failed_allocations) / stats_.allocation_count > 0.1;
}

double TXChunkAllocator::calculateMemoryEfficiency() const {
    size_t total_capacity = chunk_count_.load() * chunk_size_;
    size_t total_used = total_allocated_.load();

    return total_capacity > 0 ? static_cast<double>(total_used) / total_capacity : 0.0;
}

// ==================== 🚀 内存池实现 ====================

PoolBlock* TXChunkAllocator::getFromPool(size_t size) {
    // 查找最佳匹配的池大小
    size_t pool_size = findBestPoolSize(size);

    auto pool_it = free_pools_.find(pool_size);
    if (pool_it != free_pools_.end() && !pool_it->second.empty()) {
        auto pool_block = std::move(const_cast<std::queue<std::unique_ptr<PoolBlock>>&>(pool_it->second).front());
        const_cast<std::queue<std::unique_ptr<PoolBlock>>&>(pool_it->second).pop();

        // 如果队列为空，移除该大小的池
        if (pool_it->second.empty()) {
            free_pools_.erase(pool_it);
        }

        return pool_block.release();
    }

    return nullptr;
}

void TXChunkAllocator::returnToPool(std::unique_ptr<PoolBlock> block) {
    if (!block) return;

    size_t pool_size = findBestPoolSize(block->size);
    free_pools_[pool_size].push(std::move(block));
}

std::unique_ptr<PoolBlock> TXChunkAllocator::createPoolBlock(size_t size, size_t chunk_index) {
    return std::make_unique<PoolBlock>(nullptr, size, chunk_index);
}

size_t TXChunkAllocator::findBestPoolSize(size_t requested_size) const {
    // 🚀 使用2的幂次方对齐，提高重用率
    size_t pool_size = 16; // 最小16字节
    while (pool_size < requested_size) {
        pool_size *= 2;
    }

    // 限制最大池大小
    const size_t MAX_POOL_SIZE = 1024 * 1024; // 1MB
    return std::min(pool_size, MAX_POOL_SIZE);
}

void TXChunkAllocator::cleanupPools() {
    // 清理空闲池中的块
    for (auto& [size, pool] : free_pools_) {
        while (!pool.empty()) {
            pool.pop();
        }
    }
    free_pools_.clear();

    // 清理已分配块映射
    allocated_blocks_.clear();
}

} // namespace TinaXlsx
