//
// @file TXMemoryPool.cpp
// @brief 高性能内存池管理器实现
//

#include "TinaXlsx/TXMemoryPool.hpp"
#include <algorithm>
#include <cassert>

namespace TinaXlsx {

// ==================== TXMemoryPool::Chunk 实现 ====================

TXMemoryPool::Chunk::Chunk(size_t blockSize, size_t blockCount) {
    size_t totalSize = blockSize * blockCount;
    memory = std::make_unique<uint8_t[]>(totalSize);
    
    // 初始化空闲链表
    uint8_t* ptr = memory.get();
    freeList = reinterpret_cast<Block*>(ptr);
    freeCount = blockCount;
    
    Block* current = freeList;
    for (size_t i = 0; i < blockCount - 1; ++i) {
        Block* next = reinterpret_cast<Block*>(ptr + (i + 1) * blockSize);
        current->next = next;
        current = next;
    }
    current->next = nullptr;
}

// ==================== TXMemoryPool 实现 ====================

TXMemoryPool::TXMemoryPool(const PoolConfig& config)
    : config_(config) {
    // 🚀 优化：延迟分配，只在需要时创建chunk
    // 不再预分配chunk，减少基础内存使用
}

TXMemoryPool::~TXMemoryPool() {
    clear();
}

void* TXMemoryPool::allocate(size_t size) {
    if (size > config_.blockSize) {
        return nullptr; // 超出块大小限制
    }
    
    std::unique_lock<std::mutex> lock(mutex_, std::defer_lock);
    if (config_.threadSafe) {
        lock.lock();
    }
    
    // 尝试从全局空闲链表分配
    if (globalFreeList_) {
        Block* block = globalFreeList_;
        globalFreeList_ = block->next;
        
        size_t allocated = config_.blockSize;
        totalAllocated_ += allocated;
        currentUsage_ += allocated;
        
        size_t current = currentUsage_.load();
        size_t peak = peakUsage_.load();
        while (current > peak && !peakUsage_.compare_exchange_weak(peak, current)) {
            peak = peakUsage_.load();
        }
        
        return block;
    }
    
    // 尝试从现有chunk分配
    for (auto& chunk : chunks_) {
        if (chunk->freeCount > 0) {
            Block* block = allocateFromChunk(chunk.get());
            if (block) {
                size_t allocated = config_.blockSize;
                totalAllocated_ += allocated;
                currentUsage_ += allocated;
                
                size_t current = currentUsage_.load();
                size_t peak = peakUsage_.load();
                while (current > peak && !peakUsage_.compare_exchange_weak(peak, current)) {
                    peak = peakUsage_.load();
                }
                
                return block;
            }
        }
    }
    
    // 创建新chunk
    if (config_.autoGrow && chunks_.size() < config_.maxChunks) {
        auto newChunk = std::unique_ptr<Chunk>(createChunk());
        Block* block = allocateFromChunk(newChunk.get());
        chunks_.push_back(std::move(newChunk));
        
        if (block) {
            size_t allocated = config_.blockSize;
            totalAllocated_ += allocated;
            currentUsage_ += allocated;
            
            size_t current = currentUsage_.load();
            size_t peak = peakUsage_.load();
            while (current > peak && !peakUsage_.compare_exchange_weak(peak, current)) {
                peak = peakUsage_.load();
            }
        }
        
        return block;
    }
    
    return nullptr; // 分配失败
}

void TXMemoryPool::deallocate(void* ptr) {
    if (!ptr || !isFromPool(ptr)) {
        return;
    }
    
    std::unique_lock<std::mutex> lock(mutex_, std::defer_lock);
    if (config_.threadSafe) {
        lock.lock();
    }
    
    // 添加到全局空闲链表
    Block* block = static_cast<Block*>(ptr);
    block->next = globalFreeList_;
    globalFreeList_ = block;
    
    size_t deallocated = config_.blockSize;
    totalDeallocated_ += deallocated;
    currentUsage_ -= deallocated;
}

TXMemoryPool::PoolStats TXMemoryPool::getStats() const {
    std::unique_lock<std::mutex> lock(mutex_, std::defer_lock);
    if (config_.threadSafe) {
        lock.lock();
    }
    
    PoolStats stats;
    stats.totalAllocated = totalAllocated_.load();
    stats.totalDeallocated = totalDeallocated_.load();
    stats.currentUsage = currentUsage_.load();
    stats.peakUsage = peakUsage_.load();
    stats.totalChunks = chunks_.size();
    
    // 计算空闲块数
    Block* current = globalFreeList_;
    while (current) {
        stats.freeBlocks++;
        current = current->next;
    }
    
    for (const auto& chunk : chunks_) {
        stats.freeBlocks += chunk->freeCount;
    }
    
    return stats;
}

void TXMemoryPool::clear() {
    std::unique_lock<std::mutex> lock(mutex_, std::defer_lock);
    if (config_.threadSafe) {
        lock.lock();
    }

    chunks_.clear();
    globalFreeList_ = nullptr;

    totalAllocated_ = 0;
    totalDeallocated_ = 0;
    currentUsage_ = 0;
    peakUsage_ = 0;
}

void TXMemoryPool::shrink() {
    std::unique_lock<std::mutex> lock(mutex_, std::defer_lock);
    if (config_.threadSafe) {
        lock.lock();
    }

    // 移除完全空闲的chunk
    chunks_.erase(
        std::remove_if(chunks_.begin(), chunks_.end(),
            [this](const std::unique_ptr<Chunk>& chunk) {
                return chunk->freeCount == config_.blocksPerChunk;
            }),
        chunks_.end()
    );
}

TXMemoryPool::Chunk* TXMemoryPool::createChunk() {
    return new Chunk(config_.blockSize, config_.blocksPerChunk);
}

TXMemoryPool::Block* TXMemoryPool::allocateFromChunk(Chunk* chunk) {
    if (chunk->freeCount == 0 || !chunk->freeList) {
        return nullptr;
    }
    
    Block* block = chunk->freeList;
    chunk->freeList = block->next;
    chunk->freeCount--;
    
    return block;
}

void TXMemoryPool::deallocateToChunk(void* ptr) {
    // 这个方法在当前实现中不使用
    // 所有释放都通过全局空闲链表处理
}

bool TXMemoryPool::isFromPool(void* ptr) const {
    if (!ptr) return false;
    
    uint8_t* bytePtr = static_cast<uint8_t*>(ptr);
    
    for (const auto& chunk : chunks_) {
        uint8_t* chunkStart = chunk->memory.get();
        uint8_t* chunkEnd = chunkStart + (config_.blockSize * config_.blocksPerChunk);
        
        if (bytePtr >= chunkStart && bytePtr < chunkEnd) {
            // 检查是否对齐到块边界
            size_t offset = bytePtr - chunkStart;
            return (offset % config_.blockSize) == 0;
        }
    }
    
    return false;
}

// 注意：TXStringPool 实现已移动到 TXCompactCell.cpp 中





// ==================== TXMemoryManager 实现 ====================

TXMemoryManager& TXMemoryManager::instance() {
    static TXMemoryManager instance;
    return instance;
}

TXMemoryManager::TXMemoryManager() {
    // 创建默认池
    generalPool_ = std::make_unique<TXMemoryPool>();
}

TXMemoryManager::~TXMemoryManager() {
    clearAll();
}

TXMemoryPool& TXMemoryManager::getGeneralPool() {
    return *generalPool_;
}

TXMemoryPool& TXMemoryManager::getPool(size_t blockSize) {
    std::lock_guard<std::mutex> lock(mutex_);

    auto it = sizedPools_.find(blockSize);
    if (it != sizedPools_.end()) {
        return *it->second;
    }

    // 创建新的指定大小的池
    TXMemoryPool::PoolConfig config;
    config.blockSize = blockSize;
    config.blocksPerChunk = std::max(size_t(64), 4096 / blockSize);

    auto pool = std::make_unique<TXMemoryPool>(config);
    TXMemoryPool* poolPtr = pool.get();
    sizedPools_[blockSize] = std::move(pool);

    return *poolPtr;
}

void TXMemoryManager::clearAll() {
    std::lock_guard<std::mutex> lock(mutex_);

    if (generalPool_) {
        generalPool_->clear();
    }

    for (auto& [size, pool] : sizedPools_) {
        pool->clear();
    }
    sizedPools_.clear();
}

void TXMemoryManager::shrinkAll() {
    std::lock_guard<std::mutex> lock(mutex_);

    if (generalPool_) {
        generalPool_->shrink();
    }

    for (auto& [size, pool] : sizedPools_) {
        pool->shrink();
    }
}

TXMemoryManager::GlobalStats TXMemoryManager::getGlobalStats() const {
    std::lock_guard<std::mutex> lock(mutex_);

    GlobalStats stats;

    if (generalPool_) {
        stats.generalPool = generalPool_->getStats();
        stats.totalMemoryUsage += stats.generalPool.currentUsage;
    }

    stats.totalPools = 1 + sizedPools_.size(); // general + sized pools

    for (const auto& [size, pool] : sizedPools_) {
        auto poolStats = pool->getStats();
        stats.totalMemoryUsage += poolStats.currentUsage;
    }

    return stats;
}

} // namespace TinaXlsx
