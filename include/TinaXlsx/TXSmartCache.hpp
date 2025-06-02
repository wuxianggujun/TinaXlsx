//
// @file TXSmartCache.hpp
// @brief 智能缓存系统 - 提升频繁访问数据的性能
//

#pragma once

#include "TXTypes.hpp"
#include <unordered_map>
#include <unordered_set>
#include <vector>
#include <memory>
#include <chrono>
#include <string>

namespace TinaXlsx {

/**
 * @brief LRU缓存模板类
 * 
 * 基于最近最少使用算法的高性能缓存
 */
template<typename Key, typename Value>
class TXLRUCache {
public:
    explicit TXLRUCache(size_t capacity) : capacity_(capacity) {}
    
    /**
     * @brief 获取缓存值
     */
    bool get(const Key& key, Value& value) {
        auto it = cache_.find(key);
        if (it == cache_.end()) {
            return false;
        }
        
        // 移动到最前面（最近使用）
        moveToFront(it->second);
        value = it->second->value;
        return true;
    }
    
    /**
     * @brief 设置缓存值
     */
    void put(const Key& key, const Value& value) {
        auto it = cache_.find(key);
        if (it != cache_.end()) {
            // 更新现有值
            it->second->value = value;
            moveToFront(it->second);
            return;
        }
        
        // 添加新值
        if (cache_.size() >= capacity_) {
            evictLRU();
        }
        
        auto node = std::make_shared<Node>(key, value);
        addToFront(node);
        cache_[key] = node;
    }
    
    /**
     * @brief 清空缓存
     */
    void clear() {
        cache_.clear();
        head_ = tail_ = nullptr;
    }
    
    /**
     * @brief 获取缓存统计
     */
    struct Stats {
        size_t size = 0;
        size_t capacity = 0;
        double hitRate = 0.0;
        size_t hits = 0;
        size_t misses = 0;
    };
    
    Stats getStats() const {
        Stats stats;
        stats.size = cache_.size();
        stats.capacity = capacity_;
        stats.hits = hits_;
        stats.misses = misses_;
        if (hits_ + misses_ > 0) {
            stats.hitRate = static_cast<double>(hits_) / (hits_ + misses_);
        }
        return stats;
    }

private:
    struct Node {
        Key key;
        Value value;
        std::shared_ptr<Node> prev;
        std::shared_ptr<Node> next;
        
        Node(const Key& k, const Value& v) : key(k), value(v) {}
    };
    
    size_t capacity_;
    std::unordered_map<Key, std::shared_ptr<Node>> cache_;
    std::shared_ptr<Node> head_;
    std::shared_ptr<Node> tail_;
    
    mutable size_t hits_ = 0;
    mutable size_t misses_ = 0;
    
    void moveToFront(std::shared_ptr<Node> node) {
        if (node == head_) return;
        
        // 从当前位置移除
        if (node->prev) node->prev->next = node->next;
        if (node->next) node->next->prev = node->prev;
        if (node == tail_) tail_ = node->prev;
        
        // 添加到前面
        addToFront(node);
    }
    
    void addToFront(std::shared_ptr<Node> node) {
        node->prev = nullptr;
        node->next = head_;
        
        if (head_) head_->prev = node;
        head_ = node;
        
        if (!tail_) tail_ = node;
    }
    
    void evictLRU() {
        if (!tail_) return;
        
        cache_.erase(tail_->key);
        
        if (tail_->prev) {
            tail_->prev->next = nullptr;
            tail_ = tail_->prev;
        } else {
            head_ = tail_ = nullptr;
        }
    }
};

/**
 * @brief 字符串缓存池
 * 
 * 专门用于缓存频繁使用的字符串，减少内存分配
 */
class TXStringCache {
public:
    explicit TXStringCache(size_t maxSize = 10000);
    
    /**
     * @brief 获取或创建字符串
     */
    const std::string& intern(const std::string& str);
    
    /**
     * @brief 预加载常用字符串
     */
    void preloadCommonStrings();
    
    /**
     * @brief 获取缓存统计
     */
    struct Stats {
        size_t totalStrings = 0;
        size_t uniqueStrings = 0;
        size_t memoryUsed = 0;
        size_t memorySaved = 0;
        double deduplicationRate = 0.0;
    };
    
    Stats getStats() const;
    
    /**
     * @brief 清理缓存
     */
    void cleanup();

private:
    std::unordered_set<std::string> stringPool_;
    size_t maxSize_;
    mutable size_t totalRequests_ = 0;
    mutable size_t cacheHits_ = 0;
};

/**
 * @brief 样式缓存
 * 
 * 缓存样式对象和样式索引的映射关系
 */
class TXStyleCache {
public:
    explicit TXStyleCache(size_t capacity = 1000);
    
    /**
     * @brief 获取样式索引
     */
    bool getStyleIndex(const std::string& styleKey, uint32_t& index);
    
    /**
     * @brief 设置样式索引
     */
    void setStyleIndex(const std::string& styleKey, uint32_t index);
    
    /**
     * @brief 获取样式对象
     */
    bool getStyleObject(uint32_t index, class TXCellStyle& style);
    
    /**
     * @brief 设置样式对象
     */
    void setStyleObject(uint32_t index, const class TXCellStyle& style);
    
    /**
     * @brief 清空缓存
     */
    void clear();
    
    /**
     * @brief 获取缓存统计
     */
    struct Stats {
        size_t keyToIndexSize = 0;
        size_t indexToStyleSize = 0;
        double keyToIndexHitRate = 0.0;
        double indexToStyleHitRate = 0.0;
    };
    
    Stats getStats() const;

private:
    TXLRUCache<std::string, uint32_t> keyToIndexCache_;
    TXLRUCache<uint32_t, class TXCellStyle> indexToStyleCache_;
};

/**
 * @brief 坐标缓存
 * 
 * 缓存坐标字符串转换结果
 */
class TXCoordinateCache {
public:
    explicit TXCoordinateCache(size_t capacity = 5000);
    
    /**
     * @brief 获取坐标字符串
     */
    const std::string& getCoordinateString(const TXCoordinate& coord);
    
    /**
     * @brief 获取坐标对象
     */
    bool getCoordinate(const std::string& address, TXCoordinate& coord);
    
    /**
     * @brief 清空缓存
     */
    void clear();
    
    /**
     * @brief 获取缓存统计
     */
    struct Stats {
        size_t coordToStringSize = 0;
        size_t stringToCoordSize = 0;
        double coordToStringHitRate = 0.0;
        double stringToCoordHitRate = 0.0;
    };
    
    Stats getStats() const;

private:
    TXLRUCache<TXCoordinate, std::string> coordToStringCache_;
    TXLRUCache<std::string, TXCoordinate> stringToCoordCache_;
    
    // 坐标哈希函数
    struct CoordinateHash {
        std::size_t operator()(const TXCoordinate& coord) const {
            return std::hash<uint64_t>()(
                (static_cast<uint64_t>(coord.getRow().index()) << 32) | 
                coord.getCol().index()
            );
        }
    };
};

/**
 * @brief 智能缓存管理器
 * 
 * 统一管理所有缓存组件
 */
class TXSmartCacheManager {
public:
    TXSmartCacheManager();
    ~TXSmartCacheManager() = default;
    
    /**
     * @brief 获取字符串缓存
     */
    TXStringCache& getStringCache() { return stringCache_; }
    
    /**
     * @brief 获取样式缓存
     */
    TXStyleCache& getStyleCache() { return styleCache_; }
    
    /**
     * @brief 获取坐标缓存
     */
    TXCoordinateCache& getCoordinateCache() { return coordinateCache_; }
    
    /**
     * @brief 获取全局统计
     */
    struct GlobalStats {
        TXStringCache::Stats stringStats;
        TXStyleCache::Stats styleStats;
        TXCoordinateCache::Stats coordinateStats;
        size_t totalMemoryUsed = 0;
        size_t totalMemorySaved = 0;
    };
    
    GlobalStats getGlobalStats() const;
    
    /**
     * @brief 清空所有缓存
     */
    void clearAll();
    
    /**
     * @brief 优化缓存（清理不常用的条目）
     */
    void optimize();
    
    /**
     * @brief 预热缓存（加载常用数据）
     */
    void warmup();

private:
    TXStringCache stringCache_;
    TXStyleCache styleCache_;
    TXCoordinateCache coordinateCache_;
    
    std::chrono::steady_clock::time_point lastOptimization_;
    static constexpr std::chrono::minutes OPTIMIZATION_INTERVAL{5};
};

} // namespace TinaXlsx
