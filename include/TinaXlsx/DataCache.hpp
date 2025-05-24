/**
 * @file DataCache.hpp
 * @brief Data cache manager - Specialized for Excel data caching and management
 */

#pragma once

#include "Types.hpp"
#include <unordered_map>
#include <optional>
#include <string>

namespace TinaXlsx {

/**
 * @brief Data cache manager
 * Specialized for Excel data caching and management, implementing LRU strategy
 */
class DataCache {
public:
    using CacheKey = std::string;
    
private:
    std::unordered_map<CacheKey, TableData> tableCache_;
    std::unordered_map<CacheKey, std::pair<RowIndex, ColumnIndex>> dimensionCache_;
    size_t maxCacheSize_ = 10; // Maximum number of cached worksheets

public:
    /**
     * @brief Cache worksheet data
     * @param key Cache key (usually worksheet name + path)
     * @param data Table data
     */
    void cacheTableData(const CacheKey& key, const TableData& data) {
        if (tableCache_.size() >= maxCacheSize_) {
            // Simple LRU strategy: delete the first one
            auto it = tableCache_.begin();
            tableCache_.erase(it);
        }
        tableCache_[key] = data;
    }
    
    /**
     * @brief Get cached table data
     * @param key Cache key
     * @return std::optional<TableData> Cached data, empty if not found
     */
    std::optional<TableData> getCachedTableData(const CacheKey& key) const {
        auto it = tableCache_.find(key);
        return it != tableCache_.end() ? std::make_optional(it->second) : std::nullopt;
    }
    
    /**
     * @brief Cache dimension information
     * @param key Cache key
     * @param rows Number of rows
     * @param cols Number of columns
     */
    void cacheDimensions(const CacheKey& key, RowIndex rows, ColumnIndex cols) {
        dimensionCache_[key] = {rows, cols};
    }
    
    /**
     * @brief Get cached dimension information
     * @param key Cache key
     * @return std::optional<std::pair<RowIndex, ColumnIndex>> Dimension info
     */
    std::optional<std::pair<RowIndex, ColumnIndex>> getCachedDimensions(const CacheKey& key) const {
        auto it = dimensionCache_.find(key);
        return it != dimensionCache_.end() ? std::make_optional(it->second) : std::nullopt;
    }
    
    /**
     * @brief Check if cached data exists
     * @param key Cache key
     * @return bool Whether exists
     */
    bool hasTableData(const CacheKey& key) const {
        return tableCache_.find(key) != tableCache_.end();
    }
    
    /**
     * @brief Check if cached dimension info exists
     * @param key Cache key
     * @return bool Whether exists
     */
    bool hasDimensions(const CacheKey& key) const {
        return dimensionCache_.find(key) != dimensionCache_.end();
    }
    
    /**
     * @brief Clear specific cache
     * @param key Cache key
     */
    void clearCache(const CacheKey& key) {
        tableCache_.erase(key);
        dimensionCache_.erase(key);
    }
    
    /**
     * @brief Clear all cache
     */
    void clear() {
        tableCache_.clear();
        dimensionCache_.clear();
    }
    
    /**
     * @brief Set maximum cache size
     * @param size Maximum size
     */
    void setMaxCacheSize(size_t size) { 
        maxCacheSize_ = size; 
        // Clean up excess if current cache exceeds new limit
        while (tableCache_.size() > maxCacheSize_) {
            auto it = tableCache_.begin();
            tableCache_.erase(it);
        }
    }
    
    /**
     * @brief Get current cache size
     * @return size_t Number of cached worksheets
     */
    size_t getCacheSize() const { 
        return tableCache_.size(); 
    }
    
    /**
     * @brief Get maximum cache size
     * @return size_t Maximum cache size
     */
    size_t getMaxCacheSize() const { 
        return maxCacheSize_; 
    }
};

} // namespace TinaXlsx 