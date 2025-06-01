#include "TinaXlsx/TXMergedCells.hpp"
#include <algorithm>
#include <sstream>
#include <regex>
#include <unordered_set>

namespace TinaXlsx {

// ==================== MergeRegion实现 ====================

TXMergedCells::MergeRegion::MergeRegion(const TXRange& range) 
    : startRow(range.getStart().getRow()), startCol(range.getStart().getCol()),
      endRow(range.getEnd().getRow()), endCol(range.getEnd().getCol()) {}

bool TXMergedCells::MergeRegion::contains(row_t row, column_t col) const {
    return row >= startRow && row <= endRow && col >= startCol && col <= endCol;
}

std::pair<row_t, column_t> TXMergedCells::MergeRegion::getSize() const {
    return {row_t(endRow.index() - startRow.index() + 1), column_t(endCol.index() - startCol.index() + 1)};
}

std::size_t TXMergedCells::MergeRegion::getCellCount() const {
    auto size = getSize();
    return static_cast<std::size_t>(size.first.index()) * static_cast<std::size_t>(size.second.index());
}

std::string TXMergedCells::MergeRegion::toString() const {
    TXCoordinate start(startRow, startCol);
    TXCoordinate end(endRow, endCol);
    TXRange range(start, end);
    return range.toAddress();
}

TXMergedCells::MergeRegion TXMergedCells::MergeRegion::fromString(const std::string& rangeStr) {
    TXRange range = TXRange::fromAddress(rangeStr);
    return MergeRegion(range);
}

TXRange TXMergedCells::MergeRegion::toRange() const {
    TXCoordinate start(startRow, startCol);
    TXCoordinate end(endRow, endCol);
    return TXRange(start, end);
}

bool TXMergedCells::MergeRegion::isValid() const {
    return is_valid_coordinate(startRow, startCol) && 
           is_valid_coordinate(endRow, endCol) &&
           startRow <= endRow && startCol <= endCol;
}

bool TXMergedCells::MergeRegion::operator==(const MergeRegion& other) const {
    return startRow == other.startRow && startCol == other.startCol &&
           endRow == other.endRow && endCol == other.endCol;
}

bool TXMergedCells::MergeRegion::operator!=(const MergeRegion& other) const {
    return !(*this == other);
}

bool TXMergedCells::MergeRegion::operator<(const MergeRegion& other) const {
    if (startRow != other.startRow) return startRow < other.startRow;
    if (startCol != other.startCol) return startCol < other.startCol;
    if (endRow != other.endRow) return endRow < other.endRow;
    return endCol < other.endCol;
}

// ==================== TXMergedCells实现 ====================

TXMergedCells::TXMergedCells() = default;

// ==================== 私有辅助方法 ====================

uint64_t TXMergedCells::getCellKey(row_t row, column_t col) const {
    return (static_cast<uint64_t>(row.index()) << 32) | static_cast<uint64_t>(col.index());
}

void TXMergedCells::updateCellMapping(const MergeRegion& region, const MergeRegion* regionPtr) {
    for (row_t r = region.startRow; r <= region.endRow; ++r) {
        for (column_t c = region.startCol; c <= region.endCol; ++c) {
            uint64_t key = getCellKey(r, c);
            if (regionPtr) {
                cellToRegionMap_[key] = regionPtr;
            } else {
                cellToRegionMap_.erase(key);
            }
        }
    }
}

bool TXMergedCells::canMergeInternal(const MergeRegion& region) const {
    // 检查是否与现有合并区域重叠
    for (const auto& existingRegion : mergeRegions_) {
        if (isOverlapping(region, existingRegion)) {
            return false;
        }
    }
    return true;
}

// ==================== 合并操作 ====================

bool TXMergedCells::mergeCells(row_t startRow, column_t startCol,
                              row_t endRow, column_t endCol) {
    MergeRegion region(startRow, startCol, endRow, endCol);
    return mergeCells(region.toRange());
}

bool TXMergedCells::mergeCells(const TXRange& range) {
    lastError_.clear();
    
    MergeRegion region(range);
    
    // 验证区域有效性
    if (!isValidRegion(region)) {
        lastError_ = "Invalid merge region";
        return false;
    }
    
    // 规范化区域
    MergeRegion normalizedRegion = normalizeRegion(region);
    
    // 检查是否可以合并
    if (!canMergeInternal(normalizedRegion)) {
        lastError_ = "Region overlaps with existing merged cells";
        return false;
    }
    
    // 单个单元格不需要合并
    if (normalizedRegion.startRow == normalizedRegion.endRow && 
        normalizedRegion.startCol == normalizedRegion.endCol) {
        lastError_ = "Single cell does not need merging";
        return false;
    }
    
    // 添加到合并区域集合
    auto result = mergeRegions_.insert(normalizedRegion);
    if (result.second) {
        // 更新映射
        updateCellMapping(normalizedRegion, &(*result.first));
        return true;
    }
    
    lastError_ = "Merge region already exists";
    return false;
}

bool TXMergedCells::mergeCells(const std::string& rangeStr) {
    MergeRegion region = MergeRegion::fromString(rangeStr);
    return mergeCells(region.toRange());
}

// ==================== 拆分操作 ====================

bool TXMergedCells::unmergeCells(row_t row, column_t col) {
    lastError_.clear();
    
    uint64_t key = getCellKey(row, col);
    auto it = cellToRegionMap_.find(key);
    if (it != cellToRegionMap_.end()) {
        const MergeRegion* regionPtr = it->second;
        MergeRegion region = *regionPtr;
        
        // 清除映射
        updateCellMapping(region, nullptr);
        
        // 从集合中移除
        mergeRegions_.erase(region);
        return true;
    }
    
    lastError_ = "Cell at specified position is not merged";
    return false;
}

bool TXMergedCells::unmergeCells(const MergeRegion& region) {
    lastError_.clear();
    
    auto it = mergeRegions_.find(region);
    if (it != mergeRegions_.end()) {
        // 清除映射
        updateCellMapping(*it, nullptr);
        mergeRegions_.erase(it);
        return true;
    }
    
    lastError_ = "Specified merge region does not exist";
    return false;
}

std::size_t TXMergedCells::unmergeCellsInRange(const TXRange& range) {
    auto overlapping = getOverlappingRegions(range);
    return batchUnmergeCells(overlapping);
}

void TXMergedCells::unmergeAllCells() {
    clear();
}

// ==================== 查询操作 ====================

bool TXMergedCells::isMerged(row_t row, column_t col) const {
    uint64_t key = getCellKey(row, col);
    return cellToRegionMap_.find(key) != cellToRegionMap_.end();
}

const TXMergedCells::MergeRegion* TXMergedCells::getMergeRegion(row_t row, column_t col) const {
    uint64_t key = getCellKey(row, col);
    auto it = cellToRegionMap_.find(key);
    return (it != cellToRegionMap_.end()) ? it->second : nullptr;
}

std::vector<TXMergedCells::MergeRegion> TXMergedCells::getAllMergeRegions() const {
    return std::vector<MergeRegion>(mergeRegions_.begin(), mergeRegions_.end());
}

std::size_t TXMergedCells::getMergeCount() const {
    return mergeRegions_.size();
}

bool TXMergedCells::canMerge(row_t startRow, column_t startCol,
                            row_t endRow, column_t endCol) const {
    MergeRegion region(startRow, startCol, endRow, endCol);
    MergeRegion normalized = normalizeRegion(region);
    return isValidRegion(normalized) && canMergeInternal(normalized);
}

std::vector<TXMergedCells::MergeRegion> TXMergedCells::getOverlappingRegions(const TXRange& range) const {
    std::vector<MergeRegion> result;
    MergeRegion queryRegion(range);
    
    for (const auto& region : mergeRegions_) {
        if (isOverlapping(queryRegion, region)) {
            result.push_back(region);
        }
    }
    
    return result;
}

// ==================== 范围操作 ====================

std::vector<TXMergedCells::MergeRegion> TXMergedCells::getMergeRegionsInRange(const TXRange& range) const {
    std::vector<MergeRegion> result;
    MergeRegion queryRegion(range);
    
    for (const auto& region : mergeRegions_) {
        // 检查合并区域是否完全在查询范围内
        if (queryRegion.contains(region.startRow, region.startCol) && 
            queryRegion.contains(region.endRow, region.endCol)) {
            result.push_back(region);
        }
    }
    
    return result;
}

bool TXMergedCells::hasmergeInRange(const TXRange& range) const {
    return !getMergeRegionsInRange(range).empty();
}

// ==================== 批量操作 ====================

std::size_t TXMergedCells::batchMergeCells(const std::vector<MergeRegion>& regions) {
    std::size_t successCount = 0;
    
    // 预先验证所有区域
    std::vector<MergeRegion> validRegions;
    for (const auto& region : regions) {
        MergeRegion normalized = normalizeRegion(region);
        if (isValidRegion(normalized) && canMergeInternal(normalized)) {
            validRegions.push_back(normalized);
        }
    }
    
    // 批量插入
    for (const auto& region : validRegions) {
        auto result = mergeRegions_.insert(region);
        if (result.second) {
            updateCellMapping(region, &(*result.first));
            successCount++;
        }
    }
    
    return successCount;
}

std::size_t TXMergedCells::batchUnmergeCells(const std::vector<MergeRegion>& regions) {
    std::size_t successCount = 0;
    
    for (const auto& region : regions) {
        auto it = mergeRegions_.find(region);
        if (it != mergeRegions_.end()) {
            updateCellMapping(*it, nullptr);
            mergeRegions_.erase(it);
            successCount++;
        }
    }
    
    return successCount;
}

// ==================== 工具函数 ====================

void TXMergedCells::clear() {
    mergeRegions_.clear();
    cellToRegionMap_.clear();
    lastError_.clear();
}

bool TXMergedCells::empty() const {
    return mergeRegions_.empty();
}

const std::string& TXMergedCells::getLastError() const {
    return lastError_;
}

// ==================== 静态工具函数实现 ====================

bool TXMergedCells::isOverlapping(const MergeRegion& region1, const MergeRegion& region2) {
    return !(region1.endRow < region2.startRow || region1.startRow > region2.endRow ||
             region1.endCol < region2.startCol || region1.startCol > region2.endCol);
}

bool TXMergedCells::isValidRegion(const MergeRegion& region) {
    return region.isValid();
}

TXMergedCells::MergeRegion TXMergedCells::normalizeRegion(const MergeRegion& region) {
    MergeRegion normalized;
    normalized.startRow = std::min(region.startRow, region.endRow);
    normalized.endRow = std::max(region.startRow, region.endRow);
    normalized.startCol = std::min(region.startCol, region.endCol);
    normalized.endCol = std::max(region.startCol, region.endCol);
    return normalized;
}

// ==================== 行列调整操作 ====================

void TXMergedCells::adjustForRowInsertion(row_t insertRow, row_t count) {
    std::vector<MergeRegion> toUpdate;
    std::vector<MergeRegion> toRemove;

    for (const auto& region : mergeRegions_) {
        if (region.startRow >= insertRow) {
            // 整个区域在插入位置之后，需要向下移动
            MergeRegion newRegion = region;
            newRegion.startRow = row_t(region.startRow.index() + count.index());
            newRegion.endRow = row_t(region.endRow.index() + count.index());
            toRemove.push_back(region);
            toUpdate.push_back(newRegion);
        } else if (region.endRow >= insertRow) {
            // 区域跨越插入位置，需要扩展
            MergeRegion newRegion = region;
            newRegion.endRow = row_t(region.endRow.index() + count.index());
            toRemove.push_back(region);
            toUpdate.push_back(newRegion);
        }
    }

    // 批量更新
    batchUnmergeCells(toRemove);
    batchMergeCells(toUpdate);
}

void TXMergedCells::adjustForRowDeletion(row_t deleteRow, row_t count) {
    std::vector<MergeRegion> toUpdate;
    std::vector<MergeRegion> toRemove;

    row_t deleteEnd = row_t(deleteRow.index() + count.index() - 1);

    for (const auto& region : mergeRegions_) {
        if (region.endRow < deleteRow) {
            // 区域在删除范围之前，不受影响
            continue;
        } else if (region.startRow > deleteEnd) {
            // 整个区域在删除范围之后，需要向上移动
            MergeRegion newRegion = region;
            newRegion.startRow = row_t(region.startRow.index() - count.index());
            newRegion.endRow = row_t(region.endRow.index() - count.index());
            toRemove.push_back(region);
            toUpdate.push_back(newRegion);
        } else {
            // 区域与删除范围重叠，需要特殊处理
            if (region.startRow >= deleteRow && region.endRow <= deleteEnd) {
                // 整个区域被删除
                toRemove.push_back(region);
            } else if (region.startRow < deleteRow && region.endRow > deleteEnd) {
                // 区域包含删除范围，需要缩小
                MergeRegion newRegion = region;
                newRegion.endRow = row_t(region.endRow.index() - count.index());
                toRemove.push_back(region);
                toUpdate.push_back(newRegion);
            } else if (region.startRow < deleteRow) {
                // 区域开始在删除范围之前，结束在删除范围内
                MergeRegion newRegion = region;
                newRegion.endRow = row_t(deleteRow.index() - 1);
                toRemove.push_back(region);
                if (newRegion.startRow <= newRegion.endRow) {
                    toUpdate.push_back(newRegion);
                }
            } else {
                // 区域开始在删除范围内，结束在删除范围之后
                MergeRegion newRegion = region;
                newRegion.startRow = deleteRow;
                newRegion.endRow = row_t(region.endRow.index() - count.index());
                toRemove.push_back(region);
                if (newRegion.startRow <= newRegion.endRow) {
                    toUpdate.push_back(newRegion);
                }
            }
        }
    }

    // 批量更新
    batchUnmergeCells(toRemove);
    batchMergeCells(toUpdate);
}

void TXMergedCells::adjustForColumnInsertion(column_t insertCol, column_t count) {
    std::vector<MergeRegion> toUpdate;
    std::vector<MergeRegion> toRemove;

    for (const auto& region : mergeRegions_) {
        if (region.startCol >= insertCol) {
            // 整个区域在插入位置之后，需要向右移动
            MergeRegion newRegion = region;
            newRegion.startCol = column_t(region.startCol.index() + count.index());
            newRegion.endCol = column_t(region.endCol.index() + count.index());
            toRemove.push_back(region);
            toUpdate.push_back(newRegion);
        } else if (region.endCol >= insertCol) {
            // 区域跨越插入位置，需要扩展
            MergeRegion newRegion = region;
            newRegion.endCol = column_t(region.endCol.index() + count.index());
            toRemove.push_back(region);
            toUpdate.push_back(newRegion);
        }
    }

    // 批量更新
    batchUnmergeCells(toRemove);
    batchMergeCells(toUpdate);
}

void TXMergedCells::adjustForColumnDeletion(column_t deleteCol, column_t count) {
    std::vector<MergeRegion> toUpdate;
    std::vector<MergeRegion> toRemove;

    column_t deleteEnd = column_t(deleteCol.index() + count.index() - 1);

    for (const auto& region : mergeRegions_) {
        if (region.endCol < deleteCol) {
            // 区域在删除范围之前，不受影响
            continue;
        } else if (region.startCol > deleteEnd) {
            // 整个区域在删除范围之后，需要向左移动
            MergeRegion newRegion = region;
            newRegion.startCol = column_t(region.startCol.index() - count.index());
            newRegion.endCol = column_t(region.endCol.index() - count.index());
            toRemove.push_back(region);
            toUpdate.push_back(newRegion);
        } else {
            // 区域与删除范围重叠，需要特殊处理
            if (region.startCol >= deleteCol && region.endCol <= deleteEnd) {
                // 整个区域被删除
                toRemove.push_back(region);
            } else if (region.startCol < deleteCol && region.endCol > deleteEnd) {
                // 区域包含删除范围，需要缩小
                MergeRegion newRegion = region;
                newRegion.endCol = column_t(region.endCol.index() - count.index());
                toRemove.push_back(region);
                toUpdate.push_back(newRegion);
            } else if (region.startCol < deleteCol) {
                // 区域开始在删除范围之前，结束在删除范围内
                MergeRegion newRegion = region;
                newRegion.endCol = column_t(deleteCol.index() - 1);
                toRemove.push_back(region);
                if (newRegion.startCol <= newRegion.endCol) {
                    toUpdate.push_back(newRegion);
                }
            } else {
                // 区域开始在删除范围内，结束在删除范围之后
                MergeRegion newRegion = region;
                newRegion.startCol = deleteCol;
                newRegion.endCol = column_t(region.endCol.index() - count.index());
                toRemove.push_back(region);
                if (newRegion.startCol <= newRegion.endCol) {
                    toUpdate.push_back(newRegion);
                }
            }
        }
    }

    // 批量更新
    batchUnmergeCells(toRemove);
    batchMergeCells(toUpdate);
}

} // namespace TinaXlsx