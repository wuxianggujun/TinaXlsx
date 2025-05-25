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

bool TXMergedCells::MergeRegion::contains(TXTypes::RowIndex row, TXTypes::ColIndex col) const {
    return row >= startRow && row <= endRow && col >= startCol && col <= endCol;
}

std::pair<TXTypes::RowIndex, TXTypes::ColIndex> TXMergedCells::MergeRegion::getSize() const {
    return {endRow - startRow + 1, endCol - startCol + 1};
}

std::size_t TXMergedCells::MergeRegion::getCellCount() const {
    auto size = getSize();
    return static_cast<std::size_t>(size.first) * size.second;
}

std::string TXMergedCells::MergeRegion::toString() const {
    TXRange range(startRow, startCol, endRow, endCol);
    return range.toAddress();
}

TXMergedCells::MergeRegion TXMergedCells::MergeRegion::fromString(const std::string& rangeStr) {
    TXRange range = TXRange::fromAddress(rangeStr);
    return MergeRegion(range);
}

TXRange TXMergedCells::MergeRegion::toRange() const {
    return TXRange(startRow, startCol, endRow, endCol);
}

bool TXMergedCells::MergeRegion::isValid() const {
    return TXTypes::isValidCoordinate(startRow, startCol) && 
           TXTypes::isValidCoordinate(endRow, endCol) &&
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

// ==================== TXMergedCells::Impl实现 ====================

class TXMergedCells::Impl {
public:
    // 使用集合存储合并区域，保证有序和唯一性
    std::set<MergeRegion> mergeRegions_;
    
    // 用于快速查找包含特定单元格的合并区域
    std::unordered_map<uint64_t, const MergeRegion*> cellToRegionMap_;
    
    std::string lastError_;
    
    Impl() = default;
    
    // 生成单元格的唯一键
    uint64_t getCellKey(TXTypes::RowIndex row, TXTypes::ColIndex col) const {
        return (static_cast<uint64_t>(row) << 32) | col;
    }
    
    // 更新单元格到区域的映射
    void updateCellMapping(const MergeRegion& region, const MergeRegion* regionPtr) {
        for (TXTypes::RowIndex r = region.startRow; r <= region.endRow; ++r) {
            for (TXTypes::ColIndex c = region.startCol; c <= region.endCol; ++c) {
                uint64_t key = getCellKey(r, c);
                if (regionPtr) {
                    cellToRegionMap_[key] = regionPtr;
                } else {
                    cellToRegionMap_.erase(key);
                }
            }
        }
    }
    
    bool mergeCells(const MergeRegion& region) {
        lastError_.clear();
        
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
    
    bool unmergeCells(const MergeRegion& region) {
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
    
    bool unmergeCellsByPosition(TXTypes::RowIndex row, TXTypes::ColIndex col) {
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
    
    bool canMergeInternal(const MergeRegion& region) const {
        // 检查是否与现有合并区域重叠
        for (const auto& existingRegion : mergeRegions_) {
            if (isOverlapping(region, existingRegion)) {
                return false;
            }
        }
        return true;
    }
    
    std::vector<MergeRegion> getOverlappingRegions(const TXRange& range) const {
        std::vector<MergeRegion> result;
        MergeRegion queryRegion(range);
        
        for (const auto& region : mergeRegions_) {
            if (isOverlapping(queryRegion, region)) {
                result.push_back(region);
            }
        }
        
        return result;
    }
    
    std::size_t batchMergeCells(const std::vector<MergeRegion>& regions) {
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
    
    std::size_t batchUnmergeCells(const std::vector<MergeRegion>& regions) {
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
    
    std::string toXml() const {
        std::ostringstream oss;
        oss << "<mergeCells count=\"" << mergeRegions_.size() << "\">\n";
        
        for (const auto& region : mergeRegions_) {
            oss << "  <mergeCell ref=\"" << region.toString() << "\"/>\n";
        }
        
        oss << "</mergeCells>\n";
        return oss.str();
    }
    
    bool fromXml(const std::string& xmlStr) {
        lastError_.clear();
        clear();
        
        try {
            // 简化的XML解析（实际项目中应使用专业的XML解析器）
            std::regex mergePattern("<mergeCell\\s+ref=\"([^\"]+)\"\\s*/>");
            std::sregex_iterator iter(xmlStr.begin(), xmlStr.end(), mergePattern);
            std::sregex_iterator end;
            
            while (iter != end) {
                std::string ref = (*iter)[1].str();
                MergeRegion region = MergeRegion::fromString(ref);
                if (region.isValid()) {
                    mergeCells(region);
                }
                ++iter;
            }
            
            return true;
        } catch (...) {
            lastError_ = "XML parsing failed";
            return false;
        }
    }
    
    void clear() {
        mergeRegions_.clear();
        cellToRegionMap_.clear();
        lastError_.clear();
    }
};

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

// ==================== TXMergedCells公共接口实现 ====================

TXMergedCells::TXMergedCells() : pImpl(std::make_unique<Impl>()) {}

TXMergedCells::~TXMergedCells() = default;

TXMergedCells::TXMergedCells(TXMergedCells&& other) noexcept : pImpl(std::move(other.pImpl)) {}

TXMergedCells& TXMergedCells::operator=(TXMergedCells&& other) noexcept {
    if (this != &other) {
        pImpl = std::move(other.pImpl);
    }
    return *this;
}

// ==================== 合并操作 ====================

bool TXMergedCells::mergeCells(TXTypes::RowIndex startRow, TXTypes::ColIndex startCol,
                              TXTypes::RowIndex endRow, TXTypes::ColIndex endCol) {
    MergeRegion region(startRow, startCol, endRow, endCol);
    return pImpl->mergeCells(region);
}

bool TXMergedCells::mergeCells(const TXRange& range) {
    MergeRegion region(range);
    return pImpl->mergeCells(region);
}

bool TXMergedCells::mergeCells(const std::string& rangeStr) {
    MergeRegion region = MergeRegion::fromString(rangeStr);
    return pImpl->mergeCells(region);
}

// ==================== 拆分操作 ====================

bool TXMergedCells::unmergeCells(TXTypes::RowIndex row, TXTypes::ColIndex col) {
    return pImpl->unmergeCellsByPosition(row, col);
}

bool TXMergedCells::unmergeCells(const MergeRegion& region) {
    return pImpl->unmergeCells(region);
}

std::size_t TXMergedCells::unmergeCellsInRange(const TXRange& range) {
    auto overlapping = pImpl->getOverlappingRegions(range);
    return pImpl->batchUnmergeCells(overlapping);
}

void TXMergedCells::unmergeAllCells() {
    pImpl->clear();
}

// ==================== 查询操作 ====================

bool TXMergedCells::isMerged(TXTypes::RowIndex row, TXTypes::ColIndex col) const {
    uint64_t key = pImpl->getCellKey(row, col);
    return pImpl->cellToRegionMap_.find(key) != pImpl->cellToRegionMap_.end();
}

const TXMergedCells::MergeRegion* TXMergedCells::getMergeRegion(TXTypes::RowIndex row, TXTypes::ColIndex col) const {
    uint64_t key = pImpl->getCellKey(row, col);
    auto it = pImpl->cellToRegionMap_.find(key);
    return (it != pImpl->cellToRegionMap_.end()) ? it->second : nullptr;
}

std::vector<TXMergedCells::MergeRegion> TXMergedCells::getAllMergeRegions() const {
    return std::vector<MergeRegion>(pImpl->mergeRegions_.begin(), pImpl->mergeRegions_.end());
}

std::size_t TXMergedCells::getMergeCount() const {
    return pImpl->mergeRegions_.size();
}

bool TXMergedCells::canMerge(TXTypes::RowIndex startRow, TXTypes::ColIndex startCol,
                            TXTypes::RowIndex endRow, TXTypes::ColIndex endCol) const {
    MergeRegion region(startRow, startCol, endRow, endCol);
    MergeRegion normalized = normalizeRegion(region);
    return isValidRegion(normalized) && pImpl->canMergeInternal(normalized);
}

std::vector<TXMergedCells::MergeRegion> TXMergedCells::getOverlappingRegions(const TXRange& range) const {
    return pImpl->getOverlappingRegions(range);
}

// ==================== 范围操作 ====================

std::vector<TXMergedCells::MergeRegion> TXMergedCells::getMergeRegionsInRange(const TXRange& range) const {
    std::vector<MergeRegion> result;
    MergeRegion queryRegion(range);
    
    for (const auto& region : pImpl->mergeRegions_) {
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
    return pImpl->batchMergeCells(regions);
}

std::size_t TXMergedCells::batchUnmergeCells(const std::vector<MergeRegion>& regions) {
    return pImpl->batchUnmergeCells(regions);
}

// ==================== 工具函数 ====================

void TXMergedCells::clear() {
    pImpl->clear();
}

bool TXMergedCells::empty() const {
    return pImpl->mergeRegions_.empty();
}

const std::string& TXMergedCells::getLastError() const {
    return pImpl->lastError_;
}

std::string TXMergedCells::toXml() const {
    return pImpl->toXml();
}

bool TXMergedCells::fromXml(const std::string& xmlStr) {
    return pImpl->fromXml(xmlStr);
}

} // namespace TinaXlsx 