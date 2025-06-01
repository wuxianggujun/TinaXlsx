#include "TinaXlsx/TXCellManager.hpp"
#include "TinaXlsx/TXRange.hpp"
#include <algorithm>
#include <limits>

namespace TinaXlsx {

// ==================== 单元格访问 ====================

TXCell* TXCellManager::getCell(const Coordinate& coord) {
    if (!isValidCoordinate(coord)) {
        return nullptr;
    }

    auto it = cells_.find(coord);
    if (it != cells_.end()) {
        return &it->second;
    }

    // 创建新单元格
    cells_[coord] = TXCell();
    return &cells_[coord];
}

const TXCell* TXCellManager::getCell(const Coordinate& coord) const {
    if (!isValidCoordinate(coord)) {
        return nullptr;
    }

    auto it = cells_.find(coord);
    return (it != cells_.end()) ? &it->second : nullptr;
}

bool TXCellManager::hasCell(const Coordinate& coord) const {
    return cells_.find(coord) != cells_.end();
}

bool TXCellManager::removeCell(const Coordinate& coord) {
    auto it = cells_.find(coord);
    if (it != cells_.end()) {
        cells_.erase(it);
        return true;
    }
    return false;
}

// ==================== 值操作 ====================

bool TXCellManager::setCellValue(const Coordinate& coord, const CellValue& value) {
    if (!isValidCoordinate(coord)) {
        return false;
    }

    auto* cell = getCell(coord);
    if (cell) {
        cell->setValue(value);
        return true;
    }
    return false;
}

TXCellManager::CellValue TXCellManager::getCellValue(const Coordinate& coord) const {
    const auto* cell = getCell(coord);
    if (cell) {
        return cell->getValue();
    }
    return std::string(""); // 默认返回空字符串
}

std::size_t TXCellManager::setCellValues(const std::vector<std::pair<Coordinate, CellValue>>& values) {
    std::size_t count = 0;
    for (const auto& pair : values) {
        if (setCellValue(pair.first, pair.second)) {
            ++count;
        }
    }
    return count;
}

std::vector<std::pair<TXCellManager::Coordinate, TXCellManager::CellValue>>
TXCellManager::getCellValues(const std::vector<Coordinate>& coords) const {
    std::vector<std::pair<Coordinate, CellValue>> result;
    result.reserve(coords.size());

    for (const auto& coord : coords) {
        result.emplace_back(coord, getCellValue(coord));
    }

    return result;
}

// ==================== 范围操作 ====================

TXRange TXCellManager::getUsedRange() const {
    if (cells_.empty()) {
        return TXRange(); // 返回默认范围
    }

    row_t min_row = row_t::last(), max_row = row_t(1);
    column_t min_col = column_t::last(), max_col = column_t(1);

    bool found_data = false;
    for (const auto& pair : cells_) {
        if (!pair.second.isEmpty()) {
            const auto& coord = pair.first;
            if (!found_data) {
                // 第一个有效单元格，初始化范围
                min_row = coord.getRow();
                max_row = coord.getRow();
                min_col = coord.getCol();
                max_col = coord.getCol();
                found_data = true;
            } else {
                // 扩展范围
                min_row = std::min(min_row, coord.getRow());
                max_row = std::max(max_row, coord.getRow());
                min_col = std::min(min_col, coord.getCol());
                max_col = std::max(max_col, coord.getCol());
            }
        }
    }

    if (!found_data) {
        return TXRange(); // 没有有效数据
    }

    return TXRange(Coordinate(min_row, min_col), Coordinate(max_row, max_col));
}

row_t TXCellManager::getMaxUsedRow() const {
    row_t max_row = row_t(0);
    for (const auto& pair : cells_) {
        if (!pair.second.isEmpty()) {
            max_row = std::max(max_row, pair.first.getRow());
        }
    }
    return max_row;
}

column_t TXCellManager::getMaxUsedColumn() const {
    column_t max_col = column_t(1);
    for (const auto& pair : cells_) {
        if (!pair.second.isEmpty()) {
            max_col = std::max(max_col, pair.first.getCol());
        }
    }
    return max_col;
}

void TXCellManager::clear() {
    cells_.clear();
}

std::size_t TXCellManager::getNonEmptyCellCount() const {
    std::size_t count = 0;
    for (const auto& pair : cells_) {
        if (!pair.second.isEmpty()) {
            ++count;
        }
    }
    return count;
}

// ==================== 行列移动支持 ====================

void TXCellManager::transformCells(std::function<Coordinate(const Coordinate&)> transform) {
    CellContainer new_cells;

    for (auto& pair : cells_) {
        const auto& old_coord = pair.first;
        auto new_coord = transform(old_coord);
        
        if (new_coord.isValid()) {
            new_cells[new_coord] = std::move(pair.second);
        }
        // 如果新坐标无效，单元格被丢弃（用于删除操作）
    }

    cells_ = std::move(new_cells);
}

std::size_t TXCellManager::removeCellsInRange(const TXRange& range) {
    if (!range.isValid()) {
        return 0;
    }

    std::size_t removed_count = 0;
    auto it = cells_.begin();
    
    while (it != cells_.end()) {
        if (range.contains(it->first)) {
            it = cells_.erase(it);
            ++removed_count;
        } else {
            ++it;
        }
    }

    return removed_count;
}

// ==================== 私有辅助方法 ====================

bool TXCellManager::isValidCoordinate(const Coordinate& coord) const {
    return coord.isValid();
}

} // namespace TinaXlsx
