#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXMergedCells.hpp"
#include <unordered_map>
#include <regex>
#include <algorithm>

namespace TinaXlsx {

// TXCoordinate的哈希函数特化
struct CoordinateHash {
    std::size_t operator()(const TXCoordinate& coord) const {
        return std::hash<TXTypes::RowIndex>()(coord.getRow()) ^ 
               (std::hash<TXTypes::ColIndex>()(coord.getCol()) << 1);
    }
};

class TXSheet::Impl {
public:
    explicit Impl(const std::string& name) : name_(name), last_error_(""), merged_cells_() {}

    const std::string& getName() const {
        return name_;
    }

    void setName(const std::string& name) {
        name_ = name;
    }

    TXSheet::CellValue getCellValue(const Coordinate& coord) const {
        auto it = cells_.find(coord);
        if (it != cells_.end()) {
            return it->second.getValue();
        }
        return std::string("");  // 默认返回空字符串
    }

    bool setCellValue(const Coordinate& coord, const CellValue& value) {
        if (!coord.isValid()) {
            last_error_ = "Invalid cell coordinate";
            return false;
        }

        cells_[coord].setValue(value);
        last_error_.clear();
        return true;
    }

    TXCell* getCell(const Coordinate& coord) {
        if (!coord.isValid()) {
            return nullptr;
        }

        auto it = cells_.find(coord);
        if (it != cells_.end()) {
            return &it->second;
        }

        // 创建新的单元格
        cells_[coord] = TXCell();
        return &cells_[coord];
    }

    const TXCell* getCell(const Coordinate& coord) const {
        auto it = cells_.find(coord);
        if (it != cells_.end()) {
            return &it->second;
        }
        return nullptr;
    }

    bool insertRows(TXTypes::RowIndex row, TXTypes::RowIndex count) {
        if (!TXTypes::isValidRow(row)) {
            last_error_ = "Invalid row number";
            return false;
        }

        // 需要移动的单元格
        std::unordered_map<Coordinate, TXCell, CoordinateHash> new_cells;
        
        for (auto& pair : cells_) {
            const auto& coord = pair.first;
            if (coord.getRow() >= row) {
                // 向下移动
                Coordinate new_coord(coord.getRow() + count, coord.getCol());
                new_cells[new_coord] = std::move(pair.second);
            } else {
                new_cells[coord] = std::move(pair.second);
            }
        }

        cells_ = std::move(new_cells);
        last_error_.clear();
        return true;
    }

    bool deleteRows(TXTypes::RowIndex row, TXTypes::RowIndex count) {
        if (!TXTypes::isValidRow(row)) {
            last_error_ = "Invalid row number";
            return false;
        }

        std::unordered_map<Coordinate, TXCell, CoordinateHash> new_cells;
        
        for (auto& pair : cells_) {
            const auto& coord = pair.first;
            if (coord.getRow() < row) {
                // 保持不变
                new_cells[coord] = std::move(pair.second);
            } else if (coord.getRow() >= row + count) {
                // 向上移动
                Coordinate new_coord(coord.getRow() - count, coord.getCol());
                new_cells[new_coord] = std::move(pair.second);
            }
            // 在删除范围内的单元格被丢弃
        }

        cells_ = std::move(new_cells);
        last_error_.clear();
        return true;
    }

    bool insertColumns(TXTypes::ColIndex col, TXTypes::ColIndex count) {
        if (!TXTypes::isValidCol(col)) {
            last_error_ = "Invalid column number";
            return false;
        }

        std::unordered_map<Coordinate, TXCell, CoordinateHash> new_cells;
        
        for (auto& pair : cells_) {
            const auto& coord = pair.first;
            if (coord.getCol() >= col) {
                // 向右移动
                Coordinate new_coord(coord.getRow(), coord.getCol() + count);
                new_cells[new_coord] = std::move(pair.second);
            } else {
                new_cells[coord] = std::move(pair.second);
            }
        }

        cells_ = std::move(new_cells);
        last_error_.clear();
        return true;
    }

    bool deleteColumns(TXTypes::ColIndex col, TXTypes::ColIndex count) {
        if (!TXTypes::isValidCol(col)) {
            last_error_ = "Invalid column number";
            return false;
        }

        std::unordered_map<Coordinate, TXCell, CoordinateHash> new_cells;
        
        for (auto& pair : cells_) {
            const auto& coord = pair.first;
            if (coord.getCol() < col) {
                // 保持不变
                new_cells[coord] = std::move(pair.second);
            } else if (coord.getCol() >= col + count) {
                // 向左移动
                Coordinate new_coord(coord.getRow(), coord.getCol() - count);
                new_cells[new_coord] = std::move(pair.second);
            }
            // 在删除范围内的单元格被丢弃
        }

        cells_ = std::move(new_cells);
        last_error_.clear();
        return true;
    }

    TXTypes::RowIndex getUsedRowCount() const {
        TXTypes::RowIndex max_row = 0;
        for (const auto& pair : cells_) {
            if (!pair.second.isEmpty()) {
                max_row = std::max(max_row, pair.first.getRow());
            }
        }
        return max_row;
    }

    TXTypes::ColIndex getUsedColumnCount() const {
        TXTypes::ColIndex max_col = 0;
        for (const auto& pair : cells_) {
            if (!pair.second.isEmpty()) {
                max_col = std::max(max_col, pair.first.getCol());
            }
        }
        return max_col;
    }

    TXSheet::Range getUsedRange() const {
        if (cells_.empty()) {
            return Range();  // 返回默认范围 A1:A1
        }

        TXTypes::RowIndex min_row = TXTypes::MAX_ROWS, max_row = 0;
        TXTypes::ColIndex min_col = TXTypes::MAX_COLS, max_col = 0;

        for (const auto& pair : cells_) {
            if (!pair.second.isEmpty()) {
                const auto& coord = pair.first;
                min_row = std::min(min_row, coord.getRow());
                max_row = std::max(max_row, coord.getRow());
                min_col = std::min(min_col, coord.getCol());
                max_col = std::max(max_col, coord.getCol());
            }
        }

        if (max_row == 0) {
            return Range();  // 没有有效数据
        }

        return Range(Coordinate(min_row, min_col), Coordinate(max_row, max_col));
    }

    void clear() {
        cells_.clear();
        merged_cells_.clear();
        last_error_.clear();
    }

    std::size_t setCellValues(const std::vector<std::pair<Coordinate, CellValue>>& values) {
        std::size_t count = 0;
        for (const auto& pair : values) {
            if (setCellValue(pair.first, pair.second)) {
                ++count;
            }
        }
        return count;
    }

    std::vector<std::pair<TXSheet::Coordinate, TXSheet::CellValue>> 
    getCellValues(const std::vector<Coordinate>& coords) const {
        std::vector<std::pair<Coordinate, CellValue>> result;
        result.reserve(coords.size());
        
        for (const auto& coord : coords) {
            result.emplace_back(coord, getCellValue(coord));
        }
        
        return result;
    }

    bool setRangeValues(const TXSheet::Range& range, const std::vector<std::vector<CellValue>>& values) {
        if (!range.isValid() || values.empty()) {
            last_error_ = "Invalid range or empty values";
            return false;
        }

        TXTypes::RowIndex row_count = range.getRowCount();
        TXTypes::ColIndex col_count = range.getColCount();

        if (values.size() != static_cast<size_t>(row_count)) {
            last_error_ = "Row count mismatch";
            return false;
        }

        for (TXTypes::RowIndex i = 0; i < row_count; ++i) {
            if (values[i].size() != static_cast<size_t>(col_count)) {
                last_error_ = "Column count mismatch at row " + std::to_string(i);
                return false;
            }
            
            for (TXTypes::ColIndex j = 0; j < col_count; ++j) {
                Coordinate coord(range.getStart().getRow() + i, 
                               range.getStart().getCol() + j);
                setCellValue(coord, values[i][j]);
            }
        }

        last_error_.clear();
        return true;
    }

    std::vector<std::vector<TXSheet::CellValue>> getRangeValues(const TXSheet::Range& range) const {
        TXTypes::RowIndex row_count = range.getRowCount();
        TXTypes::ColIndex col_count = range.getColCount();
        
        std::vector<std::vector<CellValue>> result(row_count);
        
        for (TXTypes::RowIndex i = 0; i < row_count; ++i) {
            result[i].resize(col_count);
            for (TXTypes::ColIndex j = 0; j < col_count; ++j) {
                Coordinate coord(range.getStart().getRow() + i, 
                               range.getStart().getCol() + j);
                result[i][j] = getCellValue(coord);
            }
        }
        
        return result;
    }

    const std::string& getLastError() const {
        return last_error_;
    }

    // ==================== 合并单元格功能实现 ====================

    bool mergeCells(const TXSheet::Range& range) {
        if (!range.isValid()) {
            last_error_ = "Invalid range for merging";
            return false;
        }

        // 检查范围是否已经有合并的单元格
        auto overlappingRegions = merged_cells_.getOverlappingRegions(range);
        if (!overlappingRegions.empty()) {
            last_error_ = "Range overlaps with existing merged cells";
            return false;
        }

        // 添加合并区域
        if (!merged_cells_.mergeCells(range)) {
            last_error_ = "Failed to add merge region";
            return false;
        }

        // 设置主单元格和从属单元格
        auto start = range.getStart();
        auto end = range.getEnd();

        for (TXTypes::RowIndex row = start.getRow(); row <= end.getRow(); ++row) {
            for (TXTypes::ColIndex col = start.getCol(); col <= end.getCol(); ++col) {
                Coordinate coord(row, col);
                TXCell* cell = getCell(coord);
                if (cell) {
                    cell->setMerged(true);
                    if (row == start.getRow() && col == start.getCol()) {
                        // 主单元格
                        cell->setMasterCell(true);
                    } else {
                        // 从属单元格
                        cell->setMasterCell(false);
                        cell->setMasterCellPosition(start.getRow(), start.getCol());
                    }
                }
            }
        }

        last_error_.clear();
        return true;
    }

    bool unmergeCells(TXTypes::RowIndex row, TXTypes::ColIndex col) {
        const auto* region = merged_cells_.getMergeRegion(row, col);
        if (!region) {
            last_error_ = "Cell is not in a merged region";
            return false;
        }

        // 移除合并区域
        if (!merged_cells_.unmergeCells(*region)) {
            last_error_ = "Failed to remove merge region";
            return false;
        }

        // 更新单元格状态
        for (TXTypes::RowIndex r = region->startRow; r <= region->endRow; ++r) {
            for (TXTypes::ColIndex c = region->startCol; c <= region->endCol; ++c) {
                Coordinate cell_coord(r, c);
                TXCell* cell = getCell(cell_coord);
                if (cell) {
                    cell->setMerged(false);
                    cell->setMasterCell(false);
                    cell->setMasterCellPosition(0, 0);
                }
            }
        }

        last_error_.clear();
        return true;
    }

    std::size_t unmergeCellsInRange(const TXSheet::Range& range) {
        return merged_cells_.unmergeCellsInRange(range);
    }

    bool isCellMerged(TXTypes::RowIndex row, TXTypes::ColIndex col) const {
        return merged_cells_.isMerged(row, col);
    }

    TXSheet::Range getMergeRegion(TXTypes::RowIndex row, TXTypes::ColIndex col) const {
        const auto* region = merged_cells_.getMergeRegion(row, col);
        if (region) {
            return TXSheet::Range(region->startRow, region->startCol, region->endRow, region->endCol);
        }
        return TXSheet::Range(); // 返回无效范围
    }

    std::vector<TXSheet::Range> getAllMergeRegions() const {
        auto regions = merged_cells_.getAllMergeRegions();
        std::vector<TXSheet::Range> result;
        result.reserve(regions.size());
        for (const auto& region : regions) {
            result.emplace_back(region.startRow, region.startCol, region.endRow, region.endCol);
        }
        return result;
    }

    std::size_t getMergeCount() const {
        return merged_cells_.getMergeCount();
    }

    // ==================== 公式功能实现 ====================
    
    std::size_t calculateAllFormulas() {
        std::size_t count = 0;
        for (auto& pair : cells_) {
            TXCell& cell = pair.second;
            if (cell.isFormula()) {
                // 注意：这里暂时简化，实际上需要解决循环依赖等问题
                count++;
            }
        }
        return count;
    }

    std::size_t calculateFormulasInRange(const TXSheet::Range& range) {
        std::size_t count = 0;
        auto start = range.getStart();
        auto end = range.getEnd();
        
        for (TXTypes::RowIndex row = start.getRow(); row <= end.getRow(); ++row) {
            for (TXTypes::ColIndex col = start.getCol(); col <= end.getCol(); ++col) {
                Coordinate coord(row, col);
                auto it = cells_.find(coord);
                if (it != cells_.end() && it->second.isFormula()) {
                    // 注意：这里暂时简化，实际上需要解决循环依赖等问题
                    count++;
                }
            }
        }
        return count;
    }

private:
    std::string name_;
    std::unordered_map<Coordinate, TXCell, CoordinateHash> cells_;
    mutable std::string last_error_;
    TXMergedCells merged_cells_;
};

// TXSheet 实现
TXSheet::TXSheet(const std::string& name) : pImpl(std::make_unique<Impl>(name)) {}

TXSheet::~TXSheet() = default;

TXSheet::TXSheet(TXSheet&& other) noexcept : pImpl(std::move(other.pImpl)) {}

TXSheet& TXSheet::operator=(TXSheet&& other) noexcept {
    if (this != &other) {
        pImpl = std::move(other.pImpl);
    }
    return *this;
}

const std::string& TXSheet::getName() const {
    return pImpl->getName();
}

void TXSheet::setName(const std::string& name) {
    pImpl->setName(name);
}

TXSheet::CellValue TXSheet::getCellValue(TXTypes::RowIndex row, TXTypes::ColIndex col) const {
    return pImpl->getCellValue(Coordinate(row, col));
}

TXSheet::CellValue TXSheet::getCellValue(const Coordinate& coord) const {
    return pImpl->getCellValue(coord);
}

TXSheet::CellValue TXSheet::getCellValue(const std::string& address) const {
    return pImpl->getCellValue(Coordinate::fromAddress(address));
}

bool TXSheet::setCellValue(TXTypes::RowIndex row, TXTypes::ColIndex col, const CellValue& value) {
    return pImpl->setCellValue(Coordinate(row, col), value);
}

bool TXSheet::setCellValue(const Coordinate& coord, const CellValue& value) {
    return pImpl->setCellValue(coord, value);
}

bool TXSheet::setCellValue(const std::string& address, const CellValue& value) {
    return pImpl->setCellValue(Coordinate::fromAddress(address), value);
}

TXCell* TXSheet::getCell(TXTypes::RowIndex row, TXTypes::ColIndex col) {
    return pImpl->getCell(Coordinate(row, col));
}

const TXCell* TXSheet::getCell(TXTypes::RowIndex row, TXTypes::ColIndex col) const {
    return pImpl->getCell(Coordinate(row, col));
}

TXCell* TXSheet::getCell(const Coordinate& coord) {
    return pImpl->getCell(coord);
}

const TXCell* TXSheet::getCell(const Coordinate& coord) const {
    return pImpl->getCell(coord);
}

TXCell* TXSheet::getCell(const std::string& address) {
    return pImpl->getCell(Coordinate::fromAddress(address));
}

const TXCell* TXSheet::getCell(const std::string& address) const {
    return pImpl->getCell(Coordinate::fromAddress(address));
}

bool TXSheet::insertRows(TXTypes::RowIndex row, TXTypes::RowIndex count) {
    return pImpl->insertRows(row, count);
}

bool TXSheet::deleteRows(TXTypes::RowIndex row, TXTypes::RowIndex count) {
    return pImpl->deleteRows(row, count);
}

bool TXSheet::insertColumns(TXTypes::ColIndex col, TXTypes::ColIndex count) {
    return pImpl->insertColumns(col, count);
}

bool TXSheet::deleteColumns(TXTypes::ColIndex col, TXTypes::ColIndex count) {
    return pImpl->deleteColumns(col, count);
}

TXTypes::RowIndex TXSheet::getUsedRowCount() const {
    return pImpl->getUsedRowCount();
}

TXTypes::ColIndex TXSheet::getUsedColumnCount() const {
    return pImpl->getUsedColumnCount();
}

TXSheet::Range TXSheet::getUsedRange() const {
    return pImpl->getUsedRange();
}

void TXSheet::clear() {
    pImpl->clear();
}

std::size_t TXSheet::setCellValues(const std::vector<std::pair<Coordinate, CellValue>>& values) {
    return pImpl->setCellValues(values);
}

std::vector<std::pair<TXSheet::Coordinate, TXSheet::CellValue>> 
TXSheet::getCellValues(const std::vector<Coordinate>& coords) const {
    return pImpl->getCellValues(coords);
}

bool TXSheet::setRangeValues(const Range& range, const std::vector<std::vector<CellValue>>& values) {
    return pImpl->setRangeValues(range, values);
}

std::vector<std::vector<TXSheet::CellValue>> TXSheet::getRangeValues(const Range& range) const {
    return pImpl->getRangeValues(range);
}

const std::string& TXSheet::getLastError() const {
    return pImpl->getLastError();
}

// ==================== 合并单元格功能公共接口 ====================

bool TXSheet::mergeCells(TXTypes::RowIndex startRow, TXTypes::ColIndex startCol,
                        TXTypes::RowIndex endRow, TXTypes::ColIndex endCol) {
    Range range(Coordinate(startRow, startCol), Coordinate(endRow, endCol));
    return pImpl->mergeCells(range);
}

bool TXSheet::mergeCells(const Range& range) {
    return pImpl->mergeCells(range);
}

bool TXSheet::mergeCells(const std::string& rangeStr) {
    try {
        Range range = addressToRange(rangeStr);
        return pImpl->mergeCells(range);
    } catch (const std::exception& e) {
        return false;
    }
}

bool TXSheet::unmergeCells(TXTypes::RowIndex row, TXTypes::ColIndex col) {
    return pImpl->unmergeCells(row, col);
}

std::size_t TXSheet::unmergeCellsInRange(const Range& range) {
    return pImpl->unmergeCellsInRange(range);
}

bool TXSheet::isCellMerged(TXTypes::RowIndex row, TXTypes::ColIndex col) const {
    return pImpl->isCellMerged(row, col);
}

TXSheet::Range TXSheet::getMergeRegion(TXTypes::RowIndex row, TXTypes::ColIndex col) const {
    return pImpl->getMergeRegion(row, col);
}

std::vector<TXSheet::Range> TXSheet::getAllMergeRegions() const {
    return pImpl->getAllMergeRegions();
}

std::size_t TXSheet::getMergeCount() const {
    return pImpl->getMergeCount();
}

// ==================== 公式功能公共接口 ====================

std::size_t TXSheet::calculateAllFormulas() {
    return pImpl->calculateAllFormulas();
}

std::size_t TXSheet::calculateFormulasInRange(const Range& range) {
    return pImpl->calculateFormulasInRange(range);
}

bool TXSheet::setCellFormula(TXTypes::RowIndex row, TXTypes::ColIndex col, const std::string& formula) {
    TXCell* cell = getCell(row, col);
    if (!cell) {
        return false;
    }
    
    cell->setFormula(formula);
    return true;
}

std::string TXSheet::getCellFormula(TXTypes::RowIndex row, TXTypes::ColIndex col) const {
    const TXCell* cell = getCell(row, col);
    if (!cell) {
        return "";
    }
    
    return cell->getFormula();
}

std::size_t TXSheet::setCellFormulas(const std::vector<std::pair<Coordinate, std::string>>& formulas) {
    std::size_t count = 0;
    for (const auto& pair : formulas) {
        if (setCellFormula(pair.first.getRow(), pair.first.getCol(), pair.second)) {
            count++;
        }
    }
    return count;
}

// ==================== 数字格式化功能公共接口 ====================

bool TXSheet::setCellNumberFormat(TXTypes::RowIndex row, TXTypes::ColIndex col, 
                                 TXCell::NumberFormat formatType, int decimalPlaces) {
    TXCell* cell = getCell(row, col);
    if (!cell) {
        return false;
    }
    
    cell->setPredefinedFormat(formatType, decimalPlaces);
    return true;
}

bool TXSheet::setCellCustomFormat(TXTypes::RowIndex row, TXTypes::ColIndex col, 
                                 const std::string& formatString) {
    TXCell* cell = getCell(row, col);
    if (!cell) {
        return false;
    }
    
    cell->setCustomFormat(formatString);
    return true;
}

std::size_t TXSheet::setRangeNumberFormat(const Range& range, TXCell::NumberFormat formatType, 
                                         int decimalPlaces) {
    std::size_t count = 0;
    auto start = range.getStart();
    auto end = range.getEnd();
    
    for (TXTypes::RowIndex row = start.getRow(); row <= end.getRow(); ++row) {
        for (TXTypes::ColIndex col = start.getCol(); col <= end.getCol(); ++col) {
            if (setCellNumberFormat(row, col, formatType, decimalPlaces)) {
                count++;
            }
        }
    }
    return count;
}

std::string TXSheet::getCellFormattedValue(TXTypes::RowIndex row, TXTypes::ColIndex col) const {
    const TXCell* cell = getCell(row, col);
    if (!cell) {
        return "";
    }
    
    return cell->getFormattedValue();
}

std::size_t TXSheet::setCellFormats(const std::vector<std::pair<Coordinate, TXCell::NumberFormat>>& formats) {
    std::size_t count = 0;
    for (const auto& pair : formats) {
        if (setCellNumberFormat(pair.first.getRow(), pair.first.getCol(), pair.second)) {
            count++;
        }
    }
    return count;
}

} // namespace TinaXlsx 