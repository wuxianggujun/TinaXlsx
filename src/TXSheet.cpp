#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXMergedCells.hpp"
#include "TinaXlsx/TXStyle.hpp"
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXWorkbookContext.hpp"
#include <unordered_map>
#include <regex>
#include <algorithm>
#include <utility>

namespace TinaXlsx
{
    // ==================== TXSheet 实现 ====================

    TXSheet::TXSheet(const std::string& name, TXWorkbook* parentWorkbook)
        : name_(name), workbook_(parentWorkbook)
    {
    }

    TXSheet::~TXSheet() = default;

    TXSheet::TXSheet(TXSheet&& other) noexcept
        : name_(std::move(other.name_))
          , cells_(std::move(other.cells_))
          , lastError_(std::move(other.lastError_))
          , mergedCells_(std::move(other.mergedCells_))
          , workbook_(other.workbook_)
    {
        other.workbook_ = nullptr;
    }

    TXSheet& TXSheet::operator=(TXSheet&& other) noexcept
    {
        if (this != &other)
        {
            name_ = std::move(other.name_);
            cells_ = std::move(other.cells_);
            lastError_ = std::move(other.lastError_);
            mergedCells_ = std::move(other.mergedCells_);
            workbook_ = other.workbook_;
            other.workbook_ = nullptr;
        }
        return *this;
    }

    const std::string& TXSheet::getName() const
    {
        return name_;
    }

    void TXSheet::setName(const std::string& name)
    {
        name_ = name;
    }

    // ==================== 单元格值操作实现 ====================

    TXSheet::CellValue TXSheet::getCellValue(row_t row, column_t col) const
    {
        return getCellValue(Coordinate(row, col));
    }

    TXSheet::CellValue TXSheet::getCellValue(const Coordinate& coord) const
    {
        auto it = cells_.find(coord);
        if (it != cells_.end())
        {
            return it->second.getValue();
        }
        return std::string(""); // 默认返回空字符串
    }

    TXSheet::CellValue TXSheet::getCellValue(const std::string& address) const
    {
        return getCellValue(Coordinate::fromAddress(address));
    }

    bool TXSheet::setCellValue(row_t row, column_t col, const CellValue& value)
    {
        return setCellValue(Coordinate(row, col), value);
    }

    bool TXSheet::setCellValue(const Coordinate& coord, const CellValue& value)
    {
        if (!coord.isValid())
        {
            lastError_ = "Invalid cell coordinate";
            return false;
        }

        cells_[coord].setValue(value);
        lastError_.clear();
        return true;
    }

    bool TXSheet::setCellValue(const std::string& address, const CellValue& value)
    {
        return setCellValue(Coordinate::fromAddress(address), value);
    }

    // ==================== 单元格访问实现 ====================

    TXCell* TXSheet::getCell(row_t row, column_t col)
    {
        return getCell(Coordinate(row, col));
    }

    const TXCell* TXSheet::getCell(row_t row, column_t col) const
    {
        return getCell(Coordinate(row, col));
    }

    TXCell* TXSheet::getCell(const Coordinate& coord)
    {
        return getCellInternal(coord);
    }

    const TXCell* TXSheet::getCell(const Coordinate& coord) const
    {
        return getCellInternal(coord);
    }

    TXCell* TXSheet::getCell(const std::string& address)
    {
        return getCell(Coordinate::fromAddress(address));
    }

    const TXCell* TXSheet::getCell(const std::string& address) const
    {
        return getCell(Coordinate::fromAddress(address));
    }

    TXCell* TXSheet::getCellInternal(const Coordinate& coord)
    {
        if (!coord.isValid())
        {
            return nullptr;
        }

        auto it = cells_.find(coord);
        if (it != cells_.end())
        {
            return &it->second;
        }

        // 创建新单元格
        cells_[coord] = TXCell();
        return &cells_[coord];
    }

    const TXCell* TXSheet::getCellInternal(const Coordinate& coord) const
    {
        auto it = cells_.find(coord);
        return (it != cells_.end()) ? &it->second : nullptr;
    }

    bool TXSheet::insertRows(row_t row, row_t count)
    {
        if (!row.is_valid())
        {
            lastError_ = "Invalid row number";
            return false;
        }

        // 需要移动的单元格 - 使用正确的哈希函数类型
        std::unordered_map<Coordinate, TXCell, CoordinateHash> new_cells;

        for (auto& pair : cells_)
        {
            const auto& coord = pair.first;
            if (coord.getRow() >= row)
            {
                // 向下移动
                Coordinate new_coord(row_t(coord.getRow().index() + count.index()), coord.getCol());
                new_cells[new_coord] = std::move(pair.second);
            }
            else
            {
                new_cells[coord] = std::move(pair.second);
            }
        }

        cells_ = std::move(new_cells);
        lastError_.clear();
        return true;
    }

    bool TXSheet::deleteRows(row_t row, row_t count)
    {
        if (!row.is_valid())
        {
            lastError_ = "Invalid row number";
            return false;
        }

        std::unordered_map<Coordinate, TXCell, CoordinateHash> new_cells;

        for (auto& pair : cells_)
        {
            const auto& coord = pair.first;
            if (coord.getRow() < row)
            {
                // 保持不变
                new_cells[coord] = std::move(pair.second);
            }
            else if (coord.getRow().index() >= row.index() + count.index())
            {
                // 向上移动
                Coordinate new_coord(row_t(coord.getRow().index() - count.index()), coord.getCol());
                new_cells[new_coord] = std::move(pair.second);
            }
            // 在删除范围内的单元格被丢弃
        }

        cells_ = std::move(new_cells);
        lastError_.clear();
        return true;
    }

    bool TXSheet::insertColumns(column_t col, column_t count)
    {
        if (!col.is_valid())
        {
            lastError_ = "Invalid column number";
            return false;
        }

        std::unordered_map<Coordinate, TXCell, CoordinateHash> new_cells;

        for (auto& pair : cells_)
        {
            const auto& coord = pair.first;
            if (coord.getCol() >= col)
            {
                // 向右移动
                Coordinate new_coord(coord.getRow(), column_t(coord.getCol().index() + count.index()));
                new_cells[new_coord] = std::move(pair.second);
            }
            else
            {
                new_cells[coord] = std::move(pair.second);
            }
        }

        cells_ = std::move(new_cells);
        lastError_.clear();
        return true;
    }

    bool TXSheet::deleteColumns(column_t col, column_t count)
    {
        if (!col.is_valid())
        {
            lastError_ = "Invalid column number";
            return false;
        }

        std::unordered_map<Coordinate, TXCell, CoordinateHash> new_cells;

        for (auto& pair : cells_)
        {
            const auto& coord = pair.first;
            if (coord.getCol() < col)
            {
                // 保持不变
                new_cells[coord] = std::move(pair.second);
            }
            else if (coord.getCol().index() >= col.index() + count.index())
            {
                // 向左移动
                Coordinate new_coord(coord.getRow(), column_t(coord.getCol().index() - count.index()));
                new_cells[new_coord] = std::move(pair.second);
            }
            // 在删除范围内的单元格被丢弃
        }

        cells_ = std::move(new_cells);
        lastError_.clear();
        return true;
    }

    row_t TXSheet::getUsedRowCount() const
    {
        row_t max_row = row_t(0);
        for (const auto& pair : cells_)
        {
            if (!pair.second.isEmpty())
            {
                max_row = std::max(max_row, pair.first.getRow());
            }
        }
        return max_row;
    }

    column_t TXSheet::getUsedColumnCount() const
    {
        column_t max_col = column_t(1);
        for (const auto& pair : cells_)
        {
            if (!pair.second.isEmpty())
            {
                max_col = std::max(max_col, pair.first.getCol());
            }
        }
        return max_col;
    }

    TXSheet::Range TXSheet::getUsedRange() const
    {
        if (cells_.empty())
        {
            return Range(); // 返回默认范围 A1:A1
        }

        row_t min_row = row_t::last(), max_row = row_t(1);
        column_t min_col = column_t::last(), max_col = column_t(1);

        bool found_data = false;
        for (const auto& pair : cells_)
        {
            if (!pair.second.isEmpty())
            {
                const auto& coord = pair.first;
                if (!found_data)
                {
                    // 第一个有效单元格，初始化范围
                    min_row = coord.getRow();
                    max_row = coord.getRow();
                    min_col = coord.getCol();
                    max_col = coord.getCol();
                    found_data = true;
                }
                else
                {
                    // 扩展范围
                    min_row = std::min(min_row, coord.getRow());
                    max_row = std::max(max_row, coord.getRow());
                    min_col = std::min(min_col, coord.getCol());
                    max_col = std::max(max_col, coord.getCol());
                }
            }
        }

        if (!found_data)
        {
            return Range(); // 没有有效数据
        }

        return Range(Coordinate(min_row, min_col), Coordinate(max_row, max_col));
    }

    void TXSheet::clear()
    {
        cells_.clear();
        mergedCells_.clear();
        lastError_.clear();
    }

    std::size_t TXSheet::setCellValues(const std::vector<std::pair<Coordinate, CellValue>>& values)
    {
        std::size_t count = 0;
        for (const auto& pair : values)
        {
            if (setCellValue(pair.first, pair.second))
            {
                ++count;
            }
        }
        return count;
    }

    std::vector<std::pair<TXSheet::Coordinate, TXSheet::CellValue>>
    TXSheet::getCellValues(const std::vector<Coordinate>& coords) const
    {
        std::vector<std::pair<Coordinate, CellValue>> result;
        result.reserve(coords.size());

        for (const auto& coord : coords)
        {
            result.emplace_back(coord, getCellValue(coord));
        }

        return result;
    }

    bool TXSheet::setRangeValues(const TXSheet::Range& range, const std::vector<std::vector<CellValue>>& values)
    {
        if (!range.isValid() || values.empty())
        {
            lastError_ = "Invalid range or empty values";
            return false;
        }

        row_t row_count = range.getRowCount();
        column_t col_count = range.getColCount();

        if (values.size() != static_cast<size_t>(row_count.index()))
        {
            lastError_ = "Row count mismatch";
            return false;
        }

        for (u32 i = 0; i < row_count.index(); ++i)
        {
            if (values[i].size() != static_cast<size_t>(col_count.index()))
            {
                lastError_ = "Column count mismatch at row " + std::to_string(i + 1);
                return false;
            }

            for (u32 j = 0; j < col_count.index(); ++j)
            {
                Coordinate coord(row_t(range.getStart().getRow().index() + static_cast<u32>(i)),
                                 column_t(range.getStart().getCol().index() + static_cast<u32>(j)));
                setCellValue(coord, values[i][j]);
            }
        }

        lastError_.clear();
        return true;
    }

    std::vector<std::vector<TXSheet::CellValue>> TXSheet::getRangeValues(const TXSheet::Range& range) const
    {
        row_t row_count = range.getRowCount();
        column_t col_count = range.getColCount();

        std::vector<std::vector<CellValue>> result(row_count.index());

        for (u32 i = 0; i < row_count.index(); ++i)
        {
            result[i].resize(col_count.index());
            for (u32 j = 0; j < col_count.index(); ++j)
            {
                Coordinate coord(row_t(range.getStart().getRow().index() + static_cast<u32>(i)),
                                 column_t(range.getStart().getCol().index() + static_cast<u32>(j)));
                result[i][j] = getCellValue(coord);
            }
        }

        return result;
    }

    const std::string& TXSheet::getLastError() const
    {
        return lastError_;
    }

    // ==================== 合并单元格功能实现 ====================

    bool TXSheet::mergeCells(const TXSheet::Range& range)
    {
        if (!range.isValid())
        {
            lastError_ = "Invalid range for merging";
            return false;
        }

        // 检查范围是否已经有合并的单元格
        auto overlappingRegions = mergedCells_.getOverlappingRegions(range);
        if (!overlappingRegions.empty())
        {
            lastError_ = "Range overlaps with existing merged cells";
            return false;
        }

        // 添加合并区域
        if (!mergedCells_.mergeCells(range))
        {
            lastError_ = "Failed to add merge region";
            return false;
        }

        workbook_->getContext()->registerComponentFast(ExcelComponent::MergedCells);

        // 设置主单元格和从属单元格
        auto start = range.getStart();
        auto end = range.getEnd();

        for (row_t row = start.getRow(); row <= end.getRow(); ++row)
        {
            for (column_t col = start.getCol(); col <= end.getCol(); ++col)
            {
                Coordinate coord(row, col);
                TXCell* cell = getCell(coord);
                if (cell)
                {
                    cell->setMerged(true);
                    if (row == start.getRow() && col == start.getCol())
                    {
                        // 主单元格
                        cell->setMasterCell(true);
                    }
                    else
                    {
                        // 从属单元格
                        cell->setMasterCell(false);
                        cell->setMasterCellPosition(start.getRow().index(), start.getCol().index());
                    }
                }
            }
        }

        lastError_.clear();
        return true;
    }

    bool TXSheet::unmergeCells(row_t row, column_t col)
    {
        const auto* region = mergedCells_.getMergeRegion(row, col);
        if (!region)
        {
            lastError_ = "Cell is not in a merged region";
            return false;
        }

        // 移除合并区域
        if (!mergedCells_.unmergeCells(*region))
        {
            lastError_ = "Failed to remove merge region";
            return false;
        }

        // 更新单元格状态
        for (row_t r = region->startRow; r <= region->endRow; ++r)
        {
            for (column_t c = region->startCol; c <= region->endCol; ++c)
            {
                Coordinate cell_coord(r, c);
                TXCell* cell = getCell(cell_coord);
                if (cell)
                {
                    cell->setMerged(false);
                    cell->setMasterCell(false);
                    cell->setMasterCellPosition(0, 0);
                }
            }
        }

        lastError_.clear();
        return true;
    }

    std::size_t TXSheet::unmergeCellsInRange(const TXSheet::Range& range)
    {
        return mergedCells_.unmergeCellsInRange(range);
    }

    bool TXSheet::isCellMerged(row_t row, column_t col) const
    {
        return mergedCells_.isMerged(row, col);
    }

    TXSheet::Range TXSheet::getMergeRegion(row_t row, column_t col) const
    {
        const auto* region = mergedCells_.getMergeRegion(row, col);
        if (region)
        {
            TXCoordinate start(region->startRow, region->startCol);
            TXCoordinate end(region->endRow, region->endCol);
            return TXSheet::Range(start, end);
        }
        return TXSheet::Range(); // 返回无效范围
    }

    std::vector<TXSheet::Range> TXSheet::getAllMergeRegions() const
    {
        auto regions = mergedCells_.getAllMergeRegions();
        std::vector<TXSheet::Range> result;
        result.reserve(regions.size());
        for (const auto& region : regions)
        {
            result.emplace_back(TXCoordinate(region.startRow, region.startCol),
                                TXCoordinate(region.endRow, region.endCol));
        }
        return result;
    }

    std::size_t TXSheet::getMergeCount() const
    {
        return mergedCells_.getMergeCount();
    }

    // ==================== 公式功能实现 ====================

    std::size_t TXSheet::calculateAllFormulas()
    {
        std::size_t count = 0;
        for (auto& pair : cells_)
        {
            TXCell& cell = pair.second;
            if (cell.isFormula())
            {
                // 注意：这里暂时简化，实际上需要解决循环依赖等问题
                count++;
            }
        }
        return count;
    }

    std::size_t TXSheet::calculateFormulasInRange(const TXSheet::Range& range)
    {
        std::size_t count = 0;
        auto start = range.getStart();
        auto end = range.getEnd();

        for (row_t row = start.getRow(); row <= end.getRow(); ++row)
        {
            for (column_t col = start.getCol(); col <= end.getCol(); ++col)
            {
                Coordinate coord(row, col);
                auto it = cells_.find(coord);
                if (it != cells_.end() && it->second.isFormula())
                {
                    // 注意：这里暂时简化，实际上需要解决循环依赖等问题
                    count++;
                }
            }
        }
        return count;
    }

    bool TXSheet::setCellFormula(row_t row, column_t col, const std::string& formula)
    {
        TXCell* cell = getCell(row, col);
        if (!cell)
        {
            return false;
        }

        cell->setFormula(formula);
        return true;
    }

    std::string TXSheet::getCellFormula(row_t row, column_t col) const
    {
        const TXCell* cell = getCell(row, col);
        if (!cell)
        {
            return "";
        }

        return cell->getFormula();
    }

    std::size_t TXSheet::setCellFormulas(const std::vector<std::pair<Coordinate, std::string>>& formulas)
    {
        std::size_t count = 0;
        for (const auto& pair : formulas)
        {
            if (setCellFormula(pair.first.getRow(), pair.first.getCol(), pair.second))
            {
                count++;
            }
        }
        return count;
    }

    // ==================== 数字格式化功能实现 ====================

    bool TXSheet::setCellNumberFormat(row_t row, column_t col,
                                      TXNumberFormat::FormatType formatType, int decimalPlaces)
    {
        // 修正参数类型
        TXCell* cell = getCell(row, col);
        if (!cell)
        {
            return false;
        }

        cell->setPredefinedFormat(formatType, decimalPlaces);
        return true;
    }


    bool TXSheet::setCellCustomFormat(row_t row, column_t col,
                                      const std::string& formatString)
    {
        TXCell* cell = getCell(row, col);
        if (!cell)
        {
            return false;
        }

        cell->setCustomFormat(formatString);
        return true;
    }
    
    std::size_t TXSheet::setRangeNumberFormat(const Range& range, TXNumberFormat::FormatType formatType, 
                                          int decimalPlaces) {
        std::size_t count = 0;
        auto start = range.getStart();
        auto end = range.getEnd();

        for (row_t r = start.getRow(); r <= end.getRow(); ++r) { 
            for (column_t c = start.getCol(); c <= end.getCol(); ++c) { 
                if (setCellNumberFormat(r, c, formatType, decimalPlaces)) {
                    count++;
                }
            }
        }
        return count;
    }
    

    std::string TXSheet::getCellFormattedValue(row_t row, column_t col) const
    {
        const TXCell* cell = getCell(row, col);
        if (!cell)
        {
            return "";
        }

        return cell->getFormattedValue();
    }

    std::size_t TXSheet::setCellFormats(const std::vector<std::pair<Coordinate, TXNumberFormat::FormatType>>& formats) {
        std::size_t count = 0;
        for (const auto& pair : formats) {
            if (setCellNumberFormat(pair.first.getRow(), pair.first.getCol(), pair.second)) { // 假设使用 setCellNumberFormat 的默认 decimalPlaces
                count++;
            }
        }
        return count;
    }

    // ==================== 样式功能实现 ====================

    bool TXSheet::setCellStyle(row_t row, column_t col, const TXCellStyle& style)
    {
        if (!workbook_) return false;
        workbook_->getContext()->registerComponentFast(ExcelComponent::Styles);

        u32 style_id = workbook_->registerOrGetStyleFId(style);

        TXCell* cell = getCell(row, col);
        if (!cell)
        {
            return false;
        }
        cell->setStyleIndex(style_id);
        return true;
    }

    bool TXSheet::setCellStyle(const std::string& address, const TXCellStyle& style)
    {
        auto coord = Coordinate::fromAddress(address);
        return setCellStyle(coord.getRow(), coord.getCol(), style);
    }

    std::size_t TXSheet::setRangeStyle(const Range& range, const TXCellStyle& style)
    {
        std::size_t count = 0;
        auto start = range.getStart();
        auto end = range.getEnd();

        for (row_t row = start.getRow(); row <= end.getRow(); ++row)
        {
            for (column_t col = start.getCol(); col <= end.getCol(); ++col)
            {
                if (setCellStyle(row, col, style))
                {
                    count++;
                }
            }
        }
        return count;
    }

    std::size_t TXSheet::setCellStyles(const std::vector<std::pair<Coordinate, TXCellStyle>>& styles)
    {
        std::size_t count = 0;
        for (const auto& pair : styles)
        {
            if (setCellStyle(pair.first.getRow(), pair.first.getCol(), pair.second))
            {
                count++;
            }
        }
        return count;
    }

    // ==================== 合并单元格功能公共接口 ====================

    bool TXSheet::mergeCells(row_t startRow, column_t startCol,
                             row_t endRow, column_t endCol)
    {
        Range range(Coordinate(startRow, startCol), Coordinate(endRow, endCol));
        return mergeCells(range);
    }

    bool TXSheet::mergeCells(const std::string& rangeStr)
    {
        try
        {
            Range range = addressToRange(rangeStr);
            return mergeCells(range);
        }
        catch (const std::exception& e)
        {
            return false;
        }
    }

    void TXSheet::updateUsedRange()
    {
        // 内部辅助方法，用于优化已使用范围的计算
        // 这里可以添加缓存逻辑
    }
} // namespace TinaXlsx 
