#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXMergedCells.hpp"
#include "TinaXlsx/TXStyle.hpp"
#include <unordered_map>
#include <regex>
#include <algorithm>
#include <utility>

#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXWorkbookContext.hpp"

namespace TinaXlsx
{
    // TXCoordinate的哈希函数特化
    struct CoordinateHash
    {
        std::size_t operator()(const TXCoordinate& coord) const
        {
            return std::hash<row_t>()(coord.getRow()) ^
                (std::hash<column_t>()(coord.getCol()) << 1);
        }
    };

    class TXSheet::Impl
    {
    public:
        explicit Impl(std::string name, TXWorkbook* workbook) : name_(std::move(name)), workbook_(workbook)
        {
        }

        const std::string& getName() const
        {
            return name_;
        }

        void setName(const std::string& name)
        {
            name_ = name;
        }

        [[nodiscard]] TXWorkbook* getWorkbook() const
        {
            return workbook_;
        }

        [[nodiscard]] CellValue getCellValue(const Coordinate& coord) const
        {
            auto it = cells_.find(coord);
            if (it != cells_.end())
            {
                return it->second.getValue();
            }
            return std::string(""); // 默认返回空字符串
        }

        bool setCellValue(const Coordinate& coord, const CellValue& value)
        {
            if (!coord.isValid())
            {
                last_error_ = "Invalid cell coordinate";
                return false;
            }

            if (std::holds_alternative<std::string>(value))
            {
                workbook_->getContext()->registerComponentFast(ExcelComponent::SharedStrings);
            }

            cells_[coord].setValue(value);
            last_error_.clear();
            return true;
        }

        TXCell* getCell(const Coordinate& coord)
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

            // 创建新的单元格
            cells_[coord] = TXCell();
            return &cells_[coord];
        }

        const TXCell* getCell(const Coordinate& coord) const
        {
            auto it = cells_.find(coord);
            if (it != cells_.end())
            {
                return &it->second;
            }
            return nullptr;
        }

        bool insertRows(row_t row, row_t count)
        {
            if (!row.is_valid())
            {
                last_error_ = "Invalid row number";
                return false;
            }

            // 需要移动的单元格
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
            last_error_.clear();
            return true;
        }

        bool deleteRows(row_t row, row_t count)
        {
            if (!row.is_valid())
            {
                last_error_ = "Invalid row number";
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
            last_error_.clear();
            return true;
        }

        bool insertColumns(column_t col, column_t count)
        {
            if (!col.is_valid())
            {
                last_error_ = "Invalid column number";
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
            last_error_.clear();
            return true;
        }

        bool deleteColumns(column_t col, column_t count)
        {
            if (!col.is_valid())
            {
                last_error_ = "Invalid column number";
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
            last_error_.clear();
            return true;
        }

        row_t getUsedRowCount() const
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

        column_t getUsedColumnCount() const
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

        TXSheet::Range getUsedRange() const
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

        void clear()
        {
            cells_.clear();
            merged_cells_.clear();
            last_error_.clear();
        }

        std::size_t setCellValues(const std::vector<std::pair<Coordinate, CellValue>>& values)
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
        getCellValues(const std::vector<Coordinate>& coords) const
        {
            std::vector<std::pair<Coordinate, CellValue>> result;
            result.reserve(coords.size());

            for (const auto& coord : coords)
            {
                result.emplace_back(coord, getCellValue(coord));
            }

            return result;
        }

        bool setRangeValues(const TXSheet::Range& range, const std::vector<std::vector<CellValue>>& values)
        {
            if (!range.isValid() || values.empty())
            {
                last_error_ = "Invalid range or empty values";
                return false;
            }

            row_t row_count = range.getRowCount();
            column_t col_count = range.getColCount();

            if (values.size() != static_cast<size_t>(row_count.index()))
            {
                last_error_ = "Row count mismatch";
                return false;
            }

            for (u32 i = 0; i < row_count.index(); ++i)
            {
                if (values[i].size() != static_cast<size_t>(col_count.index()))
                {
                    last_error_ = "Column count mismatch at row " + std::to_string(i + 1);
                    return false;
                }

                for (u32 j = 0; j < col_count.index(); ++j)
                {
                    Coordinate coord(row_t(range.getStart().getRow().index() + static_cast<u32>(i)),
                                     column_t(range.getStart().getCol().index() + static_cast<u32>(j)));
                    setCellValue(coord, values[i][j]);
                }
            }

            last_error_.clear();
            return true;
        }

        std::vector<std::vector<TXSheet::CellValue>> getRangeValues(const TXSheet::Range& range) const
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


        const std::string& getLastError() const
        {
            return last_error_;
        }

        // ==================== 合并单元格功能实现 ====================

        bool mergeCells(const TXSheet::Range& range)
        {
            if (!range.isValid())
            {
                last_error_ = "Invalid range for merging";
                return false;
            }

            // 检查范围是否已经有合并的单元格
            auto overlappingRegions = merged_cells_.getOverlappingRegions(range);
            if (!overlappingRegions.empty())
            {
                last_error_ = "Range overlaps with existing merged cells";
                return false;
            }

            // 添加合并区域
            if (!merged_cells_.mergeCells(range))
            {
                last_error_ = "Failed to add merge region";
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

            last_error_.clear();
            return true;
        }

        bool unmergeCells(row_t row, column_t col)
        {
            const auto* region = merged_cells_.getMergeRegion(row, col);
            if (!region)
            {
                last_error_ = "Cell is not in a merged region";
                return false;
            }

            // 移除合并区域
            if (!merged_cells_.unmergeCells(*region))
            {
                last_error_ = "Failed to remove merge region";
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

            last_error_.clear();
            return true;
        }

        std::size_t unmergeCellsInRange(const TXSheet::Range& range)
        {
            return merged_cells_.unmergeCellsInRange(range);
        }

        bool isCellMerged(row_t row, column_t col) const
        {
            return merged_cells_.isMerged(row, col);
        }

        TXSheet::Range getMergeRegion(row_t row, column_t col) const
        {
            const auto* region = merged_cells_.getMergeRegion(row, col);
            if (region)
            {
                TXCoordinate start(region->startRow, region->startCol);
                TXCoordinate end(region->endRow, region->endCol);
                return TXSheet::Range(start, end);
            }
            return TXSheet::Range(); // 返回无效范围
        }

        std::vector<TXSheet::Range> getAllMergeRegions() const
        {
            auto regions = merged_cells_.getAllMergeRegions();
            std::vector<TXSheet::Range> result;
            result.reserve(regions.size());
            for (const auto& region : regions)
            {
                result.emplace_back(TXCoordinate(region.startRow, region.startCol),
                                    TXCoordinate(region.endRow, region.endCol));
            }
            return result;
        }

        std::size_t getMergeCount() const
        {
            return merged_cells_.getMergeCount();
        }

        // ==================== 公式功能实现 ====================

        std::size_t calculateAllFormulas()
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

        std::size_t calculateFormulasInRange(const TXSheet::Range& range)
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

    private:
        std::string name_;
        std::unordered_map<Coordinate, TXCell, CoordinateHash> cells_;
        std::string last_error_;
        TXMergedCells merged_cells_;
        TXWorkbook* workbook_ = nullptr;
    };

    // TXSheet 实现
    TXSheet::TXSheet(const std::string& name, TXWorkbook* parentWorkbook) : pImpl(
        std::make_unique<Impl>(name, parentWorkbook))
    {
    }

    TXSheet::~TXSheet() = default;

    TXSheet::TXSheet(TXSheet&& other) noexcept : pImpl(std::move(other.pImpl))
    {
    }

    TXSheet& TXSheet::operator=(TXSheet&& other) noexcept
    {
        if (this != &other)
        {
            pImpl = std::move(other.pImpl);
        }
        return *this;
    }

    const std::string& TXSheet::getName() const
    {
        return pImpl->getName();
    }

    void TXSheet::setName(const std::string& name)
    {
        pImpl->setName(name);
    }

    TXSheet::CellValue TXSheet::getCellValue(row_t row, column_t col) const
    {
        return pImpl->getCellValue(Coordinate(row, col));
    }

    TXSheet::CellValue TXSheet::getCellValue(const Coordinate& coord) const
    {
        return pImpl->getCellValue(coord);
    }

    TXSheet::CellValue TXSheet::getCellValue(const std::string& address) const
    {
        return pImpl->getCellValue(Coordinate::fromAddress(address));
    }

    bool TXSheet::setCellValue(row_t row, column_t col, const CellValue& value)
    {
        return pImpl->setCellValue(Coordinate(row, col), value);
    }

    bool TXSheet::setCellValue(const Coordinate& coord, const CellValue& value)
    {
        return pImpl->setCellValue(coord, value);
    }

    bool TXSheet::setCellValue(const std::string& address, const CellValue& value)
    {
        return pImpl->setCellValue(Coordinate::fromAddress(address), value);
    }

    TXCell* TXSheet::getCell(row_t row, column_t col)
    {
        return pImpl->getCell(Coordinate(row, col));
    }

    const TXCell* TXSheet::getCell(row_t row, column_t col) const
    {
        return pImpl->getCell(Coordinate(row, col));
    }

    TXCell* TXSheet::getCell(const Coordinate& coord)
    {
        return pImpl->getCell(coord);
    }

    const TXCell* TXSheet::getCell(const Coordinate& coord) const
    {
        return pImpl->getCell(coord);
    }

    TXCell* TXSheet::getCell(const std::string& address)
    {
        return pImpl->getCell(Coordinate::fromAddress(address));
    }

    const TXCell* TXSheet::getCell(const std::string& address) const
    {
        return pImpl->getCell(Coordinate::fromAddress(address));
    }

    bool TXSheet::insertRows(row_t row, row_t count)
    {
        return pImpl->insertRows(row, count);
    }

    bool TXSheet::deleteRows(row_t row, row_t count)
    {
        return pImpl->deleteRows(row, count);
    }

    bool TXSheet::insertColumns(column_t col, column_t count)
    {
        return pImpl->insertColumns(col, count);
    }

    bool TXSheet::deleteColumns(column_t col, column_t count)
    {
        return pImpl->deleteColumns(col, count);
    }

    row_t TXSheet::getUsedRowCount() const
    {
        return pImpl->getUsedRowCount();
    }

    column_t TXSheet::getUsedColumnCount() const
    {
        return pImpl->getUsedColumnCount();
    }

    TXSheet::Range TXSheet::getUsedRange() const
    {
        return pImpl->getUsedRange();
    }

    void TXSheet::clear()
    {
        pImpl->clear();
    }

    std::size_t TXSheet::setCellValues(const std::vector<std::pair<Coordinate, CellValue>>& values)
    {
        return pImpl->setCellValues(values);
    }

    std::vector<std::pair<TXSheet::Coordinate, TXSheet::CellValue>>
    TXSheet::getCellValues(const std::vector<Coordinate>& coords) const
    {
        return pImpl->getCellValues(coords);
    }

    bool TXSheet::setRangeValues(const Range& range, const std::vector<std::vector<CellValue>>& values)
    {
        return pImpl->setRangeValues(range, values);
    }

    std::vector<std::vector<TXSheet::CellValue>> TXSheet::getRangeValues(const Range& range) const
    {
        return pImpl->getRangeValues(range);
    }

    const std::string& TXSheet::getLastError() const
    {
        return pImpl->getLastError();
    }

    // ==================== 合并单元格功能公共接口 ====================

    bool TXSheet::mergeCells(row_t startRow, column_t startCol,
                             row_t endRow, column_t endCol)
    {
        Range range(Coordinate(startRow, startCol), Coordinate(endRow, endCol));
        return pImpl->mergeCells(range);
    }

    bool TXSheet::mergeCells(const Range& range)
    {
        return pImpl->mergeCells(range);
    }

    bool TXSheet::mergeCells(const std::string& rangeStr)
    {
        try
        {
            Range range = addressToRange(rangeStr);
            return pImpl->mergeCells(range);
        }
        catch (const std::exception& e)
        {
            return false;
        }
    }

    bool TXSheet::unmergeCells(row_t row, column_t col)
    {
        return pImpl->unmergeCells(row, col);
    }

    std::size_t TXSheet::unmergeCellsInRange(const Range& range)
    {
        return pImpl->unmergeCellsInRange(range);
    }

    bool TXSheet::isCellMerged(row_t row, column_t col) const
    {
        return pImpl->isCellMerged(row, col);
    }

    TXSheet::Range TXSheet::getMergeRegion(row_t row, column_t col) const
    {
        return pImpl->getMergeRegion(row, col);
    }

    std::vector<TXSheet::Range> TXSheet::getAllMergeRegions() const
    {
        return pImpl->getAllMergeRegions();
    }

    std::size_t TXSheet::getMergeCount() const
    {
        return pImpl->getMergeCount();
    }

    // ==================== 公式功能公共接口 ====================

    std::size_t TXSheet::calculateAllFormulas()
    {
        return pImpl->calculateAllFormulas();
    }

    std::size_t TXSheet::calculateFormulasInRange(const Range& range)
    {
        return pImpl->calculateFormulasInRange(range);
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

    // ==================== 数字格式化功能公共接口 ====================

    bool TXSheet::setCellNumberFormat(row_t row, column_t col,
                                      TXCell::NumberFormat formatType, int decimalPlaces)
    {
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

    std::size_t TXSheet::setRangeNumberFormat(const Range& range, TXCell::NumberFormat formatType,
                                              int decimalPlaces)
    {
        std::size_t count = 0;
        auto start = range.getStart();
        auto end = range.getEnd();

        for (row_t row = start.getRow(); row <= end.getRow(); ++row)
        {
            for (column_t col = start.getCol(); col <= end.getCol(); ++col)
            {
                if (setCellNumberFormat(row, col, formatType, decimalPlaces))
                {
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

    std::size_t TXSheet::setCellFormats(const std::vector<std::pair<Coordinate, TXCell::NumberFormat>>& formats)
    {
        std::size_t count = 0;
        for (const auto& pair : formats)
        {
            if (setCellNumberFormat(pair.first.getRow(), pair.first.getCol(), pair.second))
            {
                count++;
            }
        }
        return count;
    }

    // ==================== 样式功能公共接口 ====================

    bool TXSheet::setCellStyle(row_t row, column_t col, const TXCellStyle& style)
    {
        if (!pImpl) return false;
        TXWorkbook* workbook = pImpl->getWorkbook();
        if (!workbook) return false;

        workbook->getContext()->registerComponentFast(ExcelComponent::Styles);

        u32 style_id = workbook->registerOrGetStyleFId(style);

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
} // namespace TinaXlsx 
