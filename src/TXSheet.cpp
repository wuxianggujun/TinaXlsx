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
#include <set>
#include <unordered_set>
#include <cmath>
#include <functional>

namespace TinaXlsx
{
    // ==================== TXSheet 实现 ====================

    TXSheet::TXSheet(const std::string& name, TXWorkbook* parentWorkbook)
        : name_(name), workbook_(parentWorkbook)
    {
        // 初始化公式计算选项为默认值
        formulaOptions_.autoCalculate = true;
        formulaOptions_.iterativeCalculation = false;
        formulaOptions_.maxIterations = 100;
        formulaOptions_.maxChange = 0.001;
        formulaOptions_.precisionAsDisplayed = false;
        formulaOptions_.use1904DateSystem = false;
    }

    TXSheet::~TXSheet() = default;

    TXSheet::TXSheet(TXSheet&& other) noexcept
        : name_(std::move(other.name_))
          , cells_(std::move(other.cells_))
          , lastError_(std::move(other.lastError_))
          , mergedCells_(std::move(other.mergedCells_))
          , workbook_(other.workbook_)
          , columnWidths_(std::move(other.columnWidths_))
          , rowHeights_(std::move(other.rowHeights_))
          , sheetProtection_(std::move(other.sheetProtection_))
          , formulaOptions_(std::move(other.formulaOptions_))
          , namedRanges_(std::move(other.namedRanges_))
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
            columnWidths_ = std::move(other.columnWidths_);
            rowHeights_ = std::move(other.rowHeights_);
            sheetProtection_ = std::move(other.sheetProtection_);
            formulaOptions_ = std::move(other.formulaOptions_);
            namedRanges_ = std::move(other.namedRanges_);
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
        TXCell* cell = getCell(row, col);
        if (!cell || !workbook_)
        {
            return false;
        }

        // 确保样式组件已注册
        if (auto* context = workbook_->getContext())
        {
            context->registerComponentFast(ExcelComponent::Styles);
        }

        // 新架构：通过 TXCellStyle 统一处理
        // 1. 获取当前有效样式
        TXCellStyle style_to_apply = getCellEffectiveStyle(cell);
        
        // 2. 设置数字格式
        bool useThousandSeparator = (formatType == TXNumberFormat::FormatType::Number ||
                                   formatType == TXNumberFormat::FormatType::Currency);
        style_to_apply.setNumberFormat(formatType, decimalPlaces, useThousandSeparator);
        
        // 3. 应用完整样式
        return setCellStyle(row, col, style_to_apply);
    }


    bool TXSheet::setCellCustomFormat(row_t row, column_t col,
                                      const std::string& formatString)
    {
        TXCell* cell = getCell(row, col);
        if (!cell || !workbook_)
        {
            return false;
        }

        // 确保样式组件已注册
        if (auto* context = workbook_->getContext())
        {
            context->registerComponentFast(ExcelComponent::Styles);
        }

        // 新架构：通过 TXCellStyle 统一处理
        // 1. 获取当前有效样式
        TXCellStyle style_to_apply = getCellEffectiveStyle(cell);
        
        // 2. 设置自定义数字格式
        style_to_apply.setCustomNumberFormat(formatString);
        
        // 3. 应用完整样式
        return setCellStyle(row, col, style_to_apply);
    }

    bool TXSheet::applyCellNumberFormat(TXCell* cell, u32 numFmtId)
    {
        if (!cell || !workbook_)
        {
            return false;
        }

        auto& styleManager = workbook_->getStyleManager();
        
        // 获取当前样式
        TXCellStyle cellStyle = getCellEffectiveStyle(cell);
        
        // 创建数字格式定义 - 这里简化处理，因为我们只有numFmtId
        // 实际应用中可能需要根据numFmtId逆向解析格式类型
        TXCellStyle::NumberFormatDefinition numFmtDef;
        if (numFmtId == 0) {
            numFmtDef = TXCellStyle::NumberFormatDefinition(); // 常规格式
        } else {
            // 简化处理：作为自定义格式处理
            numFmtDef = TXCellStyle::NumberFormatDefinition("General"); // 默认
        }
        
        cellStyle.setNumberFormatDefinition(numFmtDef);
        
        // 注册或获取这个样式的XF记录
        u32 styleIndex = styleManager.registerCellStyleXF(cellStyle);

        // 设置单元格的样式索引
        cell->setStyleIndex(styleIndex);

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

        // 使用新的 TXStyleManager 方法
        auto& styleManager = workbook_->getStyleManager();
        u32 style_id = styleManager.registerCellStyleXF(style);

        TXCell* cell = getCell(row, col);
        if (!cell)
        {
            return false;
        }
        
        // 设置样式索引
        cell->setStyleIndex(style_id);
        
        // 重要：同步数字格式对象到单元格
        auto numberFormatObject = style.createNumberFormatObject();
        if (numberFormatObject) {
            cell->setNumberFormatObject(std::move(numberFormatObject));
        }
        
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

    std::size_t TXSheet::setBatchNumberFormats(const std::vector<std::pair<Coordinate, TXCellStyle::NumberFormatDefinition>>& formats) {
        if (!workbook_) return 0;
        
        // 确保样式组件已注册
        workbook_->getContext()->registerComponentFast(ExcelComponent::Styles);
        
        std::size_t count = 0;
        
        for (const auto& pair : formats) {
            const auto& coord = pair.first;
            const auto& formatDef = pair.second;
            
            TXCell* cell = getCell(coord.getRow(), coord.getCol());
            if (!cell) continue;
            
            // 获取当前样式
            TXCellStyle style = getCellEffectiveStyle(cell);
            
            // 设置新的数字格式
            style.setNumberFormatDefinition(formatDef);
            
            // 应用样式
            if (setCellStyle(coord.getRow(), coord.getCol(), style)) {
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

    TXCellStyle TXSheet::getCellEffectiveStyle(TXCell* cell)
    {
        if (!cell || !workbook_) {
            return TXCellStyle(); // 默认样式
        }
        
        // 完整实现：从样式管理器反向构造样式
        u32 styleIndex = cell->getStyleIndex();
        if (styleIndex == 0) {
            // 使用默认样式
            return TXCellStyle();
        }
        
        // 从 TXStyleManager 获取完整样式对象
        return workbook_->getStyleManager().getStyleObjectFromXfIndex(styleIndex);
    }

    // ==================== 列宽和行高操作实现 ====================

    bool TXSheet::setColumnWidth(column_t col, double width)
    {
        if (col.index() <= 0 || width < 0.0 || width > 255.0) {
            lastError_ = "Invalid column or width value";
            return false;
        }
        
        if (isOperationBlocked("formatColumns")) {
            lastError_ = "Operation blocked by sheet protection";
            return false;
        }
        
        columnWidths_[col.index()] = width;
        return true;
    }

    double TXSheet::getColumnWidth(column_t col) const
    {
        auto it = columnWidths_.find(col.index());
        return (it != columnWidths_.end()) ? it->second : 8.43; // Excel默认列宽
    }

    bool TXSheet::setRowHeight(row_t row, double height)
    {
        if (row.index() <= 0 || height < 0.0 || height > 409.0) {
            lastError_ = "Invalid row or height value";
            return false;
        }
        
        if (isOperationBlocked("formatRows")) {
            lastError_ = "Operation blocked by sheet protection";
            return false;
        }
        
        rowHeights_[row.index()] = height;
        return true;
    }

    double TXSheet::getRowHeight(row_t row) const
    {
        auto it = rowHeights_.find(row.index());
        return (it != rowHeights_.end()) ? it->second : 15.0; // Excel默认行高
    }

    double TXSheet::autoFitColumnWidth(column_t col, double minWidth, double maxWidth)
    {
        if (col.index() <= 0) {
            return getColumnWidth(col);
        }
        
        double maxContentWidth = minWidth;
        
        // 遍历该列的所有单元格，计算最大内容宽度
        for (const auto& [coord, cell] : cells_) {
            if (coord.getCol() == col) {
                std::string displayText = cell.getFormattedValue();
                if (!displayText.empty()) {
                    double textWidth = calculateTextWidth(displayText);
                    maxContentWidth = std::max(maxContentWidth, textWidth);
                }
            }
        }
        
        // 限制在最小和最大宽度之间
        double finalWidth = std::min(std::max(maxContentWidth, minWidth), maxWidth);
        setColumnWidth(col, finalWidth);
        
        return finalWidth;
    }

    double TXSheet::autoFitRowHeight(row_t row, double minHeight, double maxHeight)
    {
        if (row.index() <= 0) {
            return getRowHeight(row);
        }
        
        double maxContentHeight = minHeight;
        
        // 遍历该行的所有单元格，计算最大内容高度
        for (const auto& [coord, cell] : cells_) {
            if (coord.getRow() == row) {
                std::string displayText = cell.getFormattedValue();
                if (!displayText.empty()) {
                    double columnWidth = getColumnWidth(coord.getCol());
                    double textHeight = calculateTextHeight(displayText, 11.0, columnWidth);
                    maxContentHeight = std::max(maxContentHeight, textHeight);
                }
            }
        }
        
        // 限制在最小和最大高度之间
        double finalHeight = std::min(std::max(maxContentHeight, minHeight), maxHeight);
        setRowHeight(row, finalHeight);
        
        return finalHeight;
    }

    std::size_t TXSheet::autoFitAllColumnWidths(double minWidth, double maxWidth)
    {
        std::set<column_t> usedColumns;
        
        // 收集所有使用的列
        for (const auto& [coord, cell] : cells_) {
            usedColumns.insert(coord.getCol());
        }
        
        std::size_t count = 0;
        for (column_t col : usedColumns) {
            autoFitColumnWidth(col, minWidth, maxWidth);
            ++count;
        }
        
        return count;
    }

    std::size_t TXSheet::autoFitAllRowHeights(double minHeight, double maxHeight)
    {
        std::set<row_t> usedRows;
        
        // 收集所有使用的行
        for (const auto& [coord, cell] : cells_) {
            usedRows.insert(coord.getRow());
        }
        
        std::size_t count = 0;
        for (row_t row : usedRows) {
            autoFitRowHeight(row, minHeight, maxHeight);
            ++count;
        }
        
        return count;
    }

    // ==================== 工作表保护功能实现 ====================

    bool TXSheet::protectSheet(const std::string& password, const SheetProtection& protection)
    {
        sheetProtection_ = protection;
        sheetProtection_.isProtected = true;
        
        if (!password.empty()) {
            sheetProtection_.password = generatePasswordHash(password);
        }
        
        return true;
    }

    bool TXSheet::unprotectSheet(const std::string& password)
    {
        if (!sheetProtection_.isProtected) {
            return true; // 已经未保护
        }
        
        if (!sheetProtection_.password.empty()) {
            if (!verifyPassword(password, sheetProtection_.password)) {
                lastError_ = "Invalid password";
                return false;
            }
        }
        
        sheetProtection_.isProtected = false;
        sheetProtection_.password.clear();
        
        return true;
    }

    bool TXSheet::isSheetProtected() const
    {
        return sheetProtection_.isProtected;
    }

    const TXSheet::SheetProtection& TXSheet::getSheetProtection() const
    {
        return sheetProtection_;
    }

    bool TXSheet::setCellLocked(row_t row, column_t col, bool locked)
    {
        TXCell* cell = getCell(row, col);
        if (!cell) {
            // 创建新单元格
            setCellValue(row, col, cell_value_t{});
            cell = getCell(row, col);
        }
        
        if (cell) {
            cell->setLocked(locked);
            return true;
        }
        
        return false;
    }

    bool TXSheet::isCellLocked(row_t row, column_t col) const
    {
        const TXCell* cell = getCell(row, col);
        return cell ? cell->isLocked() : true; // 默认锁定
    }

    std::size_t TXSheet::setRangeLocked(const Range& range, bool locked)
    {
        std::size_t count = 0;
        
        for (row_t r(range.getStart().getRow().index()); r.index() <= range.getEnd().getRow().index(); ++r) {
            for (column_t c(range.getStart().getCol().index()); c.index() <= range.getEnd().getCol().index(); ++c) {
                if (setCellLocked(r, c, locked)) {
                    ++count;
                }
            }
        }
        
        return count;
    }

    // ==================== 增强公式支持实现 ====================

    void TXSheet::setFormulaCalculationOptions(const FormulaCalculationOptions& options)
    {
        formulaOptions_ = options;
    }

    const TXSheet::FormulaCalculationOptions& TXSheet::getFormulaCalculationOptions() const
    {
        return formulaOptions_;
    }

    bool TXSheet::addNamedRange(const std::string& name, const Range& range, const std::string& comment)
    {
        if (name.empty() || !range.isValid()) {
            lastError_ = "Invalid name or range";
            return false;
        }
        
        namedRanges_[name] = range;
        return true;
    }

    bool TXSheet::removeNamedRange(const std::string& name)
    {
        auto it = namedRanges_.find(name);
        if (it != namedRanges_.end()) {
            namedRanges_.erase(it);
            return true;
        }
        return false;
    }

    TXSheet::Range TXSheet::getNamedRange(const std::string& name) const
    {
        auto it = namedRanges_.find(name);
        return (it != namedRanges_.end()) ? it->second : Range();
    }

    std::unordered_map<std::string, TXSheet::Range> TXSheet::getAllNamedRanges() const
    {
        return namedRanges_;
    }

    bool TXSheet::detectCircularReferences() const
    {
        // 简化的循环引用检测
        std::unordered_set<Coordinate, CoordinateHash> visiting;
        std::unordered_set<Coordinate, CoordinateHash> visited;
        
        for (const auto& [coord, cell] : cells_) {
            if (cell.hasFormula() && visited.find(coord) == visited.end()) {
                if (detectCircularReferencesHelper(coord, visiting, visited)) {
                    return true;
                }
            }
        }
        
        return false;
    }

    std::unordered_map<TXSheet::Coordinate, std::vector<TXSheet::Coordinate>, TXSheet::CoordinateHash> 
    TXSheet::getFormulaDependencies() const
    {
        std::unordered_map<Coordinate, std::vector<Coordinate>, CoordinateHash> dependencies;
        
        for (const auto& [coord, cell] : cells_) {
            if (cell.hasFormula()) {
                // 解析公式中的单元格引用
                std::vector<Coordinate> refs = parseFormulaReferences(cell.getFormula());
                dependencies[coord] = refs;
            }
        }
        
        return dependencies;
    }

    // ==================== 私有辅助方法实现 ====================

    double TXSheet::calculateTextWidth(const std::string& text, double fontSize, const std::string& fontName) const
    {
        // 简化的文本宽度计算
        // 实际实现应该考虑字体度量
        double charWidth = fontSize * 0.6; // 近似值
        return text.length() * charWidth / 7.0; // 转换为Excel字符单位
    }

    double TXSheet::calculateTextHeight(const std::string& text, double fontSize, double columnWidth) const
    {
        // 简化的文本高度计算
        double lineHeight = fontSize * 1.2; // 行高通常是字体大小的1.2倍
        
        // 计算换行数
        double textWidth = calculateTextWidth(text, fontSize);
        double excelColumnWidth = columnWidth * 7.0; // 转换为像素近似值
        int lines = static_cast<int>(std::ceil(textWidth / excelColumnWidth));
        
        return std::max(1, lines) * lineHeight;
    }

    std::string TXSheet::generatePasswordHash(const std::string& password) const
    {
        // 简化的哈希实现（实际应使用MD5或更安全的算法）
        std::hash<std::string> hasher;
        return std::to_string(hasher(password));
    }

    bool TXSheet::verifyPassword(const std::string& password, const std::string& hash) const
    {
        return generatePasswordHash(password) == hash;
    }

    bool TXSheet::isOperationBlocked(const std::string& operation) const
    {
        if (!sheetProtection_.isProtected) {
            return false;
        }
        
        if (operation == "formatCells") return !sheetProtection_.formatCells;
        if (operation == "formatColumns") return !sheetProtection_.formatColumns;
        if (operation == "formatRows") return !sheetProtection_.formatRows;
        if (operation == "insertColumns") return !sheetProtection_.insertColumns;
        if (operation == "insertRows") return !sheetProtection_.insertRows;
        if (operation == "deleteColumns") return !sheetProtection_.deleteColumns;
        if (operation == "deleteRows") return !sheetProtection_.deleteRows;
        if (operation == "sort") return !sheetProtection_.sort;
        if (operation == "autoFilter") return !sheetProtection_.autoFilter;
        
        return false; // 默认允许
    }

    bool TXSheet::detectCircularReferencesHelper(const Coordinate& coord, 
                                                std::unordered_set<Coordinate, CoordinateHash>& visiting,
                                                std::unordered_set<Coordinate, CoordinateHash>& visited) const
    {
        if (visiting.find(coord) != visiting.end()) {
            return true; // 发现循环
        }
        
        if (visited.find(coord) != visited.end()) {
            return false; // 已经访问过
        }
        
        visiting.insert(coord);
        
        const TXCell* cell = getCell(coord);
        if (cell && cell->hasFormula()) {
            std::vector<Coordinate> refs = parseFormulaReferences(cell->getFormula());
            for (const Coordinate& ref : refs) {
                if (detectCircularReferencesHelper(ref, visiting, visited)) {
                    return true;
                }
            }
        }
        
        visiting.erase(coord);
        visited.insert(coord);
        
        return false;
    }

    std::vector<TXCoordinate> TXSheet::parseFormulaReferences(const std::string& formula) const
    {
        std::vector<Coordinate> references;
        
        // 简化的公式引用解析
        // 实际实现应该使用更复杂的解析器
        std::regex cellRefRegex(R"([A-Z]+[0-9]+)");
        std::sregex_iterator iter(formula.begin(), formula.end(), cellRefRegex);
        std::sregex_iterator end;
        
        for (; iter != end; ++iter) {
            std::string ref = iter->str();
            try {
                Coordinate coord = Coordinate::fromAddress(ref);
                references.push_back(coord);
            } catch (...) {
                // 忽略无效引用
            }
        }
        
        return references;
    }

} // namespace TinaXlsx 
