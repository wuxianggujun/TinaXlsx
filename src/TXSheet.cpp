#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXMergedCells.hpp"
#include "TinaXlsx/TXStyle.hpp"
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXWorkbookContext.hpp"
#include "TinaXlsx/TXNumberFormat.hpp"

namespace TinaXlsx {

// ==================== 构造和析构 ====================

TXSheet::TXSheet(const std::string& name, TXWorkbook* parentWorkbook)
    : name_(name), workbook_(parentWorkbook) {
    clearError();
}

TXSheet::~TXSheet() = default;

TXSheet::TXSheet(TXSheet&& other) noexcept
    : name_(std::move(other.name_))
    , workbook_(other.workbook_)
    , lastError_(std::move(other.lastError_))
    , cellManager_(std::move(other.cellManager_))
    , rowColumnManager_(std::move(other.rowColumnManager_))
    , protectionManager_(std::move(other.protectionManager_))
    , formulaManager_(std::move(other.formulaManager_))
    , mergedCells_(std::move(other.mergedCells_)) {
    other.workbook_ = nullptr;
}

TXSheet& TXSheet::operator=(TXSheet&& other) noexcept {
    if (this != &other) {
        name_ = std::move(other.name_);
        workbook_ = other.workbook_;
        lastError_ = std::move(other.lastError_);
        cellManager_ = std::move(other.cellManager_);
        rowColumnManager_ = std::move(other.rowColumnManager_);
        protectionManager_ = std::move(other.protectionManager_);
        formulaManager_ = std::move(other.formulaManager_);
        mergedCells_ = std::move(other.mergedCells_);
        other.workbook_ = nullptr;
    }
    return *this;
}

// ==================== 基本属�?====================

const std::string& TXSheet::getName() const {
    return name_;
}

void TXSheet::setName(const std::string& name) {
    name_ = name;
}

const std::string& TXSheet::getLastError() const {
    return lastError_;
}

// ==================== 单元格操作（委托给CellManager�?===================

TXSheet::CellValue TXSheet::getCellValue(row_t row, column_t col) const {
    return cellManager_.getCellValue(TXCoordinate(row, col));
}

TXSheet::CellValue TXSheet::getCellValue(const Coordinate& coord) const {
    return cellManager_.getCellValue(coord);
}

TXSheet::CellValue TXSheet::getCellValue(const std::string& address) const {
    return cellManager_.getCellValue(Coordinate::fromAddress(address));
}

bool TXSheet::setCellValue(row_t row, column_t col, const CellValue& value) {
    bool result = cellManager_.setCellValue(TXCoordinate(row, col), value);
    if (result) {
        clearError();
        notifyComponentChange(ExcelComponent::BasicWorkbook);
    } else {
        setError("Failed to set cell value");
    }
    return result;
}

bool TXSheet::setCellValue(const Coordinate& coord, const CellValue& value) {
    bool result = cellManager_.setCellValue(coord, value);
    if (result) {
        clearError();
        notifyComponentChange(ExcelComponent::BasicWorkbook);
    } else {
        setError("Failed to set cell value");
    }
    return result;
}

bool TXSheet::setCellValue(const std::string& address, const CellValue& value) {
    return setCellValue(Coordinate::fromAddress(address), value);
}

TXCell* TXSheet::getCell(row_t row, column_t col) {
    return cellManager_.getCell(TXCoordinate(row, col));
}

const TXCell* TXSheet::getCell(row_t row, column_t col) const {
    return cellManager_.getCell(TXCoordinate(row, col));
}

TXCell* TXSheet::getCell(const Coordinate& coord) {
    return cellManager_.getCell(coord);
}

const TXCell* TXSheet::getCell(const Coordinate& coord) const {
    return cellManager_.getCell(coord);
}

TXCell* TXSheet::getCell(const std::string& address) {
    return cellManager_.getCell(Coordinate::fromAddress(address));
}

const TXCell* TXSheet::getCell(const std::string& address) const {
    return cellManager_.getCell(Coordinate::fromAddress(address));
}

TXCell* TXSheet::getCellInternal(const Coordinate& coord) {
    return cellManager_.getCell(coord);
}

const TXCell* TXSheet::getCellInternal(const Coordinate& coord) const {
    return cellManager_.getCell(coord);
}

// ==================== 行列操作（委托给RowColumnManager�?===================

bool TXSheet::insertRows(row_t row, row_t count) {
    if (!protectionManager_.isOperationAllowed(TXSheetProtectionManager::OperationType::InsertRows)) {
        setError("Operation blocked by sheet protection");
        return false;
    }

    bool result = rowColumnManager_.insertRows(row, count, cellManager_);
    if (result) {
        clearError();
        mergedCells_.adjustForRowInsertion(row, count);
        notifyComponentChange(ExcelComponent::BasicWorkbook);
    } else {
        setError("Failed to insert rows");
    }
    return result;
}

bool TXSheet::deleteRows(row_t row, row_t count) {
    if (!protectionManager_.isOperationAllowed(TXSheetProtectionManager::OperationType::DeleteRows)) {
        setError("Operation blocked by sheet protection");
        return false;
    }

    bool result = rowColumnManager_.deleteRows(row, count, cellManager_);
    if (result) {
        clearError();
        mergedCells_.adjustForRowDeletion(row, count);
        notifyComponentChange(ExcelComponent::BasicWorkbook);
    } else {
        setError("Failed to delete rows");
    }
    return result;
}

bool TXSheet::insertColumns(column_t col, column_t count) {
    if (!protectionManager_.isOperationAllowed(TXSheetProtectionManager::OperationType::InsertColumns)) {
        setError("Operation blocked by sheet protection");
        return false;
    }

    bool result = rowColumnManager_.insertColumns(col, count, cellManager_);
    if (result) {
        clearError();
        mergedCells_.adjustForColumnInsertion(col, count);
        notifyComponentChange(ExcelComponent::BasicWorkbook);
    } else {
        setError("Failed to insert columns");
    }
    return result;
}

bool TXSheet::deleteColumns(column_t col, column_t count) {
    if (!protectionManager_.isOperationAllowed(TXSheetProtectionManager::OperationType::DeleteColumns)) {
        setError("Operation blocked by sheet protection");
        return false;
    }

    bool result = rowColumnManager_.deleteColumns(col, count, cellManager_);
    if (result) {
        clearError();
        mergedCells_.adjustForColumnDeletion(col, count);
        notifyComponentChange(ExcelComponent::BasicWorkbook);
    } else {
        setError("Failed to delete columns");
    }
    return result;
}

bool TXSheet::setColumnWidth(column_t col, double width) {
    if (!protectionManager_.isOperationAllowed(TXSheetProtectionManager::OperationType::FormatColumns)) {
        setError("Operation blocked by sheet protection");
        return false;
    }

    bool result = rowColumnManager_.setColumnWidth(col, width);
    if (result) {
        clearError();
    } else {
        setError("Failed to set column width");
    }
    return result;
}

double TXSheet::getColumnWidth(column_t col) const {
    return rowColumnManager_.getColumnWidth(col);
}

bool TXSheet::setRowHeight(row_t row, double height) {
    if (!protectionManager_.isOperationAllowed(TXSheetProtectionManager::OperationType::FormatRows)) {
        setError("Operation blocked by sheet protection");
        return false;
    }

    bool result = rowColumnManager_.setRowHeight(row, height);
    if (result) {
        clearError();
    } else {
        setError("Failed to set row height");
    }
    return result;
}

double TXSheet::getRowHeight(row_t row) const {
    return rowColumnManager_.getRowHeight(row);
}

double TXSheet::autoFitColumnWidth(column_t col, double minWidth, double maxWidth) {
    return rowColumnManager_.autoFitColumnWidth(col, cellManager_, minWidth, maxWidth);
}

double TXSheet::autoFitRowHeight(row_t row, double minHeight, double maxHeight) {
    return rowColumnManager_.autoFitRowHeight(row, cellManager_, minHeight, maxHeight);
}

std::size_t TXSheet::autoFitAllColumnWidths(double minWidth, double maxWidth) {
    return rowColumnManager_.autoFitAllColumnWidths(cellManager_, minWidth, maxWidth);
}

std::size_t TXSheet::autoFitAllRowHeights(double minHeight, double maxHeight) {
    return rowColumnManager_.autoFitAllRowHeights(cellManager_, minHeight, maxHeight);
}

// ==================== 范围信息（委托给CellManager�?===================

row_t TXSheet::getUsedRowCount() const {
    return cellManager_.getMaxUsedRow();
}

column_t TXSheet::getUsedColumnCount() const {
    return cellManager_.getMaxUsedColumn();
}

TXSheet::Range TXSheet::getUsedRange() const {
    return cellManager_.getUsedRange();
}

// ==================== 批量操作（委托给CellManager�?===================

std::size_t TXSheet::setCellValues(const std::vector<std::pair<Coordinate, CellValue>>& values) {
    std::size_t count = cellManager_.setCellValues(values);
    if (count > 0) {
        notifyComponentChange(ExcelComponent::BasicWorkbook);
    }
    return count;
}

std::vector<std::pair<TXSheet::Coordinate, TXSheet::CellValue>> 
TXSheet::getCellValues(const std::vector<Coordinate>& coords) const {
    return cellManager_.getCellValues(coords);
}

// ==================== 工作表保护（委托给ProtectionManager�?===================

bool TXSheet::protectSheet(const std::string& password, const SheetProtection& protection) {
    bool result = protectionManager_.protectSheet(password, protection);
    if (result) {
        clearError();
        notifyComponentChange(ExcelComponent::BasicWorkbook);
    } else {
        setError("Failed to protect sheet");
    }
    return result;
}

bool TXSheet::unprotectSheet(const std::string& password) {
    bool result = protectionManager_.unprotectSheet(password);
    if (result) {
        clearError();
        notifyComponentChange(ExcelComponent::BasicWorkbook);
    } else {
        setError("Failed to unprotect sheet");
    }
    return result;
}

bool TXSheet::isSheetProtected() const {
    return protectionManager_.isSheetProtected();
}

const TXSheet::SheetProtection& TXSheet::getSheetProtection() const {
    return protectionManager_.getSheetProtection();
}

bool TXSheet::setCellLocked(row_t row, column_t col, bool locked) {
    return protectionManager_.setCellLocked(TXCoordinate(row, col), locked, cellManager_);
}

bool TXSheet::isCellLocked(row_t row, column_t col) const {
    return protectionManager_.isCellLocked(TXCoordinate(row, col), cellManager_);
}

std::size_t TXSheet::setRangeLocked(const Range& range, bool locked) {
    return protectionManager_.setRangeLocked(range, locked, cellManager_);
}

// ==================== 公式操作（委托给FormulaManager�?===================

std::size_t TXSheet::calculateAllFormulas() {
    return formulaManager_.calculateAllFormulas(cellManager_);
}

std::size_t TXSheet::calculateFormulasInRange(const Range& range) {
    return formulaManager_.calculateFormulasInRange(range, cellManager_);
}

bool TXSheet::setCellFormula(row_t row, column_t col, const std::string& formula) {
    bool result = formulaManager_.setCellFormula(TXCoordinate(row, col), formula, cellManager_);
    if (result) {
        clearError();
        notifyComponentChange(ExcelComponent::BasicWorkbook);
    } else {
        setError("Failed to set cell formula");
    }
    return result;
}

std::string TXSheet::getCellFormula(row_t row, column_t col) const {
    return formulaManager_.getCellFormula(TXCoordinate(row, col), cellManager_);
}

std::size_t TXSheet::setCellFormulas(const std::vector<std::pair<Coordinate, std::string>>& formulas) {
    std::size_t count = 0;
    for (const auto& pair : formulas) {
        if (setCellFormula(pair.first.getRow(), pair.first.getCol(), pair.second)) {
            ++count;
        }
    }
    return count;
}

void TXSheet::setFormulaCalculationOptions(const FormulaCalculationOptions& options) {
    formulaManager_.setCalculationOptions(options);
}

const TXSheet::FormulaCalculationOptions& TXSheet::getFormulaCalculationOptions() const {
    return formulaManager_.getCalculationOptions();
}

bool TXSheet::addNamedRange(const std::string& name, const Range& range, const std::string& comment) {
    return formulaManager_.addNamedRange(name, range, comment);
}

bool TXSheet::removeNamedRange(const std::string& name) {
    return formulaManager_.removeNamedRange(name);
}

TXSheet::Range TXSheet::getNamedRange(const std::string& name) const {
    return formulaManager_.getNamedRange(name);
}

std::unordered_map<std::string, TXSheet::Range> TXSheet::getAllNamedRanges() const {
    return formulaManager_.getAllNamedRanges();
}

bool TXSheet::detectCircularReferences() const {
    return formulaManager_.detectCircularReferences(cellManager_);
}

std::unordered_map<TXSheet::Coordinate, std::vector<TXSheet::Coordinate>, TXSheet::CoordinateHash>
TXSheet::getFormulaDependencies() const {
    auto deps = formulaManager_.getFormulaDependencies(cellManager_);
    std::unordered_map<Coordinate, std::vector<Coordinate>, CoordinateHash> result;
    for (const auto& pair : deps) {
        result[pair.first] = pair.second;
    }
    return result;
}

// ==================== 格式化操�?====================

bool TXSheet::setCellNumberFormat(row_t row, column_t col, TXNumberFormat::FormatType formatType, int decimalPlaces) {
    TXCell* cell = getCell(row, col);
    if (!cell || !workbook_) {
        return false;
    }

    // 确保样式组件已注�?    notifyComponentChange(ExcelComponent::Styles);

    // 获取当前有效样式
    TXCellStyle styleToApply = getCellEffectiveStyle(cell);

    // 设置数字格式
    bool useThousandSeparator = (formatType == TXNumberFormat::FormatType::Number ||
                               formatType == TXNumberFormat::FormatType::Currency);
    styleToApply.setNumberFormat(formatType, decimalPlaces, useThousandSeparator);

    // 应用完整样式
    return setCellStyle(row, col, styleToApply);
}

bool TXSheet::setCellCustomFormat(row_t row, column_t col, const std::string& formatString) {
    TXCell* cell = getCell(row, col);
    if (!cell || !workbook_) {
        return false;
    }

    // 确保样式组件已注册
    notifyComponentChange(ExcelComponent::Styles);

    // 获取当前有效样式
    TXCellStyle styleToApply = getCellEffectiveStyle(cell);

    // 设置自定义数字格式
    styleToApply.setCustomNumberFormat(formatString);

    // 应用完整样式
    return setCellStyle(row, col, styleToApply);
}

std::size_t TXSheet::setRangeNumberFormat(const Range& range, TXNumberFormat::FormatType formatType, int decimalPlaces) {
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

std::string TXSheet::getCellFormattedValue(row_t row, column_t col) const {
    const TXCell* cell = getCell(row, col);
    if (!cell) {
        return "";
    }
    return cell->getFormattedValue();
}

std::size_t TXSheet::setCellFormats(const std::vector<std::pair<Coordinate, TXNumberFormat::FormatType>>& formats) {
    std::size_t count = 0;
    for (const auto& pair : formats) {
        if (setCellNumberFormat(pair.first.getRow(), pair.first.getCol(), pair.second)) {
            count++;
        }
    }
    return count;
}

// ==================== 样式操作 ====================

bool TXSheet::setCellStyle(row_t row, column_t col, const TXCellStyle& style) {
    if (!workbook_) {
        setError("No workbook associated");
        return false;
    }

    if (!protectionManager_.isOperationAllowed(TXSheetProtectionManager::OperationType::FormatCells)) {
        setError("Operation blocked by sheet protection");
        return false;
    }

    notifyComponentChange(ExcelComponent::Styles);

    // 使用样式管理器注册样式
    auto& styleManager = workbook_->getStyleManager();
    u32 styleId = styleManager.registerCellStyleXF(style);

    TXCell* cell = getCell(row, col);
    if (!cell) {
        setError("Failed to get cell");
        return false;
    }

    // 设置样式索引
    cell->setStyleIndex(styleId);

    // 同步数字格式对象到单元格
    auto numberFormatObject = style.createNumberFormatObject();
    if (numberFormatObject) {
        cell->setNumberFormatObject(std::move(numberFormatObject));
    }

    clearError();
    return true;
}

bool TXSheet::setCellStyle(const std::string& address, const TXCellStyle& style) {
    auto coord = Coordinate::fromAddress(address);
    return setCellStyle(coord.getRow(), coord.getCol(), style);
}

std::size_t TXSheet::setRangeStyle(const Range& range, const TXCellStyle& style) {
    std::size_t count = 0;
    auto start = range.getStart();
    auto end = range.getEnd();

    for (row_t row = start.getRow(); row <= end.getRow(); ++row) {
        for (column_t col = start.getCol(); col <= end.getCol(); ++col) {
            if (setCellStyle(row, col, style)) {
                count++;
            }
        }
    }
    return count;
}

std::size_t TXSheet::setCellStyles(const std::vector<std::pair<Coordinate, TXCellStyle>>& styles) {
    std::size_t count = 0;
    for (const auto& pair : styles) {
        if (setCellStyle(pair.first.getRow(), pair.first.getCol(), pair.second)) {
            count++;
        }
    }
    return count;
}

std::size_t TXSheet::setBatchNumberFormats(const std::vector<std::pair<Coordinate, TXCellStyle::NumberFormatDefinition>>& formats) {
    if (!workbook_) return 0;

    // 确保样式组件已注册
    notifyComponentChange(ExcelComponent::Styles);

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

// ==================== 私有辅助方法 ====================

bool TXSheet::applyCellNumberFormat(TXCell* cell, u32 numFmtId) {
    if (!cell || !workbook_) {
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

TXCellStyle TXSheet::getCellEffectiveStyle(TXCell* cell) {
    if (!cell || !workbook_) {
        return TXCellStyle(); // 默认样式
    }

    // 从样式管理器获取完整样式对象
    u32 styleIndex = cell->getStyleIndex();
    if (styleIndex == 0) {
        // 使用默认样式
        return TXCellStyle();
    }

    // 从TXStyleManager 获取完整样式对象
    return workbook_->getStyleManager().getStyleObjectFromXfIndex(styleIndex);
}

void TXSheet::updateUsedRange() {
    // 内部辅助方法，用于优化已使用范围的计算
    // 这里可以添加缓存逻辑
}

std::vector<TXSheet::Coordinate> TXSheet::parseFormulaReferences(const std::string& formula) const {
    return formulaManager_.parseFormulaReferences(formula);
}



// ==================== 清空操作 ====================

void TXSheet::clear() {
    cellManager_.clear();
    rowColumnManager_.clear();
    protectionManager_.clear();
    formulaManager_.clear();
    mergedCells_.clear();
    clearError();
}

void TXSheet::notifyComponentChange(ExcelComponent component) const {
    if (workbook_ && workbook_->getContext()) {
        workbook_->getContext()->registerComponentFast(component);
    }
}

// ==================== 范围操作方法 ====================

bool TXSheet::setRangeValues(const Range& range, const std::vector<std::vector<CellValue>>& values) {
    if (values.empty()) {
        setError("Empty values array");
        return false;
    }

    auto start = range.getStart();
    auto end = range.getEnd();

    std::size_t rowCount = values.size();
    std::size_t colCount = values[0].size();

    // 检查范围是否匹配
    if (static_cast<std::size_t>(end.getRow().index() - start.getRow().index() + 1) != rowCount ||
        static_cast<std::size_t>(end.getCol().index() - start.getCol().index() + 1) != colCount) {
        setError("Range size does not match values array size");
        return false;
    }

    // 设置值
    for (std::size_t i = 0; i < rowCount; ++i) {
        if (values[i].size() != colCount) {
            setError("Inconsistent row sizes in values array");
            return false;
        }

        for (std::size_t j = 0; j < colCount; ++j) {
            row_t row = row_t(start.getRow().index() + i);
            column_t col = column_t(start.getCol().index() + j);

            if (!setCellValue(row, col, values[i][j])) {
                return false;
            }
        }
    }

    clearError();
    return true;
}

std::vector<std::vector<TXSheet::CellValue>> TXSheet::getRangeValues(const Range& range) const {
    std::vector<std::vector<CellValue>> result;

    auto start = range.getStart();
    auto end = range.getEnd();

    for (row_t row = start.getRow(); row <= end.getRow(); ++row) {
        std::vector<CellValue> rowValues;

        for (column_t col = start.getCol(); col <= end.getCol(); ++col) {
            rowValues.push_back(getCellValue(row, col));
        }

        result.push_back(std::move(rowValues));
    }

    return result;
}

// ==================== 合并单元格方法 ====================

bool TXSheet::mergeCells(const Range& range) {
    return mergedCells_.mergeCells(range);
}

bool TXSheet::mergeCells(row_t startRow, column_t startCol, row_t endRow, column_t endCol) {
    return mergedCells_.mergeCells(startRow, startCol, endRow, endCol);
}

bool TXSheet::unmergeCells(row_t row, column_t col) {
    return mergedCells_.unmergeCells(row, col);
}

bool TXSheet::isCellMerged(row_t row, column_t col) const {
    return mergedCells_.isMerged(row, col);
}

TXRange TXSheet::getMergeRegion(row_t row, column_t col) const {
    const auto* region = mergedCells_.getMergeRegion(row, col);
    if (region) {
        return region->toRange();
    }
    return TXRange(); // 返回无效范围
}

std::vector<TXRange> TXSheet::getAllMergeRegions() const {
    auto regions = mergedCells_.getAllMergeRegions();
    std::vector<TXRange> result;
    result.reserve(regions.size());

    for (const auto& region : regions) {
        result.push_back(region.toRange());
    }

    return result;
}

std::size_t TXSheet::getMergeCount() const {
    return mergedCells_.getMergeCount();
}

} // namespace TinaXlsx
