#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCompactCell.hpp"
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

TXSheet::~TXSheet() {
    // 在析构前先清空workbook_指针，防止在管理器析构时访问无效指针
    workbook_ = nullptr;
}

TXSheet::TXSheet(TXSheet&& other) noexcept
    : name_(std::move(other.name_))
    , workbook_(other.workbook_)
    , lastError_(std::move(other.lastError_))
    , cellManager_(std::move(other.cellManager_))
    , rowColumnManager_(std::move(other.rowColumnManager_))
    , protectionManager_(std::move(other.protectionManager_))
    , formulaManager_(std::move(other.formulaManager_))
    , mergedCells_(std::move(other.mergedCells_))
    , charts_(std::move(other.charts_))
    , nextChartId_(other.nextChartId_) {
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
        charts_ = std::move(other.charts_);
        nextChartId_ = other.nextChartId_;
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

TXCompactCell* TXSheet::getCell(row_t row, column_t col) {
    return cellManager_.getCell(TXCoordinate(row, col));
}

const TXCompactCell* TXSheet::getCell(row_t row, column_t col) const {
    return cellManager_.getCell(TXCoordinate(row, col));
}

TXCompactCell* TXSheet::getCell(const Coordinate& coord) {
    return cellManager_.getCell(coord);
}

const TXCompactCell* TXSheet::getCell(const Coordinate& coord) const {
    return cellManager_.getCell(coord);
}

TXCompactCell* TXSheet::getCell(const std::string& address) {
    return cellManager_.getCell(Coordinate::fromAddress(address));
}

const TXCompactCell* TXSheet::getCell(const std::string& address) const {
    return cellManager_.getCell(Coordinate::fromAddress(address));
}

TXCompactCell* TXSheet::getCellInternal(const Coordinate& coord) {
    return cellManager_.getCell(coord);
}

const TXCompactCell* TXSheet::getCellInternal(const Coordinate& coord) const {
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

// ==================== 高性能批量操作实现 ====================

std::size_t TXSheet::setCellValuesBatch(const std::vector<std::pair<TXCoordinate, CellValue>>& values) {
    // 转换为CellManager需要的格式
    std::vector<std::pair<Coordinate, CellValue>> convertedValues;
    convertedValues.reserve(values.size());

    for (const auto& pair : values) {
        convertedValues.emplace_back(Coordinate(pair.first.getRow(), pair.first.getCol()), pair.second);
    }

    std::size_t count = cellManager_.setCellValues(convertedValues);
    if (count > 0) {
        notifyComponentChange(ExcelComponent::BasicWorkbook);
    }
    return count;
}

std::size_t TXSheet::setRangeValues(row_t startRow, column_t startCol,
                                   const std::vector<std::vector<CellValue>>& values) {
    std::size_t count = cellManager_.setRangeValues(startRow, startCol, values);
    if (count > 0) {
        notifyComponentChange(ExcelComponent::BasicWorkbook);
    }
    return count;
}

std::size_t TXSheet::setRowValues(row_t row, column_t startCol, const std::vector<CellValue>& values) {
    std::size_t count = cellManager_.setRowValues(row, startCol, values);
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
    TXCoordinate coord(row, col);

    // 首先通过保护管理器设置锁定状态
    bool result = protectionManager_.setCellLocked(coord, locked, cellManager_);

    if (result && workbook_) {
        // 然后更新单元格样式以反映锁定状态
        TXCompactCell* cell = cellManager_.getCell(coord);
        if (cell) {
            // 获取当前样式
            auto& styleManager = workbook_->getStyleManager();
            TXCellStyle currentStyle;

            u32 currentStyleIndex = cell->getStyleIndex();
            if (currentStyleIndex > 0) {
                // 如果有现有样式，获取它
                currentStyle = styleManager.getStyleObjectFromXfIndex(currentStyleIndex);
            }

            // 更新锁定状态
            currentStyle.setLocked(locked);

            // 注册新样式并应用到单元格
            u32 newStyleIndex = styleManager.registerCellStyleXF(currentStyle);
            cell->setStyleIndex(newStyleIndex);
        }
    }

    return result;
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
    TXCompactCell* cell = getCell(row, col);
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
    TXCompactCell* cell = getCell(row, col);
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
    const TXCompactCell* cell = getCell(row, col);
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

    TXCompactCell* cell = getCell(row, col);
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

        TXCompactCell* cell = getCell(coord.getRow(), coord.getCol());
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

bool TXSheet::applyCellNumberFormat(TXCompactCell* cell, u32 numFmtId) {
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

TXCellStyle TXSheet::getCellEffectiveStyle(TXCompactCell* cell) {
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
    charts_.clear();
    nextChartId_ = 1;
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
            row_t row = row_t(start.getRow().index() + static_cast<row_t::index_t>(i));
            column_t col = column_t(start.getCol().index() + static_cast<column_t::index_t>(j));

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

// ==================== 数据验证功能实现 ====================

bool TXSheet::addDataValidation(const TXRange& range, const TXDataValidation& validation) {
    if (!range.isValid()) {
        setError("Invalid range for data validation");
        return false;
    }

    // 检查是否已存在相同范围的验证规则
    for (auto& pair : dataValidations_) {
        if (pair.first == range) {
            // 更新现有规则
            pair.second = validation;
            notifyComponentChange(ExcelComponent::BasicWorkbook);
            clearError();
            return true;
        }
    }

    // 添加新的验证规则
    dataValidations_.emplace_back(range, validation);
    notifyComponentChange(ExcelComponent::BasicWorkbook);
    clearError();
    return true;
}

bool TXSheet::removeDataValidation(const TXRange& range) {
    auto it = std::find_if(dataValidations_.begin(), dataValidations_.end(),
                          [&range](const auto& pair) { return pair.first == range; });

    if (it != dataValidations_.end()) {
        dataValidations_.erase(it);
        notifyComponentChange(ExcelComponent::BasicWorkbook);
        clearError();
        return true;
    }

    setError("Data validation not found for the specified range");
    return false;
}

void TXSheet::clearDataValidations() {
    if (!dataValidations_.empty()) {
        dataValidations_.clear();
        notifyComponentChange(ExcelComponent::BasicWorkbook);
    }
    clearError();
}

size_t TXSheet::getDataValidationCount() const {
    return dataValidations_.size();
}

bool TXSheet::hasDataValidation(const TXRange& range) const {
    return std::any_of(dataValidations_.begin(), dataValidations_.end(),
                      [&range](const auto& pair) { return pair.first == range; });
}

const std::vector<std::pair<TXRange, TXDataValidation>>& TXSheet::getDataValidations() const {
    return dataValidations_;
}

// ==================== 数据筛选功能实现 ====================

TXAutoFilter* TXSheet::enableAutoFilter(const TXRange& range) {
    autoFilter_ = std::make_unique<TXAutoFilter>(range);
    return autoFilter_.get();
}

void TXSheet::disableAutoFilter() {
    autoFilter_.reset();
}

TXAutoFilter* TXSheet::getAutoFilter() const {
    return autoFilter_.get();
}

bool TXSheet::hasAutoFilter() const {
    return autoFilter_ != nullptr;
}

// ==================== 图表操作实现 ====================

TXChart* TXSheet::addChart(std::unique_ptr<TXChart> chart) {
    if (!chart) {
        setError("Cannot add null chart");
        return nullptr;
    }

    // 设置图表名称（如果没有设置）
    if (chart->getName().empty()) {
        chart->setName("Chart" + std::to_string(nextChartId_++));
    }

    TXChart* chartPtr = chart.get();
    charts_.push_back(std::move(chart));

    clearError();
    notifyComponentChange(ExcelComponent::BasicWorkbook);
    return chartPtr;
}

TXColumnChart* TXSheet::addColumnChart(const std::string& title, const TXRange& dataRange,
                                       const std::pair<row_t, column_t>& position) {
    auto chart = std::make_unique<TXColumnChart>();
    chart->setName(title);  // 设置图表名称用于管理
    chart->setTitle(title); // 设置图表标题用于显示
    chart->setDataRange(this, dataRange);
    chart->setPosition(position.first, position.second);

    TXColumnChart* chartPtr = static_cast<TXColumnChart*>(chart.get());
    if (addChart(std::move(chart))) {
        return chartPtr;
    }
    return nullptr;
}

TXLineChart* TXSheet::addLineChart(const std::string& title, const TXRange& dataRange,
                                   const std::pair<row_t, column_t>& position) {
    auto chart = std::make_unique<TXLineChart>();
    chart->setName(title);  // 设置图表名称用于管理
    chart->setTitle(title); // 设置图表标题用于显示
    chart->setDataRange(this, dataRange);
    chart->setPosition(position.first, position.second);

    TXLineChart* chartPtr = static_cast<TXLineChart*>(chart.get());
    if (addChart(std::move(chart))) {
        return chartPtr;
    }
    return nullptr;
}

TXPieChart* TXSheet::addPieChart(const std::string& title, const TXRange& dataRange,
                                 const std::pair<row_t, column_t>& position) {
    auto chart = std::make_unique<TXPieChart>();
    chart->setName(title);  // 设置图表名称用于管理
    chart->setTitle(title); // 设置图表标题用于显示
    chart->setDataRange(this, dataRange);
    chart->setPosition(position.first, position.second);

    TXPieChart* chartPtr = static_cast<TXPieChart*>(chart.get());
    if (addChart(std::move(chart))) {
        return chartPtr;
    }
    return nullptr;
}

TXScatterChart* TXSheet::addScatterChart(const std::string& title, const TXRange& dataRange,
                                         const std::pair<row_t, column_t>& position) {
    auto chart = std::make_unique<TXScatterChart>();
    chart->setName(title);  // 设置图表名称用于管理
    chart->setTitle(title); // 设置图表标题用于显示
    chart->setDataRange(this, dataRange);
    chart->setPosition(position.first, position.second);

    TXScatterChart* chartPtr = static_cast<TXScatterChart*>(chart.get());
    if (addChart(std::move(chart))) {
        return chartPtr;
    }
    return nullptr;
}

bool TXSheet::removeChart(const std::string& chartName) {
    auto it = std::find_if(charts_.begin(), charts_.end(),
                           [&chartName](const std::unique_ptr<TXChart>& chart) {
                               return chart->getName() == chartName;
                           });

    if (it != charts_.end()) {
        charts_.erase(it);
        clearError();
        return true;
    }

    setError("Chart not found: " + chartName);
    return false;
}

TXChart* TXSheet::getChart(const std::string& chartName) {
    auto it = std::find_if(charts_.begin(), charts_.end(),
                           [&chartName](const std::unique_ptr<TXChart>& chart) {
                               return chart->getName() == chartName;
                           });

    return (it != charts_.end()) ? it->get() : nullptr;
}

const TXChart* TXSheet::getChart(const std::string& chartName) const {
    auto it = std::find_if(charts_.begin(), charts_.end(),
                           [&chartName](const std::unique_ptr<TXChart>& chart) {
                               return chart->getName() == chartName;
                           });

    return (it != charts_.end()) ? it->get() : nullptr;
}

std::vector<TXChart*> TXSheet::getAllCharts() {
    std::vector<TXChart*> result;
    result.reserve(charts_.size());
    for (const auto& chart : charts_) {
        result.push_back(chart.get());
    }
    return result;
}

std::vector<const TXChart*> TXSheet::getAllCharts() const {
    std::vector<const TXChart*> result;
    result.reserve(charts_.size());
    for (const auto& chart : charts_) {
        result.push_back(chart.get());
    }
    return result;
}

std::size_t TXSheet::getChartCount() const {
    return charts_.size();
}

} // namespace TinaXlsx
