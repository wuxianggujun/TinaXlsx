#include "TinaXlsx/TXRowColumnManager.hpp"
#include "TinaXlsx/TXCellManager.hpp"
#include <algorithm>
#include <cmath>
#include <set>

namespace TinaXlsx {

// ==================== 行操作 ====================

bool TXRowColumnManager::insertRows(row_t row, row_t count, TXCellManager& cellManager) {
    if (!isValidRow(row) || count.index() == 0) {
        return false;
    }

    // 创建坐标变换函数
    auto transform = [row, count](const TXCoordinate& coord) -> TXCoordinate {
        if (coord.getRow() >= row) {
            // 向下移动
            return TXCoordinate(row_t(coord.getRow().index() + count.index()), coord.getCol());
        }
        return coord; // 保持不变
    };

    // 应用变换
    cellManager.transformCells(transform);

    // 更新行高映射
    std::unordered_map<row_t::index_t, double> newRowHeights;
    for (const auto& pair : rowHeights_) {
        row_t::index_t oldRow = pair.first;
        double height = pair.second;
        
        if (oldRow >= row.index()) {
            newRowHeights[oldRow + count.index()] = height;
        } else {
            newRowHeights[oldRow] = height;
        }
    }
    rowHeights_ = std::move(newRowHeights);

    // 更新隐藏行映射
    std::unordered_map<row_t::index_t, bool> newHiddenRows;
    for (const auto& pair : hiddenRows_) {
        row_t::index_t oldRow = pair.first;
        bool hidden = pair.second;
        
        if (oldRow >= row.index()) {
            newHiddenRows[oldRow + count.index()] = hidden;
        } else {
            newHiddenRows[oldRow] = hidden;
        }
    }
    hiddenRows_ = std::move(newHiddenRows);

    return true;
}

bool TXRowColumnManager::deleteRows(row_t row, row_t count, TXCellManager& cellManager) {
    if (!isValidRow(row) || count.index() == 0) {
        return false;
    }

    // 删除指定范围内的单元格
    TXRange deleteRange(
        TXCoordinate(row, column_t(1)),
        TXCoordinate(row_t(row.index() + count.index() - 1), column_t::last())
    );
    cellManager.removeCellsInRange(deleteRange);

    // 创建坐标变换函数
    auto transform = [row, count](const TXCoordinate& coord) -> TXCoordinate {
        if (coord.getRow() < row) {
            return coord; // 保持不变
        } else if (coord.getRow().index() >= row.index() + count.index()) {
            // 向上移动
            return TXCoordinate(row_t(coord.getRow().index() - count.index()), coord.getCol());
        } else {
            // 在删除范围内，返回无效坐标
            return TXCoordinate();
        }
    };

    // 应用变换
    cellManager.transformCells(transform);

    // 更新行高映射
    std::unordered_map<row_t::index_t, double> newRowHeights;
    for (const auto& pair : rowHeights_) {
        row_t::index_t oldRow = pair.first;
        double height = pair.second;
        
        if (oldRow < row.index()) {
            newRowHeights[oldRow] = height;
        } else if (oldRow >= row.index() + count.index()) {
            newRowHeights[oldRow - count.index()] = height;
        }
        // 在删除范围内的行高被丢弃
    }
    rowHeights_ = std::move(newRowHeights);

    // 更新隐藏行映射
    std::unordered_map<row_t::index_t, bool> newHiddenRows;
    for (const auto& pair : hiddenRows_) {
        row_t::index_t oldRow = pair.first;
        bool hidden = pair.second;
        
        if (oldRow < row.index()) {
            newHiddenRows[oldRow] = hidden;
        } else if (oldRow >= row.index() + count.index()) {
            newHiddenRows[oldRow - count.index()] = hidden;
        }
        // 在删除范围内的隐藏状态被丢弃
    }
    hiddenRows_ = std::move(newHiddenRows);

    return true;
}

bool TXRowColumnManager::setRowHeight(row_t row, double height) {
    if (!isValidRow(row) || !isValidSize(height, false)) {
        return false;
    }
    
    rowHeights_[row.index()] = height;
    return true;
}

double TXRowColumnManager::getRowHeight(row_t row) const {
    auto it = rowHeights_.find(row.index());
    return (it != rowHeights_.end()) ? it->second : DEFAULT_ROW_HEIGHT;
}

bool TXRowColumnManager::setRowHidden(row_t row, bool hidden) {
    if (!isValidRow(row)) {
        return false;
    }
    
    if (hidden) {
        hiddenRows_[row.index()] = true;
    } else {
        hiddenRows_.erase(row.index());
    }
    return true;
}

bool TXRowColumnManager::isRowHidden(row_t row) const {
    auto it = hiddenRows_.find(row.index());
    return (it != hiddenRows_.end()) ? it->second : false;
}

// ==================== 列操作 ====================

bool TXRowColumnManager::insertColumns(column_t col, column_t count, TXCellManager& cellManager) {
    if (!isValidColumn(col) || count.index() == 0) {
        return false;
    }

    // 创建坐标变换函数
    auto transform = [col, count](const TXCoordinate& coord) -> TXCoordinate {
        if (coord.getCol() >= col) {
            // 向右移动
            return TXCoordinate(coord.getRow(), column_t(coord.getCol().index() + count.index()));
        }
        return coord; // 保持不变
    };

    // 应用变换
    cellManager.transformCells(transform);

    // 更新列宽映射
    std::unordered_map<column_t::index_t, double> newColumnWidths;
    for (const auto& pair : columnWidths_) {
        column_t::index_t oldCol = pair.first;
        double width = pair.second;
        
        if (oldCol >= col.index()) {
            newColumnWidths[oldCol + count.index()] = width;
        } else {
            newColumnWidths[oldCol] = width;
        }
    }
    columnWidths_ = std::move(newColumnWidths);

    // 更新隐藏列映射
    std::unordered_map<column_t::index_t, bool> newHiddenColumns;
    for (const auto& pair : hiddenColumns_) {
        column_t::index_t oldCol = pair.first;
        bool hidden = pair.second;
        
        if (oldCol >= col.index()) {
            newHiddenColumns[oldCol + count.index()] = hidden;
        } else {
            newHiddenColumns[oldCol] = hidden;
        }
    }
    hiddenColumns_ = std::move(newHiddenColumns);

    return true;
}

bool TXRowColumnManager::deleteColumns(column_t col, column_t count, TXCellManager& cellManager) {
    if (!isValidColumn(col) || count.index() == 0) {
        return false;
    }

    // 删除指定范围内的单元格
    TXRange deleteRange(
        TXCoordinate(row_t(1), col),
        TXCoordinate(row_t::last(), column_t(col.index() + count.index() - 1))
    );
    cellManager.removeCellsInRange(deleteRange);

    // 创建坐标变换函数
    auto transform = [col, count](const TXCoordinate& coord) -> TXCoordinate {
        if (coord.getCol() < col) {
            return coord; // 保持不变
        } else if (coord.getCol().index() >= col.index() + count.index()) {
            // 向左移动
            return TXCoordinate(coord.getRow(), column_t(coord.getCol().index() - count.index()));
        } else {
            // 在删除范围内，返回无效坐标
            return TXCoordinate();
        }
    };

    // 应用变换
    cellManager.transformCells(transform);

    // 更新列宽映射
    std::unordered_map<column_t::index_t, double> newColumnWidths;
    for (const auto& pair : columnWidths_) {
        column_t::index_t oldCol = pair.first;
        double width = pair.second;
        
        if (oldCol < col.index()) {
            newColumnWidths[oldCol] = width;
        } else if (oldCol >= col.index() + count.index()) {
            newColumnWidths[oldCol - count.index()] = width;
        }
        // 在删除范围内的列宽被丢弃
    }
    columnWidths_ = std::move(newColumnWidths);

    // 更新隐藏列映射
    std::unordered_map<column_t::index_t, bool> newHiddenColumns;
    for (const auto& pair : hiddenColumns_) {
        column_t::index_t oldCol = pair.first;
        bool hidden = pair.second;
        
        if (oldCol < col.index()) {
            newHiddenColumns[oldCol] = hidden;
        } else if (oldCol >= col.index() + count.index()) {
            newHiddenColumns[oldCol - count.index()] = hidden;
        }
        // 在删除范围内的隐藏状态被丢弃
    }
    hiddenColumns_ = std::move(newHiddenColumns);

    return true;
}

bool TXRowColumnManager::setColumnWidth(column_t col, double width) {
    if (!isValidColumn(col) || !isValidSize(width, true)) {
        return false;
    }
    
    columnWidths_[col.index()] = width;
    return true;
}

double TXRowColumnManager::getColumnWidth(column_t col) const {
    auto it = columnWidths_.find(col.index());
    return (it != columnWidths_.end()) ? it->second : DEFAULT_COLUMN_WIDTH;
}

bool TXRowColumnManager::setColumnHidden(column_t col, bool hidden) {
    if (!isValidColumn(col)) {
        return false;
    }
    
    if (hidden) {
        hiddenColumns_[col.index()] = true;
    } else {
        hiddenColumns_.erase(col.index());
    }
    return true;
}

bool TXRowColumnManager::isColumnHidden(column_t col) const {
    auto it = hiddenColumns_.find(col.index());
    return (it != hiddenColumns_.end()) ? it->second : false;
}

// ==================== 自动调整 ====================

double TXRowColumnManager::autoFitColumnWidth(column_t col, const TXCellManager& cellManager, 
                                             double minWidth, double maxWidth) {
    if (!isValidColumn(col)) {
        return getColumnWidth(col);
    }
    
    double maxContentWidth = minWidth;
    
    // 遍历该列的所有单元格，计算最大内容宽度
    for (auto it = cellManager.begin(); it != cellManager.end(); ++it) {
        const auto& coord = it->first;
        const auto& cell = it->second;
        
        if (coord.getCol() == col && !cell.isEmpty()) {
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

double TXRowColumnManager::autoFitRowHeight(row_t row, const TXCellManager& cellManager,
                                           double minHeight, double maxHeight) {
    if (!isValidRow(row)) {
        return getRowHeight(row);
    }
    
    double maxContentHeight = minHeight;
    
    // 遍历该行的所有单元格，计算最大内容高度
    for (auto it = cellManager.begin(); it != cellManager.end(); ++it) {
        const auto& coord = it->first;
        const auto& cell = it->second;
        
        if (coord.getRow() == row && !cell.isEmpty()) {
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

std::size_t TXRowColumnManager::autoFitAllColumnWidths(const TXCellManager& cellManager,
                                                      double minWidth, double maxWidth) {
    std::set<column_t> usedColumns;

    // 收集所有使用的列
    for (auto it = cellManager.begin(); it != cellManager.end(); ++it) {
        if (!it->second.isEmpty()) {
            usedColumns.insert(it->first.getCol());
        }
    }

    std::size_t count = 0;
    for (column_t col : usedColumns) {
        autoFitColumnWidth(col, cellManager, minWidth, maxWidth);
        ++count;
    }

    return count;
}

std::size_t TXRowColumnManager::autoFitAllRowHeights(const TXCellManager& cellManager,
                                                    double minHeight, double maxHeight) {
    std::set<row_t> usedRows;

    // 收集所有使用的行
    for (auto it = cellManager.begin(); it != cellManager.end(); ++it) {
        if (!it->second.isEmpty()) {
            usedRows.insert(it->first.getRow());
        }
    }

    std::size_t count = 0;
    for (row_t row : usedRows) {
        autoFitRowHeight(row, cellManager, minHeight, maxHeight);
        ++count;
    }

    return count;
}

// ==================== 批量操作 ====================

std::size_t TXRowColumnManager::setRowHeights(const std::vector<std::pair<row_t, double>>& heights) {
    std::size_t count = 0;
    for (const auto& pair : heights) {
        if (setRowHeight(pair.first, pair.second)) {
            ++count;
        }
    }
    return count;
}

std::size_t TXRowColumnManager::setColumnWidths(const std::vector<std::pair<column_t, double>>& widths) {
    std::size_t count = 0;
    for (const auto& pair : widths) {
        if (setColumnWidth(pair.first, pair.second)) {
            ++count;
        }
    }
    return count;
}

void TXRowColumnManager::clear() {
    rowHeights_.clear();
    columnWidths_.clear();
    hiddenRows_.clear();
    hiddenColumns_.clear();
}

// ==================== 私有辅助方法 ====================

double TXRowColumnManager::calculateTextWidth(const std::string& text, double fontSize,
                                             const std::string& fontName) const {
    // 简化的文本宽度计算
    // 实际实现应该考虑字体度量
    double charWidth = fontSize * 0.6; // 近似值

    // 根据字体名称调整宽度系数（简化处理）
    double fontFactor = 1.0;
    if (fontName == "Arial" || fontName == "Calibri") {
        fontFactor = 0.9;
    } else if (fontName == "Times New Roman") {
        fontFactor = 0.8;
    }

    return text.length() * charWidth * fontFactor / 7.0; // 转换为Excel字符单位
}

double TXRowColumnManager::calculateTextHeight(const std::string& text, double fontSize,
                                              double columnWidth) const {
    // 简化的文本高度计算
    double lineHeight = fontSize * 1.2; // 行高通常是字体大小的1.2倍

    // 计算换行数
    double textWidth = calculateTextWidth(text, fontSize);
    double excelColumnWidth = columnWidth * 7.0; // 转换为像素近似值
    int lines = static_cast<int>(std::ceil(textWidth / excelColumnWidth));

    return std::max(1, lines) * lineHeight;
}

bool TXRowColumnManager::isValidRow(row_t row) const {
    return row.is_valid() && row.index() > 0 && row.index() <= 1048576; // Excel最大行数
}

bool TXRowColumnManager::isValidColumn(column_t col) const {
    return col.is_valid() && col.index() > 0 && col.index() <= 16384; // Excel最大列数
}

bool TXRowColumnManager::isValidSize(double size, bool isWidth) const {
    if (isWidth) {
        return size >= 0.0 && size <= 255.0; // Excel列宽范围
    } else {
        return size >= 0.0 && size <= 409.0; // Excel行高范围
    }
}

} // namespace TinaXlsx
