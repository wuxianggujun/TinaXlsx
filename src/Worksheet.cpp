/**
 * @file Worksheet.cpp
 * @brief Excel工作表类实现
 */

#include "TinaXlsx/Worksheet.hpp"
#include <xlsxwriter.h>
#include <memory>

namespace TinaXlsx {

struct Worksheet::Impl {
    lxw_worksheet* worksheet = nullptr;
    lxw_workbook* workbook = nullptr;
    std::string name;
    RowIndex currentRow = 0;
    
    Impl(lxw_worksheet* ws, lxw_workbook* wb, const std::string& n) 
        : worksheet(ws), workbook(wb), name(n) {}
};

Worksheet::Worksheet(lxw_worksheet* worksheet, lxw_workbook* workbook, const std::string& name)
    : pImpl_(std::make_unique<Impl>(worksheet, workbook, name)) {
}

Worksheet::Worksheet(Worksheet&& other) noexcept 
    : pImpl_(std::move(other.pImpl_)) {
}

Worksheet& Worksheet::operator=(Worksheet&& other) noexcept {
    if (this != &other) {
        pImpl_ = std::move(other.pImpl_);
    }
    return *this;
}

Worksheet::~Worksheet() = default;

std::string Worksheet::getName() const {
    return pImpl_->name;
}

void Worksheet::writeCell(const CellPosition& position, const CellValue& value, Format* format) {
    writeCell(position.row, position.column, value, format);
}

void Worksheet::writeCell(RowIndex row, ColumnIndex column, const CellValue& value, Format* format) {
    if (!pImpl_->worksheet) return;
    
    lxw_format* fmt = format ? format->getInternalFormat() : nullptr;
    
    // 使用类型安全枚举而不是硬编码数字
    const CellValueType type = getCellValueType(value);
    
    switch (type) {
        case CellValueType::String:
            worksheet_write_string(pImpl_->worksheet, row, column, 
                                 std::get<std::string>(value).c_str(), fmt);
            break;
        case CellValueType::Double:
            worksheet_write_number(pImpl_->worksheet, row, column, 
                                 std::get<double>(value), fmt);
            break;
        case CellValueType::Integer:
            worksheet_write_number(pImpl_->worksheet, row, column, 
                                 static_cast<double>(std::get<Integer>(value)), fmt);
            break;
        case CellValueType::Boolean:
            worksheet_write_boolean(pImpl_->worksheet, row, column, 
                                  std::get<bool>(value), fmt);
            break;
        case CellValueType::Empty:
            worksheet_write_blank(pImpl_->worksheet, row, column, fmt);
            break;
        default:
            worksheet_write_blank(pImpl_->worksheet, row, column, fmt);
            break;
    }
}

void Worksheet::writeString(const CellPosition& position, const std::string& value, Format* format) {
    if (!pImpl_->worksheet) return;
    
    lxw_format* fmt = format ? format->getInternalFormat() : nullptr;
    worksheet_write_string(pImpl_->worksheet, position.row, position.column, value.c_str(), fmt);
}

void Worksheet::writeNumber(const CellPosition& position, double value, Format* format) {
    if (!pImpl_->worksheet) return;
    
    lxw_format* fmt = format ? format->getInternalFormat() : nullptr;
    worksheet_write_number(pImpl_->worksheet, position.row, position.column, value, fmt);
}

void Worksheet::writeInteger(const CellPosition& position, Integer value, Format* format) {
    writeNumber(position, static_cast<double>(value), format);
}

void Worksheet::writeBoolean(const CellPosition& position, bool value, Format* format) {
    if (!pImpl_->worksheet) return;
    
    lxw_format* fmt = format ? format->getInternalFormat() : nullptr;
    worksheet_write_boolean(pImpl_->worksheet, position.row, position.column, value, fmt);
}

void Worksheet::writeDateTime(const CellPosition& position, double datetime, Format* format) {
    if (!pImpl_->worksheet) return;
    
    lxw_format* fmt = format ? format->getInternalFormat() : nullptr;
    worksheet_write_number(pImpl_->worksheet, position.row, position.column, datetime, fmt);
}

void Worksheet::writeFormula(const CellPosition& position, const std::string& formula, Format* format) {
    if (!pImpl_->worksheet) return;
    
    lxw_format* fmt = format ? format->getInternalFormat() : nullptr;
    worksheet_write_formula(pImpl_->worksheet, position.row, position.column, formula.c_str(), fmt);
}

void Worksheet::writeUrl(const CellPosition& position, const std::string& url, const std::string& text, Format* format) {
    if (!pImpl_->worksheet) return;
    
    lxw_format* fmt = format ? format->getInternalFormat() : nullptr;
    if (text.empty()) {
        worksheet_write_url(pImpl_->worksheet, position.row, position.column, url.c_str(), fmt);
    } else {
        worksheet_write_url_opt(pImpl_->worksheet, position.row, position.column, url.c_str(), fmt, text.c_str(), nullptr);
    }
}

void Worksheet::writeRow(RowIndex rowIndex, const RowData& rowData, ColumnIndex startColumn, Format* format) {
    if (!pImpl_->worksheet) return;
    
    lxw_format* fmt = format ? format->getInternalFormat() : nullptr;
    
    // 优化：直接调用底层libxlsxwriter函数，减少中间函数调用
    for (ColumnIndex col = 0; col < rowData.size(); ++col) {
        const ColumnIndex currentCol = startColumn + col;
        const CellValue& value = rowData[col];
        
        // 使用类型安全枚举而不是硬编码数字
        const CellValueType type = getCellValueType(value);
        
        switch (type) {
            case CellValueType::String:
                worksheet_write_string(pImpl_->worksheet, rowIndex, currentCol, 
                                     std::get<std::string>(value).c_str(), fmt);
                break;
            case CellValueType::Double:
                worksheet_write_number(pImpl_->worksheet, rowIndex, currentCol, 
                                     std::get<double>(value), fmt);
                break;
            case CellValueType::Integer:
                worksheet_write_number(pImpl_->worksheet, rowIndex, currentCol, 
                                     static_cast<double>(std::get<Integer>(value)), fmt);
                break;
            case CellValueType::Boolean:
                worksheet_write_boolean(pImpl_->worksheet, rowIndex, currentCol, 
                                      std::get<bool>(value), fmt);
                break;
            case CellValueType::Empty:
                worksheet_write_blank(pImpl_->worksheet, rowIndex, currentCol, fmt);
                break;
            default:
                worksheet_write_blank(pImpl_->worksheet, rowIndex, currentCol, fmt);
                break;
        }
    }
}

void Worksheet::writeTable(const CellPosition& startPosition, const TableData& tableData, Format* format) {
    if (!pImpl_->worksheet) return;
    
    for (RowIndex row = 0; row < tableData.size(); ++row) {
        writeRow(startPosition.row + row, tableData[row], startPosition.column, format);
    }
}

void Worksheet::writeBatch(const CellRange& range, const TableData& tableData, Format* format) {
    if (!range.isValid()) {
        throw InvalidArgumentException("Invalid cell range");
    }
    
    if (!pImpl_->worksheet) return;
    
    lxw_format* fmt = format ? format->getInternalFormat() : nullptr;
    const RowIndex maxRows = std::min(range.rowCount(), static_cast<RowIndex>(tableData.size()));
    
    // 优化：使用更直接的循环和内联写入操作
    for (RowIndex row = 0; row < maxRows; ++row) {
        const auto& rowData = tableData[row];
        const ColumnIndex maxCols = std::min(range.columnCount(), static_cast<ColumnIndex>(rowData.size()));
        const RowIndex actualRow = range.start.row + row;
        
        for (ColumnIndex col = 0; col < maxCols; ++col) {
            const ColumnIndex actualCol = range.start.column + col;
            const CellValue& value = rowData[col];
            
            // 内联优化的写入操作 - 使用类型安全枚举
            const CellValueType type = getCellValueType(value);
            
            switch (type) {
                case CellValueType::String:
                    worksheet_write_string(pImpl_->worksheet, actualRow, actualCol, 
                                         std::get<std::string>(value).c_str(), fmt);
                    break;
                case CellValueType::Double:
                    worksheet_write_number(pImpl_->worksheet, actualRow, actualCol, 
                                         std::get<double>(value), fmt);
                    break;
                case CellValueType::Integer:
                    worksheet_write_number(pImpl_->worksheet, actualRow, actualCol, 
                                         static_cast<double>(std::get<Integer>(value)), fmt);
                    break;
                case CellValueType::Boolean:
                    worksheet_write_boolean(pImpl_->worksheet, actualRow, actualCol, 
                                          std::get<bool>(value), fmt);
                    break;
                case CellValueType::Empty:
                    worksheet_write_blank(pImpl_->worksheet, actualRow, actualCol, fmt);
                    break;
                default:
                    worksheet_write_blank(pImpl_->worksheet, actualRow, actualCol, fmt);
                    break;
            }
        }
    }
}

void Worksheet::mergeCells(const CellRange& range, const CellValue& value, Format* format) {
    if (!pImpl_->worksheet || !range.isValid()) return;
    
    // 先写入值到左上角单元格
    std::string cellText = "";
    if (!std::holds_alternative<std::monostate>(value)) {
        writeCell(range.start, value, format);
        // 提取文本值以传递给worksheet_merge_range
        std::visit([&cellText](const auto& v) {
            using T = std::decay_t<decltype(v)>;
            if constexpr (std::is_same_v<T, std::string>) {
                cellText = v;
            } else if constexpr (std::is_same_v<T, double>) {
                cellText = FastConvert::doubleToString(v);
            } else if constexpr (std::is_same_v<T, Integer>) {
                cellText = FastConvert::integerToString(v);
            } else if constexpr (std::is_same_v<T, bool>) {
                cellText = FastConvert::boolToString(v);
            }
        }, value);
    }
    
    // 合并单元格，使用实际的文本值
    lxw_format* fmt = format ? format->getInternalFormat() : nullptr;
    worksheet_merge_range(pImpl_->worksheet, 
                         range.start.row, range.start.column,
                         range.end.row, range.end.column,
                         cellText.c_str(), fmt);
}

void Worksheet::setRowHeight(RowIndex rowIndex, double height) {
    if (!pImpl_->worksheet) return;
    
    worksheet_set_row(pImpl_->worksheet, rowIndex, height, nullptr);
}

void Worksheet::setRowHeight(RowIndex startRow, RowIndex endRow, double height) {
    for (RowIndex row = startRow; row <= endRow; ++row) {
        setRowHeight(row, height);
    }
}

void Worksheet::setColumnWidth(ColumnIndex columnIndex, double width) {
    if (!pImpl_->worksheet) return;
    
    worksheet_set_column(pImpl_->worksheet, columnIndex, columnIndex, width, nullptr);
}

void Worksheet::setColumnWidth(ColumnIndex startColumn, ColumnIndex endColumn, double width) {
    if (!pImpl_->worksheet) return;
    
    worksheet_set_column(pImpl_->worksheet, startColumn, endColumn, width, nullptr);
}

void Worksheet::autoFitColumn(ColumnIndex columnIndex) {
    // libxlsxwriter没有内置的自动调整列宽功能
    // 这里设置一个合理的默认宽度
    setColumnWidth(columnIndex, 15.0);
}

void Worksheet::hideRow(RowIndex rowIndex) {
    if (!pImpl_->worksheet) return;
    
    lxw_row_col_options options = {};
    options.hidden = LXW_TRUE;
    worksheet_set_row_opt(pImpl_->worksheet, rowIndex, LXW_DEF_ROW_HEIGHT, nullptr, &options);
}

void Worksheet::hideColumn(ColumnIndex columnIndex) {
    if (!pImpl_->worksheet) return;
    
    lxw_row_col_options options = {};
    options.hidden = LXW_TRUE;
    worksheet_set_column_opt(pImpl_->worksheet, columnIndex, columnIndex, LXW_DEF_COL_WIDTH, nullptr, &options);
}

void Worksheet::setRowFormat(RowIndex rowIndex, Format* format) {
    if (!pImpl_->worksheet) return;
    
    lxw_format* fmt = format ? format->getInternalFormat() : nullptr;
    worksheet_set_row(pImpl_->worksheet, rowIndex, LXW_DEF_ROW_HEIGHT, fmt);
}

void Worksheet::setColumnFormat(ColumnIndex columnIndex, Format* format) {
    if (!pImpl_->worksheet) return;
    
    lxw_format* fmt = format ? format->getInternalFormat() : nullptr;
    worksheet_set_column(pImpl_->worksheet, columnIndex, columnIndex, LXW_DEF_COL_WIDTH, fmt);
}

void Worksheet::freezePanes(RowIndex row, ColumnIndex column) {
    if (!pImpl_->worksheet) return;
    
    worksheet_freeze_panes(pImpl_->worksheet, row, column);
}

void Worksheet::splitPanes(RowIndex row, ColumnIndex column) {
    if (!pImpl_->worksheet) return;
    
    worksheet_split_panes(pImpl_->worksheet, row * 15.0, column * 8.43); // 近似像素值
}

void Worksheet::setPrintArea(const CellRange& range) {
    if (!pImpl_->worksheet || !range.isValid()) return;
    
    worksheet_print_area(pImpl_->worksheet, 
                        range.start.row, range.start.column,
                        range.end.row, range.end.column);
}

void Worksheet::setRepeatRows(RowIndex startRow, RowIndex endRow) {
    if (!pImpl_->worksheet) return;
    
    worksheet_repeat_rows(pImpl_->worksheet, startRow, endRow);
}

void Worksheet::setRepeatColumns(ColumnIndex startColumn, ColumnIndex endColumn) {
    if (!pImpl_->worksheet) return;
    
    worksheet_repeat_columns(pImpl_->worksheet, startColumn, endColumn);
}

void Worksheet::setLandscape(bool landscape) {
    if (!pImpl_->worksheet) return;
    
    if (landscape) {
        worksheet_set_landscape(pImpl_->worksheet);
    } else {
        worksheet_set_portrait(pImpl_->worksheet);
    }
}

void Worksheet::setPaperSize(int paperSize) {
    if (!pImpl_->worksheet) return;
    
    worksheet_set_paper(pImpl_->worksheet, paperSize);
}

void Worksheet::setMargins(double left, double right, double top, double bottom) {
    if (!pImpl_->worksheet) return;
    
    worksheet_set_margins(pImpl_->worksheet, left, right, top, bottom);
}

void Worksheet::setHeader(const std::string& header) {
    if (!pImpl_->worksheet) return;
    
    worksheet_set_header(pImpl_->worksheet, header.c_str());
}

void Worksheet::setFooter(const std::string& footer) {
    if (!pImpl_->worksheet) return;
    
    worksheet_set_footer(pImpl_->worksheet, footer.c_str());
}

void Worksheet::protect(const std::string& password) {
    if (!pImpl_->worksheet) return;
    
    if (password.empty()) {
        worksheet_protect(pImpl_->worksheet, nullptr, nullptr);
    } else {
        worksheet_protect(pImpl_->worksheet, password.c_str(), nullptr);
    }
}

void Worksheet::unprotect() {
    if (!pImpl_->worksheet) return;
    
    worksheet_protect(pImpl_->worksheet, "", nullptr);
}

void Worksheet::showGridlines(bool show) {
    if (!pImpl_->worksheet) return;
    
    if (show) {
        worksheet_gridlines(pImpl_->worksheet, LXW_SHOW_SCREEN_GRIDLINES | LXW_SHOW_PRINT_GRIDLINES);
    } else {
        worksheet_gridlines(pImpl_->worksheet, LXW_HIDE_ALL_GRIDLINES);
    }
}

void Worksheet::setZoom(int scale) {
    if (!pImpl_->worksheet) return;
    
    worksheet_set_zoom(pImpl_->worksheet, scale);
}

void Worksheet::setSelection(const CellRange& range) {
    if (!pImpl_->worksheet || !range.isValid()) return;
    
    worksheet_set_selection(pImpl_->worksheet, 
                           range.start.row, range.start.column,
                           range.end.row, range.end.column);
}

void Worksheet::insertImage(const CellPosition& position, const std::string& filename, double xScale, double yScale) {
    if (!pImpl_->worksheet) return;
    
    lxw_image_options options = {};
    options.x_scale = xScale;
    options.y_scale = yScale;
    
    worksheet_insert_image_opt(pImpl_->worksheet, position.row, position.column, filename.c_str(), &options);
}

void Worksheet::insertChart(const CellPosition& position, int chartType, const CellRange& dataRange) {
    if (!pImpl_->worksheet) return;
    
    // 创建图表需要使用workbook_add_chart，这里简化处理
    // 实际实现需要更复杂的图表配置
}

void Worksheet::addDataValidation(const CellRange& range, int validationType, int criteria, 
                                 const std::string& value1, const std::string& value2) {
    if (!pImpl_->worksheet || !range.isValid()) return;
    
    lxw_data_validation validation = {};
    validation.validate = validationType;
    validation.criteria = criteria;
    
    if (!value1.empty()) {
        validation.value_formula = const_cast<char*>(value1.c_str());
    }
    if (!value2.empty()) {
        validation.maximum_formula = const_cast<char*>(value2.c_str());
    }
    
    worksheet_data_validation_range(pImpl_->worksheet, 
                                   range.start.row, range.start.column,
                                   range.end.row, range.end.column,
                                   &validation);
}

void Worksheet::addConditionalFormat(const CellRange& range, int type, int criteria, double value, Format* format) {
    if (!pImpl_->worksheet || !range.isValid()) return;
    
    lxw_conditional_format conditionalFormat = {};
    conditionalFormat.type = type;
    conditionalFormat.criteria = criteria;
    conditionalFormat.value = value;
    
    if (format) {
        conditionalFormat.format = format->getInternalFormat();
    }
    
    worksheet_conditional_format_range(pImpl_->worksheet,
                                      range.start.row, range.start.column,
                                      range.end.row, range.end.column,
                                      &conditionalFormat);
}

void Worksheet::addAutoFilter(const CellRange& range) {
    if (!pImpl_->worksheet || !range.isValid()) return;
    
    worksheet_autofilter(pImpl_->worksheet, 
                        range.start.row, range.start.column,
                        range.end.row, range.end.column);
}

lxw_worksheet* Worksheet::getInternalWorksheet() const {
    return pImpl_->worksheet;
}

lxw_workbook* Worksheet::getInternalWorkbook() const {
    return pImpl_->workbook;
}

RowIndex Worksheet::appendRow(const RowData& rowData, Format* format) {
    writeRow(pImpl_->currentRow, rowData, 0, format);
    return pImpl_->currentRow++;
}

void Worksheet::writeBatchWithCallback(RowIndex startRow, RowIndex rowCount, const RowWriteCallback& callback) {
    if (!callback) {
        throw InvalidArgumentException("回调函数不能为空");
    }
    
    for (RowIndex i = 0; i < rowCount; ++i) {
        callback(startRow + i);
    }
}

} // namespace TinaXlsx 