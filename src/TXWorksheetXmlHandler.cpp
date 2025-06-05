//
// Created by wuxianggujun on 2025/5/29.
//

#include "TinaXlsx/TXWorksheetXmlHandler.hpp"
#include "TinaXlsx/TXCompactCell.hpp"
#include "TinaXlsx/TXFormula.hpp"
#include "TinaXlsx/TXNumberUtils.hpp"
#include "TinaXlsx/TXPugiStreamWriter.hpp"
#include "TinaXlsx/TXBatchXMLGenerator.hpp"
#include "TinaXlsx/TXUnifiedMemoryManager.hpp"
#include <variant>

#include "TinaXlsx/TXSharedStringsPool.hpp"

namespace TinaXlsx
{

    TXResult<void> TXWorksheetXmlHandler::save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) {
        const TXSheet* sheet = context.sheets[m_sheetIndex].get();
        TXRange usedRange = sheet->getUsedRange();

        // 估算单元格数量，决定使用哪种写入策略
        size_t estimatedCells = 0;
        if (usedRange.isValid()) {
            estimatedCells = (usedRange.getEnd().getRow().index() - usedRange.getStart().getRow().index() + 1) *
                           (usedRange.getEnd().getCol().index() - usedRange.getStart().getCol().index() + 1);
        }

        // 对于大量数据使用流式写入，小量数据使用DOM方式
        if (estimatedCells > 5000) {
            return saveWithStreamWriter(zipWriter, context);
        }

        // 🚀 可选：使用批处理优化（当批处理器可用时）
        if (batch_generator_ && estimatedCells > 1000) {
            auto batch_result = saveWithBatchProcessor(zipWriter, context);
            if (batch_result.isOk()) {
                return batch_result;
            }
            // 如果批处理失败，回退到DOM方式
        }

        // 小数据量使用原有的DOM方式
        XmlNodeBuilder worksheet("worksheet");
        worksheet.addAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                 .addAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        // 添加维度信息
        XmlNodeBuilder dimension("dimension");
        if (usedRange.isValid()) {
            dimension.addAttribute("ref", usedRange.toAddress());
        } else {
            dimension.addAttribute("ref", "A1:A1");
        }
        worksheet.addChild(dimension);

        // 添加列宽信息
        auto& rowColManager = sheet->getRowColumnManager();
        const auto& customColumnWidths = rowColManager.getCustomColumnWidths();
        if (!customColumnWidths.empty()) {
            XmlNodeBuilder cols("cols");

            for (const auto& pair : customColumnWidths) {
                column_t::index_t colIndex = pair.first;
                double width = pair.second;

                XmlNodeBuilder col("col");

                // 格式化宽度值，保留合理的小数位数
                std::ostringstream widthStream;
                widthStream << std::fixed << std::setprecision(2) << width;
                std::string widthStr = widthStream.str();
                // 移除尾随的零和小数点
                widthStr.erase(widthStr.find_last_not_of('0') + 1, std::string::npos);
                if (widthStr.back() == '.') {
                    widthStr.pop_back();
                }

                col.addAttribute("min", std::to_string(colIndex))
                   .addAttribute("max", std::to_string(colIndex))
                   .addAttribute("width", widthStr)
                   .addAttribute("customWidth", "1");

                cols.addChild(col);
            }

            worksheet.addChild(cols);
        }

        // 构建工作表数据
        XmlNodeBuilder sheetData("sheetData");

        if (usedRange.isValid()) {
            // 遍历所有使用的行
            for (row_t row = usedRange.getStart().getRow(); row <= usedRange.getEnd().getRow(); ++row) {
                XmlNodeBuilder rowNode("row");
                rowNode.addAttribute("r", std::to_string(row.index()));

                bool hasData = false;
                // 遍历这一行的所有列
                for (column_t col = usedRange.getStart().getCol(); col <= usedRange.getEnd().getCol(); ++col) {
                    const TXCompactCell* cell = sheet->getCell(row, col);

                    if (cell && (!cell->isEmpty() || cell->getStyleIndex() != 0)) {
                        std::string cellRef = column_t::column_string_from_index(col.index()) + std::to_string(row.index());
                        XmlNodeBuilder cellNode = buildCellNode(cell, cellRef, context);
                        rowNode.addChild(cellNode);
                        hasData = true;
                    }
                }

                // 只添加非空行
                if (hasData) {
                    sheetData.addChild(rowNode);
                }
            }
        }

        worksheet.addChild(sheetData);

        // 添加工作表保护信息
        auto& protectionManager = sheet->getProtectionManager();
        if (protectionManager.isSheetProtected()) {
            const auto& protection = protectionManager.getSheetProtection();
            XmlNodeBuilder sheetProtection("sheetProtection");

            // 重要：添加sheet="1"属性表示工作表本身被保护
            sheetProtection.addAttribute("sheet", "1");

            // 添加现代Excel的SHA-512密码保护属性
            if (!protection.passwordHash.empty()) {
                // 现代Excel格式：使用algorithmName, hashValue, saltValue, spinCount
                sheetProtection.addAttribute("algorithmName", protection.algorithmName);
                sheetProtection.addAttribute("hashValue", protection.passwordHash);
                sheetProtection.addAttribute("saltValue", protection.saltValue);
                sheetProtection.addAttribute("spinCount", std::to_string(protection.spinCount));
            }

            // 添加保护选项属性（只有当值为false时才添加，因为默认值通常是true）
            if (!protection.selectLockedCells) {
                sheetProtection.addAttribute("selectLockedCells", "0");
            }
            if (!protection.selectUnlockedCells) {
                sheetProtection.addAttribute("selectUnlockedCells", "0");
            }
            if (!protection.formatCells) {
                sheetProtection.addAttribute("formatCells", "0");
            }
            if (!protection.formatColumns) {
                sheetProtection.addAttribute("formatColumns", "0");
            }
            if (!protection.formatRows) {
                sheetProtection.addAttribute("formatRows", "0");
            }
            if (!protection.insertColumns) {
                sheetProtection.addAttribute("insertColumns", "0");
            }
            if (!protection.insertRows) {
                sheetProtection.addAttribute("insertRows", "0");
            }
            if (!protection.deleteColumns) {
                sheetProtection.addAttribute("deleteColumns", "0");
            }
            if (!protection.deleteRows) {
                sheetProtection.addAttribute("deleteRows", "0");
            }

            worksheet.addChild(sheetProtection);
        }

        // 添加合并单元格（如果有）
        auto mergeRegions = sheet->getAllMergeRegions();
        if (!mergeRegions.empty()) {
            XmlNodeBuilder mergeCells("mergeCells");
            mergeCells.addAttribute("count", std::to_string(mergeRegions.size()));

            for (const auto& range : mergeRegions) {
                XmlNodeBuilder mergeCell("mergeCell");
                mergeCell.addAttribute("ref", range.toAddress());
                mergeCells.addChild(mergeCell);
            }

            worksheet.addChild(mergeCells);
        }

        // 添加数据验证（如果有）
        if (sheet->getDataValidationCount() > 0) {
            XmlNodeBuilder dataValidations = buildDataValidationsNode(sheet);
            worksheet.addChild(dataValidations);
        }

        // 添加自动筛选（如果有）
        if (sheet->hasAutoFilter()) {
            XmlNodeBuilder autoFilter = buildAutoFilterNode(sheet);
            worksheet.addChild(autoFilter);
        }

        // 添加透视表引用（如果有透视表）
        if (!m_pivotTables.empty()) {
            XmlNodeBuilder pivotTablesNode("pivotTables");
            pivotTablesNode.addAttribute("count", std::to_string(m_pivotTables.size()));

            for (size_t i = 0; i < m_pivotTables.size(); ++i) {
                XmlNodeBuilder pivotTable("pivotTable");
                pivotTable.addAttribute("cacheId", std::to_string(i + 1));
                pivotTable.addAttribute("name", "PivotTable" + std::to_string(i + 1));
                pivotTable.addAttribute("r:id", "rId" + std::to_string(i + 1));
                pivotTablesNode.addChild(pivotTable);
            }

            worksheet.addChild(pivotTablesNode);
        }

        // 添加绘图引用（如果有图表）
        if (sheet->getChartCount() > 0) {
            XmlNodeBuilder drawing("drawing");
            // 如果有透视表，图表的关系ID需要调整
            std::string drawingRId = m_pivotTables.empty() ? "rId1" : "rId" + std::to_string(m_pivotTables.size() + 1);
            drawing.addAttribute("r:id", drawingRId);
            worksheet.addChild(drawing);
        }

        TXXmlWriter writer;
        auto setRootResult = writer.setRootNode(worksheet);
        if (setRootResult.isError())
        {
            return Err<void>(setRootResult.error().getCode(), "Failed to set root node: " + setRootResult.error().getMessage());
        }

        auto xmlContentResult = writer.generateXmlString();
        if (xmlContentResult.isError())
        {
            return Err<void>(xmlContentResult.error().getCode(), "Failed to generate XML: " + xmlContentResult.error().getMessage());
        }

        std::vector<uint8_t> xmlData(xmlContentResult.value().begin(), xmlContentResult.value().end());
        auto writeResult = zipWriter.write(std::string(partName()), xmlData);
        if (writeResult.isError()) {
            return Err<void>(writeResult.error().getCode(), "Failed to write " + std::string(partName()) + ": " + writeResult.error().getMessage());
        }
        return Ok();
    }

    TXResult<void> TXWorksheetXmlHandler::saveWithStreamWriter(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) {
        const TXSheet* sheet = context.sheets[m_sheetIndex].get();
        TXRange usedRange = sheet->getUsedRange();

        // 估算单元格数量
        size_t estimatedCells = 0;
        if (usedRange.isValid()) {
            estimatedCells = (usedRange.getEnd().getRow().index() - usedRange.getStart().getRow().index() + 1) *
                           (usedRange.getEnd().getCol().index() - usedRange.getStart().getCol().index() + 1);
        }

        // 创建高性能流式写入器
        auto writer = TXWorksheetWriterFactory::createWriter(estimatedCells);

        // 开始写入工作表
        std::string rangeRef = usedRange.isValid() ? usedRange.toAddress() : "A1:A1";
        auto& rowColManager = sheet->getRowColumnManager();
        const auto& customColumnWidths = rowColManager.getCustomColumnWidths();
        bool hasCustomColumns = !customColumnWidths.empty();

        writer->startWorksheet(rangeRef, hasCustomColumns);

        // 写入列宽信息
        for (const auto& pair : customColumnWidths) {
            writer->writeColumnWidth(pair.first, pair.second);
        }

        // 开始写入工作表数据
        writer->startSheetData();

        if (usedRange.isValid()) {
            // 🚀 性能优化：使用迭代器直接遍历存在的单元格，避免大量哈希查找
            const auto& cellManager = sheet->getCellManager();

            // 按行分组单元格数据
            std::map<u32, std::vector<std::pair<u32, const TXCompactCell*>>> cellsByRow;

            // 只遍历实际存在的单元格
            for (auto it = cellManager.cbegin(); it != cellManager.cend(); ++it) {
                const auto& coord = it->first;
                const TXCompactCell* cell = &it->second;

                // 检查单元格是否在使用范围内
                if (coord.getRow() >= usedRange.getStart().getRow() &&
                    coord.getRow() <= usedRange.getEnd().getRow() &&
                    coord.getCol() >= usedRange.getStart().getCol() &&
                    coord.getCol() <= usedRange.getEnd().getCol()) {

                    // 只处理非空单元格或有样式的单元格
                    if (!cell->isEmpty() || cell->getStyleIndex() != 0) {
                        cellsByRow[coord.getRow().index()].emplace_back(coord.getCol().index(), cell);
                    }
                }
            }

            // 按行写入数据
            for (const auto& [rowIndex, rowCells] : cellsByRow) {
                writer->startRow(rowIndex);

                for (const auto& [colIndex, cell] : rowCells) {
                    std::string cellRef = column_t::column_string_from_index(colIndex) + std::to_string(rowIndex);
                    u32 styleIndex = cell->getStyleIndex();

                    // 根据单元格类型写入数据
                    TXCompactCell::CellType type = cell->getType();
                    if (type == TXCompactCell::CellType::String) {
                        const std::string& str = cell->getStringValue();
                        if (shouldUseInlineString(str)) {
                            writer->writeCellInlineString(cellRef, str, styleIndex);
                        } else {
                            u32 index = context.sharedStringsPool.add(str);
                            writer->writeCellSharedString(cellRef, index, styleIndex);
                        }
                    } else if (type == TXCompactCell::CellType::Number) {
                        writer->writeCellNumber(cellRef, cell->getNumberValue(), styleIndex);
                    } else if (type == TXCompactCell::CellType::Integer) {
                        writer->writeCellInteger(cellRef, cell->getIntegerValue(), styleIndex);
                    } else if (type == TXCompactCell::CellType::Boolean) {
                        writer->writeCellBoolean(cellRef, cell->getBooleanValue(), styleIndex);
                    }
                }

                writer->endRow();
            }
        }

        // 结束工作表数据和工作表
        writer->endSheetData();
        writer->endWorksheet();

        // 写入ZIP文件
        auto writeResult = writer->writeToZip(zipWriter, std::string(partName()));
        if (writeResult.isError()) {
            return writeResult;
        }

        return Ok();
    }

    bool TXWorksheetXmlHandler::shouldUseInlineString(const std::string& str) const
    {
        // 策略1: 极短字符串（1个字符）使用内联，节省共享字符串池空间
        if (str.length() <= 1) return true;
    
        // 策略2: 包含特殊XML字符的字符串使用内联（避免XML转义复杂性）
        if (str.find_first_of("<>&\"'") != std::string::npos) return true;
    
        // 策略3: 包含控制字符的字符串使用内联（避免XML解析问题）
        if (str.find_first_of("\n\r\t") != std::string::npos) return true;
        
        // 策略4: 非常长的字符串（>100字符）使用内联（避免共享字符串池膨胀）
        if (str.length() > 100) return true;
    
        // 策略5: 2-100字符的普通字符串使用共享字符串（最大化复用效果）
        return false;
    }

    XmlNodeBuilder TXWorksheetXmlHandler::buildCellNode(const TXCompactCell* cell, const std::string& cellRef,
                                                        const TXWorkbookContext& context) const
    {
        XmlNodeBuilder cellNode("c");
        cellNode.addAttribute("r", cellRef);

        if (!cell) return cellNode;

        // 处理样式
        if (u32 styleIndex = cell->getStyleIndex(); styleIndex != 0)
        {
            cellNode.addAttribute("s", std::to_string(styleIndex));
        }
        
        // 获取单元格值和类型
        const cell_value_t& value = cell->getValue();
        const TXCompactCell::CellType cellType = cell->getType();

        // 处理公式单元格
        if (cellType == TXCompactCell::CellType::Formula)
        {
            // 添加公式节点
            if (cell->isFormula()) {
                const TXFormula* formula = cell->getFormulaObject();
                if (formula) {
                    XmlNodeBuilder fNode("f");
                    std::string formulaStr = formula->getFormulaString();
                    // 确保公式不包含等号前缀
                    if (!formulaStr.empty() && formulaStr[0] == '=') {
                        formulaStr = formulaStr.substr(1);
                    }
                    fNode.setText(formulaStr);
                    cellNode.addChild(fNode);

                    // 如果有缓存的计算结果，也添加值节点
                    if (!std::holds_alternative<std::monostate>(value)) {
                        if (std::holds_alternative<double>(value)) {
                            XmlNodeBuilder vNode("v");
                            vNode.setText(TXNumberUtils::formatForExcelXml(std::get<double>(value)));
                            cellNode.addChild(vNode);
                        } else if (std::holds_alternative<int64_t>(value)) {
                            XmlNodeBuilder vNode("v");
                            vNode.setText(std::to_string(std::get<int64_t>(value)));
                            cellNode.addChild(vNode);
                        } else if (std::holds_alternative<std::string>(value)) {
                            // 公式结果是字符串
                            const std::string& str = std::get<std::string>(value);
                            if (shouldUseInlineString(str)) {
                                cellNode.addAttribute("t", "inlineStr");
                                XmlNodeBuilder isNode("is");
                                XmlNodeBuilder tNode("t");
                                tNode.setText(str);
                                isNode.addChild(tNode);
                                cellNode.addChild(isNode);
                            } else {
                                u32 index = context.sharedStringsPool.add(str);
                                cellNode.addAttribute("t", "s");
                                cellNode.addChild(XmlNodeBuilder("v").setText(std::to_string(index)));
                            }
                        }
                    }
                }
            }
        }
        else if (cellType == TXCompactCell::CellType::String)
        {
            const std::string& str = cell->getStringValue();
            
            // 根据字符串长度和内容决定使用内联还是共享
            if (shouldUseInlineString(str)) {
                // 内联字符串 - 直接嵌入XML
                cellNode.addAttribute("t", "inlineStr");
                XmlNodeBuilder isNode("is");
                XmlNodeBuilder tNode("t");
                tNode.setText(str);
                isNode.addChild(tNode);
                cellNode.addChild(isNode);
            } else {
                // 共享字符串 - 添加到池
                u32 index = context.sharedStringsPool.add(str);
            
                cellNode.addAttribute("t", "s");
                cellNode.addChild(XmlNodeBuilder("v").setText(std::to_string(index)));
            }
        }
        else if (std::holds_alternative<double>(value))
        {
            // 浮点数类型
            XmlNodeBuilder vNode("v");
            vNode.setText(TXNumberUtils::formatForExcelXml(std::get<double>(value)));
            cellNode.addChild(vNode);
        }
        else if (std::holds_alternative<int64_t>(value))
        {
            // 整数类型
            XmlNodeBuilder vNode("v");
            vNode.setText(std::to_string(std::get<int64_t>(value)));
            cellNode.addChild(vNode);
        }
        else if (std::holds_alternative<bool>(value))
        {
            // 布尔类型
            cellNode.addAttribute("t", "b");
            XmlNodeBuilder vNode("v");
            vNode.setText(std::get<bool>(value) ? "1" : "0");
            cellNode.addChild(vNode);
        }


        return cellNode;
    }

    XmlNodeBuilder TXWorksheetXmlHandler::buildDataValidationsNode(const TXSheet* sheet) const {
        XmlNodeBuilder dataValidations("dataValidations");

        const auto& validations = sheet->getDataValidations();
        dataValidations.addAttribute("count", std::to_string(validations.size()));

        for (const auto& pair : validations) {
            const TXRange& range = pair.first;
            const TXDataValidation& validation = pair.second;

            XmlNodeBuilder dataValidation("dataValidation");

            // 设置验证类型（必须在前面）
            switch (validation.getType()) {
                case DataValidationType::Whole:
                    dataValidation.addAttribute("type", "whole");
                    break;
                case DataValidationType::Decimal:
                    dataValidation.addAttribute("type", "decimal");
                    break;
                case DataValidationType::List:
                    dataValidation.addAttribute("type", "list");
                    break;
                case DataValidationType::Date:
                    dataValidation.addAttribute("type", "date");
                    break;
                case DataValidationType::Time:
                    dataValidation.addAttribute("type", "time");
                    break;
                case DataValidationType::TextLength:
                    dataValidation.addAttribute("type", "textLength");
                    break;
                case DataValidationType::Custom:
                    dataValidation.addAttribute("type", "custom");
                    break;
                default:
                    continue; // 跳过无效类型
            }

            // 设置允许空白（Excel默认行为）
            dataValidation.addAttribute("allowBlank", "1");

            // 设置范围
            dataValidation.addAttribute("sqref", range.toAddress());

            // 设置操作符（列表验证不需要操作符）
            if (validation.getType() != DataValidationType::List) {
                switch (validation.getOperator()) {
                    case DataValidationOperator::Between:
                        dataValidation.addAttribute("operator", "between");
                        break;
                    case DataValidationOperator::NotBetween:
                        dataValidation.addAttribute("operator", "notBetween");
                        break;
                    case DataValidationOperator::Equal:
                        dataValidation.addAttribute("operator", "equal");
                        break;
                    case DataValidationOperator::NotEqual:
                        dataValidation.addAttribute("operator", "notEqual");
                        break;
                    case DataValidationOperator::GreaterThan:
                        dataValidation.addAttribute("operator", "greaterThan");
                        break;
                    case DataValidationOperator::LessThan:
                        dataValidation.addAttribute("operator", "lessThan");
                        break;
                    case DataValidationOperator::GreaterThanOrEqual:
                        dataValidation.addAttribute("operator", "greaterThanOrEqual");
                        break;
                    case DataValidationOperator::LessThanOrEqual:
                        dataValidation.addAttribute("operator", "lessThanOrEqual");
                        break;
                }
            }

            // 对于列表验证，保持最简单的设置
            if (validation.getType() == DataValidationType::List) {
                // 不设置 showDropDown 属性，让Excel使用默认行为
                // 根据搜索结果，showDropDown=false 或不设置表示显示下拉箭头
            } else {
                // 非列表验证使用完整的属性设置

                // 设置下拉箭头显示
                // 注意：Excel中 showDropDown="1" 表示隐藏下拉箭头
                if (!validation.getShowDropDown()) {
                    dataValidation.addAttribute("showDropDown", "1");
                }

                // 设置错误警告
                if (validation.getShowErrorMessage()) {
                    dataValidation.addAttribute("showErrorMessage", "1");

                    // 设置错误样式
                    switch (validation.getErrorStyle()) {
                        case DataValidationErrorStyle::Stop:
                            dataValidation.addAttribute("errorStyle", "stop");
                            break;
                        case DataValidationErrorStyle::Warning:
                            dataValidation.addAttribute("errorStyle", "warning");
                            break;
                        case DataValidationErrorStyle::Information:
                            dataValidation.addAttribute("errorStyle", "information");
                            break;
                    }

                    // 添加错误标题和消息
                    if (!validation.getErrorTitle().empty()) {
                        dataValidation.addAttribute("errorTitle", validation.getErrorTitle());
                    }
                    if (!validation.getErrorMessage().empty()) {
                        dataValidation.addAttribute("error", validation.getErrorMessage());
                    }
                }

                // 设置输入提示
                if (validation.getShowInputMessage()) {
                    dataValidation.addAttribute("showInputMessage", "1");

                    if (!validation.getPromptTitle().empty()) {
                        dataValidation.addAttribute("promptTitle", validation.getPromptTitle());
                    }
                    if (!validation.getPromptMessage().empty()) {
                        dataValidation.addAttribute("prompt", validation.getPromptMessage());
                    }
                }
            }

            // 添加公式节点
            if (!validation.getFormula1().empty()) {
                XmlNodeBuilder formula1("formula1");
                formula1.setText(validation.getFormula1());
                dataValidation.addChild(formula1);
            }

            if (!validation.getFormula2().empty()) {
                XmlNodeBuilder formula2("formula2");
                formula2.setText(validation.getFormula2());
                dataValidation.addChild(formula2);
            }

            dataValidations.addChild(dataValidation);
        }

        return dataValidations;
    }

    XmlNodeBuilder TXWorksheetXmlHandler::buildAutoFilterNode(const TXSheet* sheet) const {
        const TXAutoFilter* autoFilter = sheet->getAutoFilter();
        if (!autoFilter) {
            return XmlNodeBuilder("autoFilter");  // 返回空节点
        }

        XmlNodeBuilder autoFilterNode("autoFilter");

        // 设置筛选范围
        autoFilterNode.addAttribute("ref", autoFilter->getRange().toAbsoluteAddress());

        // 添加筛选条件（如果有）
        const auto& conditions = autoFilter->getFilterConditions();

        // 按列分组筛选条件
        std::map<u32, std::vector<FilterCondition>> conditionsByColumn;
        for (const auto& condition : conditions) {
            conditionsByColumn[condition.columnIndex].push_back(condition);
        }

        // 为每列生成筛选XML
        for (const auto& [columnIndex, columnConditions] : conditionsByColumn) {
            XmlNodeBuilder filterColumn("filterColumn");
            filterColumn.addAttribute("colId", std::to_string(columnIndex));

            if (columnConditions.size() == 1) {
                // 单个条件
                const auto& condition = columnConditions[0];

                switch (condition.operator_) {
                    case FilterOperator::Equal:
                    case FilterOperator::NotEqual:
                    case FilterOperator::GreaterThan:
                    case FilterOperator::LessThan:
                    case FilterOperator::GreaterThanOrEqual:
                    case FilterOperator::LessThanOrEqual:
                    case FilterOperator::Contains:
                    case FilterOperator::NotContains:
                    case FilterOperator::BeginsWith:
                    case FilterOperator::EndsWith: {
                        XmlNodeBuilder customFilters("customFilters");
                        XmlNodeBuilder customFilter("customFilter");

                        // 设置操作符
                        std::string op;
                        switch (condition.operator_) {
                            case FilterOperator::Equal: op = "equal"; break;
                            case FilterOperator::NotEqual: op = "notEqual"; break;
                            case FilterOperator::GreaterThan: op = "greaterThan"; break;
                            case FilterOperator::LessThan: op = "lessThan"; break;
                            case FilterOperator::GreaterThanOrEqual: op = "greaterThanOrEqual"; break;
                            case FilterOperator::LessThanOrEqual: op = "lessThanOrEqual"; break;
                            default: op = "equal"; break;
                        }
                        customFilter.addAttribute("operator", op);
                        customFilter.addAttribute("val", condition.value1);

                        customFilters.addChild(customFilter);
                        filterColumn.addChild(customFilters);
                        break;
                    }
                    case FilterOperator::Top10:
                    case FilterOperator::Bottom10: {
                        XmlNodeBuilder top10("top10");
                        top10.addAttribute("top", condition.operator_ == FilterOperator::Top10 ? "1" : "0");
                        top10.addAttribute("val", condition.value1);
                        filterColumn.addChild(top10);
                        break;
                    }
                    default:
                        // 其他类型暂不支持
                        continue;
                }
            } else if (columnConditions.size() == 2) {
                // 多个条件（通常是范围筛选）
                XmlNodeBuilder customFilters("customFilters");
                customFilters.addAttribute("and", "1");  // AND逻辑

                for (const auto& condition : columnConditions) {
                    XmlNodeBuilder customFilter("customFilter");

                    // 设置操作符
                    std::string op;
                    switch (condition.operator_) {
                        case FilterOperator::Equal: op = "equal"; break;
                        case FilterOperator::NotEqual: op = "notEqual"; break;
                        case FilterOperator::GreaterThan: op = "greaterThan"; break;
                        case FilterOperator::LessThan: op = "lessThan"; break;
                        case FilterOperator::GreaterThanOrEqual: op = "greaterThanOrEqual"; break;
                        case FilterOperator::LessThanOrEqual: op = "lessThanOrEqual"; break;
                        default: op = "equal"; break;
                    }
                    customFilter.addAttribute("operator", op);
                    customFilter.addAttribute("val", condition.value1);

                    customFilters.addChild(customFilter);
                }

                filterColumn.addChild(customFilters);
            }

            autoFilterNode.addChild(filterColumn);
        }

        return autoFilterNode;
    }

    std::vector<std::shared_ptr<TXPivotTable>> TXWorksheetXmlHandler::getPivotTablesForSheet(const std::string& sheetName, const TXWorkbookContext& context) const {
        // 这里需要从工作簿上下文中获取透视表信息
        // 由于当前的TXWorkbookContext没有透视表信息，我们需要通过其他方式获取
        // 暂时返回空列表，这个方法需要在TXWorkbookContext中添加透视表支持后才能正确实现
        return {};
    }

    void TXWorksheetXmlHandler::setPivotTables(const std::vector<std::shared_ptr<TXPivotTable>>& pivotTables) {
        m_pivotTables = pivotTables;
    }

    // ==================== 🚀 批处理方法实现 ====================

    void TXWorksheetXmlHandler::initializeBatchGenerator() {
        // 初始化内存管理器
        TXUnifiedMemoryManager::Config memory_config;
        memory_config.chunk_size = 32 * 1024 * 1024;      // 32MB块
        memory_config.memory_limit = 512 * 1024 * 1024;   // 512MB限制
        memory_config.enable_monitoring = true;

        memory_manager_ = std::make_unique<TXUnifiedMemoryManager>(memory_config);

        // 初始化批处理XML生成器
        TXBatchXMLGenerator::XMLGeneratorConfig xml_config;
        xml_config.enable_memory_pooling = true;
        xml_config.enable_parallel_generation = true;
        xml_config.batch_size = 5000;
        xml_config.pretty_print = false;  // 生产环境不需要格式化

        batch_generator_ = std::make_unique<TXBatchXMLGenerator>(*memory_manager_, xml_config);
        batch_generator_->loadDefaultTemplates();
    }

    TXResult<void> TXWorksheetXmlHandler::saveWithBatchProcessor(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) {
        // ❌ 暂时禁用批处理器，因为它缺少关键的Excel XML格式支持
        // 需要修复以下问题：
        // 1. 缺少单元格引用（r属性，如A1, B2等）
        // 2. 缺少维度信息（dimension元素）
        // 3. 缺少字符串类型的正确处理（inlineStr vs 共享字符串）
        // 4. 缺少完整的Excel XML结构

        return Err<void>(TXErrorCode::Unknown,
                       "Batch processor needs Excel XML format fixes - falling back to DOM method");
    }

}
