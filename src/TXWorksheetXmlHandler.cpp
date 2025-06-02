//
// Created by wuxianggujun on 2025/5/29.
//

#include "TinaXlsx/TXWorksheetXmlHandler.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXNumberUtils.hpp"
#include <variant>
#include <sstream>

#include "TinaXlsx/TXSharedStringsPool.hpp"

namespace TinaXlsx
{

    TXResult<void> TXWorksheetXmlHandler::saveWithStreamWriter(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) {
        const TXSheet* sheet = context.sheets[m_sheetIndex].get();
        TXRange usedRange = sheet->getUsedRange();

        // 使用字符串流进行高效XML生成
        std::ostringstream xmlStream;
        // 注意：ostringstream没有reserve方法，但我们可以通过str()预分配
        std::string buffer;
        buffer.reserve(1024 * 1024); // 预分配1MB缓冲区

        // XML声明和工作表开始
        xmlStream << "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
        xmlStream << "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                     "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n";

        // 维度信息
        std::string rangeRef = usedRange.isValid() ? usedRange.toAddress() : "A1:A1";
        xmlStream << "<dimension ref=\"" << rangeRef << "\"/>\n";

        // 列宽信息（简化处理）
        auto& rowColManager = sheet->getRowColumnManager();
        const auto& customColumnWidths = rowColManager.getCustomColumnWidths();
        if (!customColumnWidths.empty()) {
            xmlStream << "<cols>\n";
            for (const auto& pair : customColumnWidths) {
                // pair.first 是 column_t::index_t 类型，直接使用
                xmlStream << "<col min=\"" << pair.first << "\" max=\"" << pair.first
                         << "\" width=\"" << pair.second << "\" customWidth=\"1\"/>\n";
            }
            xmlStream << "</cols>\n";
        }

        // 开始sheetData
        xmlStream << "<sheetData>\n";

        if (usedRange.isValid()) {
            // 按行流式写入单元格数据
            for (row_t row = usedRange.getStart().getRow(); row <= usedRange.getEnd().getRow(); ++row) {
                bool hasRowData = false;
                std::ostringstream rowStream;

                // 先检查这一行是否有数据，并构建行内容
                for (column_t col = usedRange.getStart().getCol(); col <= usedRange.getEnd().getCol(); ++col) {
                    const TXCell* cell = sheet->getCell(row, col);

                    if (cell && (!cell->isEmpty() || cell->getStyleIndex() != 0)) {
                        if (!hasRowData) {
                            hasRowData = true;
                            rowStream << "<row r=\"" << row.index() << "\">\n";
                        }

                        std::string cellRef = column_t::column_string_from_index(col.index()) + std::to_string(row.index());
                        rowStream << "<c r=\"" << cellRef << "\"";

                        // 样式索引
                        if (cell->getStyleIndex() != 0) {
                            rowStream << " s=\"" << cell->getStyleIndex() << "\"";
                        }

                        // 单元格类型和值
                        TXCell::CellType type = cell->getType();
                        if (type == TXCell::CellType::String) {
                            const std::string& str = cell->getStringValue();
                            if (shouldUseInlineString(str)) {
                                rowStream << " t=\"inlineStr\"><is><t>" << str << "</t></is></c>\n";
                            } else {
                                u32 index = context.sharedStringsPool.add(str);
                                rowStream << " t=\"s\"><v>" << index << "</v></c>\n";
                            }
                        } else if (type == TXCell::CellType::Number) {
                            rowStream << "><v>" << cell->getNumberValue() << "</v></c>\n";
                        } else if (type == TXCell::CellType::Integer) {
                            rowStream << "><v>" << cell->getIntegerValue() << "</v></c>\n";
                        } else if (type == TXCell::CellType::Boolean) {
                            rowStream << " t=\"b\"><v>" << (cell->getBooleanValue() ? "1" : "0") << "</v></c>\n";
                        } else {
                            rowStream << "/>\n"; // 空单元格
                        }
                    }
                }

                if (hasRowData) {
                    rowStream << "</row>\n";
                    xmlStream << rowStream.str();
                }
            }
        }

        // 结束sheetData和worksheet
        xmlStream << "</sheetData>\n";
        xmlStream << "</worksheet>\n";

        // 写入ZIP文件
        std::string xmlContent = xmlStream.str();
        std::vector<uint8_t> xmlData(xmlContent.begin(), xmlContent.end());

        auto writeResult = zipWriter.write(std::string(partName()), xmlData);
        if (writeResult.isError()) {
            return Err<void>(writeResult.error().getCode(),
                           "Failed to write " + std::string(partName()) + ": " + writeResult.error().getMessage());
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

    XmlNodeBuilder TXWorksheetXmlHandler::buildCellNode(const TXCell* cell, const std::string& cellRef,
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
        const TXCell::CellType cellType = cell->getType();

        // 处理公式单元格
        if (cellType == TXCell::CellType::Formula)
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
        else if (cellType == TXCell::CellType::String)
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
}
