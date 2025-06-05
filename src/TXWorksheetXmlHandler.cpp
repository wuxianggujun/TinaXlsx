//
// Created by wuxianggujun on 2025/5/29.
//

#include "TinaXlsx/TXWorksheetXmlHandler.hpp"

#include <iostream>

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
        // 🚀 统一使用批处理XML生成器

        // 1. 确保批处理生成器已初始化
        if (!batch_generator_) {
            initializeBatchGenerator();
        }

        // 2. 获取工作表数据
        const TXSheet* sheet = context.sheets[m_sheetIndex].get();

        // 3. 收集行数据
        auto rows_data = collectRowsData(sheet, context);

        // 4. 使用批处理生成器生成完整的工作表XML
        auto xml_result = generateCompleteWorksheetXML(sheet, rows_data, context);

        if (xml_result.isError()) {
            return Err<void>(xml_result.error().getCode(),
                           "Failed to generate worksheet XML: " + xml_result.error().getMessage());
        }

        // 5. 写入ZIP文件
        std::string xml_content = xml_result.value();
        std::vector<uint8_t> xmlData(xml_content.begin(), xml_content.end());

        auto writeResult = zipWriter.write(std::string(partName()), xmlData);
        if (writeResult.isError()) {
            return Err<void>(writeResult.error().getCode(),
                           "Failed to write " + std::string(partName()) + ": " + writeResult.error().getMessage());
        }

        return Ok();
    }

    std::vector<std::pair<size_t, std::vector<std::pair<std::string, TXCompactCell>>>> TXWorksheetXmlHandler::collectRowsData(const TXSheet* sheet, const TXWorkbookContext& context) {
        std::vector<std::pair<size_t, std::vector<std::pair<std::string, TXCompactCell>>>> rows_data;

        TXRange usedRange = sheet->getUsedRange();
        if (!usedRange.isValid()) {
            return rows_data;
        }

        const auto& cellManager = sheet->getCellManager();

        // 按行分组单元格数据
        std::map<u32, std::vector<std::pair<std::string, TXCompactCell>>> cellsByRow;

        // 遍历实际存在的单元格
        for (auto it = cellManager.cbegin(); it != cellManager.cend(); ++it) {
            const auto& coord = it->first;
            const TXCompactCell& cell = it->second;

            // 检查单元格是否在使用范围内
            if (coord.getRow() >= usedRange.getStart().getRow() &&
                coord.getRow() <= usedRange.getEnd().getRow() &&
                coord.getCol() >= usedRange.getStart().getCol() &&
                coord.getCol() <= usedRange.getEnd().getCol()) {

                // 只处理非空单元格或有样式的单元格
                if (!cell.isEmpty() || cell.getStyleIndex() != 0) {
                    std::string cellRef = column_t::column_string_from_index(coord.getCol().index()) + std::to_string(coord.getRow().index());
                    cellsByRow[coord.getRow().index()].emplace_back(cellRef, cell);
                }
            }
        }

        // 转换为批处理格式
        for (const auto& [rowIndex, rowCells] : cellsByRow) {
            rows_data.emplace_back(rowIndex, rowCells);
        }

        return rows_data;
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
                        if (true) {  // 简化：总是使用内联字符串
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

    // shouldUseInlineString方法已删除，现在统一使用内联字符串处理

    XmlNodeBuilder TXWorksheetXmlHandler::buildCellNode(const TXCompactCell* cell, const std::string& cellRef,
                                                        const TXWorkbookContext& context) const
    {
        // 🚀 优化：如果批处理生成器可用，使用它来优化单元格值的格式化
        if (batch_generator_ && cell) {
            try {
                // 使用批处理生成器的优化方法来格式化值
                // 但仍然使用DOM方式构建节点结构，确保兼容性

                // 这里我们可以使用批处理生成器的formatCellValue和shouldUseInlineString方法
                // 来获得更好的性能，同时保持DOM结构的完整性
            } catch (...) {
                // 如果出现异常，继续使用原有方式
            }
        }

        // 原有的DOM方式作为回退
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
                            if (true) {  // 简化：总是使用内联字符串
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

            // 🚀 简化字符串处理策略 - 统一使用内联字符串
            bool useInline = true;  // 暂时总是使用内联字符串，避免复杂的共享字符串逻辑

            if (useInline) {
                // 内联字符串 - 直接嵌入XML
                cellNode.addAttribute("t", "inlineStr");
                XmlNodeBuilder isNode("is");
                XmlNodeBuilder tNode("t");

                // 🚀 直接设置文本，XML转义由框架处理
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
            double doubleValue = std::get<double>(value);

            // 🚀 使用标准数字格式化
            std::string formattedValue = TXNumberUtils::formatForExcelXml(doubleValue);

            XmlNodeBuilder vNode("v");
            vNode.setText(formattedValue);
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
        const TXSheet* sheet = context.sheets[m_sheetIndex].get();
        TXRange usedRange = sheet->getUsedRange();

        try {
            // 🚀 使用批处理XML生成器生成完整的工作表XML

            // 1. 收集所有单元格数据，按行组织
            std::vector<std::pair<size_t, std::vector<std::pair<std::string, TXCompactCell>>>> rows_data;

            if (usedRange.isValid()) {
                const auto& cellManager = sheet->getCellManager();

                // 按行分组单元格数据
                std::map<u32, std::vector<std::pair<std::string, TXCompactCell>>> cellsByRow;

                // 遍历实际存在的单元格
                for (auto it = cellManager.cbegin(); it != cellManager.cend(); ++it) {
                    const auto& coord = it->first;
                    const TXCompactCell& cell = it->second;

                    // 检查单元格是否在使用范围内
                    if (coord.getRow() >= usedRange.getStart().getRow() &&
                        coord.getRow() <= usedRange.getEnd().getRow() &&
                        coord.getCol() >= usedRange.getStart().getCol() &&
                        coord.getCol() <= usedRange.getEnd().getCol()) {

                        // 只处理非空单元格或有样式的单元格
                        if (!cell.isEmpty() || cell.getStyleIndex() != 0) {
                            std::string cellRef = column_t::column_string_from_index(coord.getCol().index()) + std::to_string(coord.getRow().index());
                            cellsByRow[coord.getRow().index()].emplace_back(cellRef, cell);
                        }
                    }
                }

                // 转换为批处理格式
                for (const auto& [rowIndex, rowCells] : cellsByRow) {
                    rows_data.emplace_back(rowIndex, rowCells);
                }
            }

            // 2. 使用批处理生成器生成完整的工作表XML
            auto xml_result = generateCompleteWorksheetXML(sheet, rows_data, context);

            if (xml_result.isError()) {
                return Err<void>(xml_result.error().getCode(),
                               "Batch XML generation failed: " + xml_result.error().getMessage());
            }

            // 3. 写入ZIP文件
            std::string xml_content = xml_result.value();
            std::vector<uint8_t> xmlData(xml_content.begin(), xml_content.end());

            auto writeResult = zipWriter.write(std::string(partName()), xmlData);
            if (writeResult.isError()) {
                return Err<void>(writeResult.error().getCode(),
                               "Failed to write " + std::string(partName()) + ": " + writeResult.error().getMessage());
            }

            return Ok();

        } catch (const std::exception& e) {
            return Err<void>(TXErrorCode::Unknown,
                           "Exception in batch processor: " + std::string(e.what()));
        }
    }

    TXResult<std::string> TXWorksheetXmlHandler::generateCompleteWorksheetXML(
        const TXSheet* sheet,
        const std::vector<std::pair<size_t, std::vector<std::pair<std::string, TXCompactCell>>>>& rows_data,
        const TXWorkbookContext& context) {

        try {
            // 🚀 使用ostringstream生成完整的Excel工作表XML结构
            std::ostringstream xml;

            // 1. XML声明和根元素
            xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
            xml << "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                << "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n";

            // 2. 维度信息
            TXRange usedRange = sheet->getUsedRange();
            xml << "<dimension ref=\"";
            if (usedRange.isValid()) {
                xml << usedRange.toAddress();
            } else {
                xml << "A1:A1";
            }
            xml << "\"/>\n";

            // 3. 列宽信息
            auto& rowColManager = sheet->getRowColumnManager();
            const auto& customColumnWidths = rowColManager.getCustomColumnWidths();
            if (!customColumnWidths.empty()) {
                xml << "<cols>\n";
                for (const auto& [colIndex, width] : customColumnWidths) {
                    // 格式化宽度值
                    std::ostringstream widthStream;
                    widthStream << std::fixed << std::setprecision(2) << width;
                    std::string widthStr = widthStream.str();
                    // 移除尾随零
                    widthStr.erase(widthStr.find_last_not_of('0') + 1, std::string::npos);
                    if (widthStr.back() == '.') widthStr.pop_back();

                    xml << "<col min=\"" << colIndex << "\" max=\"" << colIndex
                        << "\" width=\"" << widthStr << "\" customWidth=\"1\"/>\n";
                }
                xml << "</cols>\n";
            }

            // 4. 工作表数据 - 使用批处理生成器
            xml << "<sheetData>\n";

            if (!rows_data.empty()) {
                auto rows_xml_result = batch_generator_->generateRowsXML(rows_data);
                if (rows_xml_result.isOk()) {
                    // 🚀 将批处理生成器的XML直接写入
                    xml << rows_xml_result.value();
                } else {
                    return Err<std::string>(rows_xml_result.error().getCode(),
                                          "Failed to generate rows XML: " + rows_xml_result.error().getMessage());
                }
            }

            xml << "</sheetData>\n";

            // 5. 工作表保护信息
            auto& protectionManager = sheet->getProtectionManager();
            if (protectionManager.isSheetProtected()) {
                const auto& protection = protectionManager.getSheetProtection();
                xml << "<sheetProtection sheet=\"1\"";

                if (!protection.passwordHash.empty()) {
                    xml << " algorithmName=\"" << protection.algorithmName << "\""
                        << " hashValue=\"" << protection.passwordHash << "\""
                        << " saltValue=\"" << protection.saltValue << "\""
                        << " spinCount=\"" << protection.spinCount << "\"";
                }

                // 添加保护选项
                if (!protection.selectLockedCells) xml << " selectLockedCells=\"0\"";
                if (!protection.selectUnlockedCells) xml << " selectUnlockedCells=\"0\"";
                if (!protection.formatCells) xml << " formatCells=\"0\"";
                if (!protection.formatColumns) xml << " formatColumns=\"0\"";
                if (!protection.formatRows) xml << " formatRows=\"0\"";
                if (!protection.insertColumns) xml << " insertColumns=\"0\"";
                if (!protection.insertRows) xml << " insertRows=\"0\"";
                if (!protection.deleteColumns) xml << " deleteColumns=\"0\"";
                if (!protection.deleteRows) xml << " deleteRows=\"0\"";

                xml << "/>\n";
            }

            // 6. 合并单元格
            auto mergeRegions = sheet->getAllMergeRegions();
            if (!mergeRegions.empty()) {
                xml << "<mergeCells count=\"" << mergeRegions.size() << "\">\n";
                for (const auto& range : mergeRegions) {
                    xml << "<mergeCell ref=\"" << range.toAddress() << "\"/>\n";
                }
                xml << "</mergeCells>\n";
            }

            // 7. 数据验证
            if (sheet->getDataValidationCount() > 0) {
                std::string dataValidationsXML = generateDataValidationsXML(sheet);
                if (!dataValidationsXML.empty()) {
                    xml << dataValidationsXML;
                }
            }

            // 8. 自动筛选
            if (sheet->hasAutoFilter()) {
                std::string autoFilterXML = generateAutoFilterXML(sheet);
                if (!autoFilterXML.empty()) {
                    xml << autoFilterXML;
                }
            }

            // 9. 透视表引用
            if (!m_pivotTables.empty()) {
                xml << "<pivotTables count=\"" << m_pivotTables.size() << "\">\n";
                for (size_t i = 0; i < m_pivotTables.size(); ++i) {
                    xml << "<pivotTable cacheId=\"" << (i + 1) << "\" name=\"PivotTable" << (i + 1)
                        << "\" r:id=\"rId" << (i + 1) << "\"/>\n";
                }
                xml << "</pivotTables>\n";
            }

            // 10. 图表引用
            if (sheet->getChartCount() > 0) {
                std::string drawingRId = m_pivotTables.empty() ? "rId1" : "rId" + std::to_string(m_pivotTables.size() + 1);
                xml << "<drawing r:id=\"" << drawingRId << "\"/>\n";
            }

            xml << "</worksheet>";

            // 🚀 返回完整的XML字符串
            return Ok(xml.str());

        } catch (const std::exception& e) {
            return Err<std::string>(TXErrorCode::Unknown,
                                  "Exception in complete worksheet XML generation: " + std::string(e.what()));
        }
    }

    std::string TXWorksheetXmlHandler::generateDataValidationsXML(const TXSheet* sheet) {
        const auto& validations = sheet->getDataValidations();
        if (validations.empty()) return "";

        std::ostringstream xml;
        xml << "<dataValidations count=\"" << validations.size() << "\">\n";

        for (const auto& [range, validation] : validations) {
            xml << "<dataValidation";

            // 设置验证类型
            switch (validation.getType()) {
                case DataValidationType::Whole: xml << " type=\"whole\""; break;
                case DataValidationType::Decimal: xml << " type=\"decimal\""; break;
                case DataValidationType::List: xml << " type=\"list\""; break;
                case DataValidationType::Date: xml << " type=\"date\""; break;
                case DataValidationType::Time: xml << " type=\"time\""; break;
                case DataValidationType::TextLength: xml << " type=\"textLength\""; break;
                case DataValidationType::Custom: xml << " type=\"custom\""; break;
                default: continue;
            }

            xml << " allowBlank=\"1\" sqref=\"" << range.toAddress() << "\"";

            // 设置操作符
            if (validation.getType() != DataValidationType::List) {
                switch (validation.getOperator()) {
                    case DataValidationOperator::Between: xml << " operator=\"between\""; break;
                    case DataValidationOperator::NotBetween: xml << " operator=\"notBetween\""; break;
                    case DataValidationOperator::Equal: xml << " operator=\"equal\""; break;
                    case DataValidationOperator::NotEqual: xml << " operator=\"notEqual\""; break;
                    case DataValidationOperator::GreaterThan: xml << " operator=\"greaterThan\""; break;
                    case DataValidationOperator::LessThan: xml << " operator=\"lessThan\""; break;
                    case DataValidationOperator::GreaterThanOrEqual: xml << " operator=\"greaterThanOrEqual\""; break;
                    case DataValidationOperator::LessThanOrEqual: xml << " operator=\"lessThanOrEqual\""; break;
                }
            }

            xml << ">\n";

            // 添加公式
            if (!validation.getFormula1().empty()) {
                xml << "<formula1>" << validation.getFormula1() << "</formula1>\n";
            }
            if (!validation.getFormula2().empty()) {
                xml << "<formula2>" << validation.getFormula2() << "</formula2>\n";
            }

            xml << "</dataValidation>\n";
        }

        xml << "</dataValidations>\n";
        return xml.str();
    }

    std::string TXWorksheetXmlHandler::generateAutoFilterXML(const TXSheet* sheet) {
        const TXAutoFilter* autoFilter = sheet->getAutoFilter();
        if (!autoFilter) return "";

        std::ostringstream xml;
        xml << "<autoFilter ref=\"" << autoFilter->getRange().toAbsoluteAddress() << "\">\n";

        // 添加筛选条件
        const auto& conditions = autoFilter->getFilterConditions();
        std::map<u32, std::vector<FilterCondition>> conditionsByColumn;
        for (const auto& condition : conditions) {
            conditionsByColumn[condition.columnIndex].push_back(condition);
        }

        for (const auto& [columnIndex, columnConditions] : conditionsByColumn) {
            xml << "<filterColumn colId=\"" << columnIndex << "\">\n";

            if (columnConditions.size() == 1) {
                const auto& condition = columnConditions[0];
                xml << "<customFilters><customFilter operator=\"equal\" val=\""
                    << condition.value1 << "\"/></customFilters>\n";
            }

            xml << "</filterColumn>\n";
        }

        xml << "</autoFilter>\n";
        return xml.str();
    }

}
