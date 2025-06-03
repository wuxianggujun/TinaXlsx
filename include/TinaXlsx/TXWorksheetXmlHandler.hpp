//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TXXmlHandler.hpp"
#include "TXXmlReader.hpp"
#include "TXXmlWriter.hpp"
#include "TXSheet.hpp"
#include "TXCompactCell.hpp"
#include "TXRange.hpp"
#include "TXTypes.hpp"
#include "TXStreamXmlReader.hpp"
#include "TXSIMDXmlParser.hpp"
#include <sstream>
#include <iomanip>
#include <memory>
#include <vector>

namespace TinaXlsx
{
    // 前向声明
    class TXPivotTable;

    class TXWorksheetXmlHandler : public TXXmlHandler
    {
    public:
        explicit TXWorksheetXmlHandler(u64 sheetIndex): m_sheetIndex(sheetIndex)
        {
        }

        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override
        {
            // 🚀 性能优化：使用SIMD优化的XML解析器
            if (m_sheetIndex >= context.sheets.size()) {
                return Err<void>(TXErrorCode::InvalidArgument, "Sheet index out of range");
            }

            // 读取XML数据
            auto xmlData = zipReader.read(partName());
            if (xmlData.isError()) {
                return Err<void>(xmlData.error().getCode(), "Failed to read " + partName());
            }

            const std::vector<uint8_t>& fileBytes = xmlData.value();
            std::string xmlContent(fileBytes.begin(), fileBytes.end());

            // 🚀 使用SIMD优化的解析器
            TXSIMDWorksheetParser parser(context.sheets[m_sheetIndex].get());
            size_t cellCount = parser.parse(xmlContent);

            // 输出统计信息（调试用）
            const auto& stats = parser.getStats();
            // TODO: 添加日志系统后输出统计信息
            // printf("SIMD解析: %zu行, %zu单元格, %.2fms\n",
            //        stats.totalRows, stats.totalCells, stats.parseTimeMs);

            return Ok();
        }

        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override
        {
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
                            XmlNodeBuilder cellNode = buildCellNode(cell, cellRef,context);
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

        [[nodiscard]] std::string partName() const override {
            return "xl/worksheets/sheet" + std::to_string(m_sheetIndex + 1) + ".xml";
        }

        /**
         * @brief 设置透视表信息（用于传递透视表数据）
         * @param pivotTables 透视表列表
         */
        void setPivotTables(const std::vector<std::shared_ptr<class TXPivotTable>>& pivotTables);

    private:
        bool shouldUseInlineString(const std::string& str) const;

        /**
         * @brief 使用流式写入器保存（高性能版本）
         * @param zipWriter ZIP写入器
         * @param context 工作簿上下文
         * @return 保存结果
         */
        TXResult<void> saveWithStreamWriter(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context);

        /**
         * @brief 构建单个单元格节点
         * @param cell 单元格对象
         * @param cellRef 单元格引用（如A1）
         * @param context 工作簿上下文
         * @return 单元格节点
         */
        XmlNodeBuilder buildCellNode(const TXCompactCell* cell, const std::string& cellRef,const TXWorkbookContext& context) const;

        /**
         * @brief 构建数据验证节点
         * @param sheet 工作表对象
         * @return 数据验证节点
         */
        XmlNodeBuilder buildDataValidationsNode(const TXSheet* sheet) const;

        /**
         * @brief 构建自动筛选节点
         * @param sheet 工作表对象
         * @return 自动筛选节点
         */
        XmlNodeBuilder buildAutoFilterNode(const TXSheet* sheet) const;

        /**
         * @brief 获取工作表的透视表列表
         * @param sheetName 工作表名称
         * @param context 工作簿上下文
         * @return 透视表列表
         */
        std::vector<std::shared_ptr<class TXPivotTable>> getPivotTablesForSheet(const std::string& sheetName, const TXWorkbookContext& context) const;

    private:
        std::vector<std::shared_ptr<class TXPivotTable>> m_pivotTables;  ///< 当前工作表的透视表列表

        u64 m_sheetIndex;
    };
}
