//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TXXmlHandler.hpp"
#include "TXXmlReader.hpp"
#include "TXXmlWriter.hpp"
#include "TXSheet.hpp"
#include "TXCell.hpp"
#include "TXRange.hpp"
#include "TXTypes.hpp"
#include <sstream>
#include <iomanip>

namespace TinaXlsx
{
    class TXWorksheetXmlHandler : public TXXmlHandler
    {
    public:
        explicit TXWorksheetXmlHandler(u64 sheetIndex): m_sheetIndex(sheetIndex)
        {
        }

        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override
        {
            auto xmlData = zipReader.read(partName());
            if (xmlData.isError())
            {
                return Err<void>(xmlData.error().getCode(), "Failed to read " + partName());
            }
            const std::vector<uint8_t>& fileBytes = xmlData.value(); // Get the actual std::vector<uint8_t>

            std::string xmlContent(fileBytes.begin(), fileBytes.end());
            TXXmlReader reader;
            auto parseResult = reader.parseFromString(xmlContent);
            if (parseResult.isError())
            {
                return Err<void>(parseResult.error().getCode(), "Failed to parse worksheet.xml: " + parseResult.error().getMessage());
            }
            
            // 解析 sheetData 节点，填充 context.sheets[sheetIndex_]
            auto cellNodesResult = reader.findNodes("//sheetData/row/c");
            if (cellNodesResult.isError())
            {
                return Err<void>(cellNodesResult.error().getCode(), "Failed to find cell nodes: " + cellNodesResult.error().getMessage());
            }
            
            for (const auto& cellNode : cellNodesResult.value())
            {
                auto refIter = cellNode.attributes.find("r");
                if (refIter != cellNode.attributes.end())
                {
                    std::string ref = refIter->second;
                    std::string value = cellNode.value;
                    context.sheets[m_sheetIndex]->setCellValue(ref, value);
                }
            }
            return Ok();
        }

        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override
        {
            const TXSheet* sheet = context.sheets[m_sheetIndex].get();
            XmlNodeBuilder worksheet("worksheet");
            worksheet.addAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                     .addAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            // 添加维度信息
            XmlNodeBuilder dimension("dimension");
            TXRange usedRange = sheet->getUsedRange();
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
                        const TXCell* cell = sheet->getCell(row, col);
                        
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

                // 添加密码哈希（如果有）
                if (!protection.passwordHash.empty()) {
                    sheetProtection.addAttribute("password", protection.passwordHash);
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

    private:
        bool shouldUseInlineString(const std::string& str) const;
        /**
         * @brief 构建单个单元格节点
         * @param cell 单元格对象
         * @param cellRef 单元格引用（如A1）
         * @param context 工作簿上下文
         * @return 单元格节点
         */
        XmlNodeBuilder buildCellNode(const TXCell* cell, const std::string& cellRef,const TXWorkbookContext& context) const;

        u64 m_sheetIndex;
    };
}
