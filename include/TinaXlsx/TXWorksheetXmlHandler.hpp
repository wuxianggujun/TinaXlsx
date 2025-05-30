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

namespace TinaXlsx
{
    class TXWorksheetXmlHandler : public TXXmlHandler
    {
    public:
        explicit TXWorksheetXmlHandler(u64 sheetIndex): m_sheetIndex(sheetIndex)
        {
        }

        bool load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override
        {
            auto xmlData = zipReader.read(std::string(partName()));
            if (xmlData.empty())
            {
                m_lastError = "Failed to read " + std::string(partName());
                return false;
            }

            std::string xmlContent(xmlData.begin(), xmlData.end());
            TXXmlReader reader;
            if (!reader.parseFromString(xmlContent))
            {
                m_lastError = "Failed to parse worksheet.xml: " + reader.getLastError();
                return false;
            }
            // 解析 sheetData 节点，填充 context.sheets[sheetIndex_]
            auto cellNodes = reader.findNodes("//sheetData/row/c");
            for (const auto& cellNode : cellNodes)
            {
                std::string ref = cellNode.attributes.at("r");
                std::string value = cellNode.value;
                context.sheets[m_sheetIndex]->setCellValue(ref, value);
            }
            return true;
        }

        bool save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override
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
            writer.setRootNode(worksheet);
            std::string xmlContent = writer.generateXmlString();
            std::vector<uint8_t> xmlData(xmlContent.begin(), xmlContent.end());
            if (!zipWriter.write(std::string(partName()), xmlData)) {
                m_lastError = "Failed to write " + std::string(partName());
                return false;
            }
            return true;
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
