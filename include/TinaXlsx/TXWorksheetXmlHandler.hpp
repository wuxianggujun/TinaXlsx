//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TinaXlsx/TXXmlHandler.hpp"
#include "TinaXlsx/TXXmlReader.hpp"
#include "TinaXlsx/TXXmlWriter.hpp"
#include "TinaXlsx/TXSheet.hpp"

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

            XmlNodeBuilder sheetData("sheetData");

            // TODO: 这里应该根据 sheet 提供的行和单元格的迭代接口，生成 sheetData 节点下面的子节点
            // 示例：假设 sheet 提供了行和单元格的迭代接口
            // for (const auto& row : sheet->rows()) {
            //     XmlNodeBuilder rowNode("row");
            //     for (const auto& cell : row.cells()) {
            //         XmlNodeBuilder c("c");
            //         c.addAttribute("r", cell.ref());
            //         c.addChild(XmlNodeBuilder("v").setText(cell.value()));
            //         rowNode.addChild(c);
            //     }
            //     sheetData.addChild(rowNode);
            // }
            
            worksheet.addChild(sheetData);

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

        [[nodiscard]] std::string_view partName() const override {
            return "xl/worksheets/sheet" + std::to_string(m_sheetIndex + 1) + ".xml";
        }

    private:
        u64 m_sheetIndex;
    };
}
