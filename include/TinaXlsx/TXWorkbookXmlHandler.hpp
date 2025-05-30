//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TXXmlHandler.hpp"
#include "TXXmlReader.hpp"
#include "TXXmlWriter.hpp"

namespace TinaXlsx
{
    class TXWorkbookXmlHandler : public TXXmlHandler
    {
    public:
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
                m_lastError = "Failed to parse workbook.xml: " + reader.getLastError();
                return false;
            }
            // 解析 sheets 节点，填充 context.sheets
            auto sheetNodes = reader.findNodes("//sheets/sheet");
            for (const auto& sheetNode : sheetNodes)
            {
                std::string name = sheetNode.attributes.at("name");
                std::string sheetId = sheetNode.attributes.at("sheetId");
                auto sheet = std::make_unique<TXSheet>(name, nullptr);
                context.sheets.push_back(std::move(sheet));
            }

            return true;
        }

        bool save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override
        {
            XmlNodeBuilder workbook("workbook");
            workbook.addAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                    .addAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            // 创建sheets节点
            XmlNodeBuilder sheetNode("sheets");
            for (std::size_t i = 0; i < context.sheets.size(); ++i)
            {
                XmlNodeBuilder sheet("sheet");
                sheet.addAttribute("name", context.sheets[i]->getName())
                     .addAttribute("sheetId", std::to_string(i + 1))
                     .addAttribute("r:id", "rId" + std::to_string(i + 1));
                sheetNode.addChild(sheet);
            }
            workbook.addChild(sheetNode);

            TXXmlWriter xmlWriter;
            xmlWriter.setRootNode(workbook);
            std::string xmlContent = xmlWriter.generateXmlString();
            std::vector<uint8_t> xmlData(xmlContent.begin(), xmlContent.end());
            if (!zipWriter.write(std::string(partName()), xmlData))
            {
                m_lastError = "Failed to write " + std::string(partName());
                return false;
            }
            return true;
        }


        std::string partName() const override
        {
            return "xl/workbook.xml";
        }
    };
}
