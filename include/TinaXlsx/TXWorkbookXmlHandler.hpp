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
            auto xmlData = zipReader.read(partName());
            if (xmlData.isError())
            {
                m_lastError = "Failed to read " + std::string(partName());
                return false;
            }
            const std::vector<uint8_t>& fileBytes = xmlData.value(); // Get the actual std::vector<uint8_t>

            std::string xmlContent(fileBytes.begin(), fileBytes.end());
            TXXmlReader reader;
            auto parseResult = reader.parseFromString(xmlContent);
            if (parseResult.isError())
            {
                m_lastError = "Failed to parse workbook.xml: " + parseResult.error().getMessage();
                return false;
            }
            
            // 解析 sheets 节点，填充 context.sheets
            auto sheetNodesResult = reader.findNodes("//sheets/sheet");
            if (sheetNodesResult.isError())
            {
                m_lastError = "Failed to find sheet nodes: " + sheetNodesResult.error().getMessage();
                return false;
            }
            
            for (const auto& sheetNode : sheetNodesResult.value())
            {
                auto nameIter = sheetNode.attributes.find("name");
                auto sheetIdIter = sheetNode.attributes.find("sheetId");
                if (nameIter != sheetNode.attributes.end() && sheetIdIter != sheetNode.attributes.end())
                {
                    std::string name = nameIter->second;
                    std::string sheetId = sheetIdIter->second;
                    auto sheet = std::make_unique<TXSheet>(name, nullptr);
                    context.sheets.push_back(std::move(sheet));
                }
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
            auto setRootResult = xmlWriter.setRootNode(workbook);
            if (setRootResult.isError())
            {
                m_lastError = "Failed to set root node: " + setRootResult.error().getMessage();
                return false;
            }
            
            auto xmlContentResult = xmlWriter.generateXmlString();
            if (xmlContentResult.isError())
            {
                m_lastError = "Failed to generate XML: " + xmlContentResult.error().getMessage();
                return false;
            }
            
            std::vector<uint8_t> xmlData(xmlContentResult.value().begin(), xmlContentResult.value().end());
            auto writeResult = zipWriter.write(std::string(partName()), xmlData);
            if (writeResult.isError())
            {
                m_lastError = "Failed to write " + std::string(partName()) + ": " + writeResult.error().getMessage();
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
