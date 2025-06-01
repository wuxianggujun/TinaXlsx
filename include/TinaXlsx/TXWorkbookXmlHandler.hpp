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
        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override
        {
            auto xmlData = zipReader.read(partName());
            if (xmlData.isError())
            {
                return Err<void>(xmlData.error().getCode(), "Failed to read " + std::string(partName()));
            }
            const std::vector<uint8_t>& fileBytes = xmlData.value(); // Get the actual std::vector<uint8_t>

            std::string xmlContent(fileBytes.begin(), fileBytes.end());
            TXXmlReader reader;
            auto parseResult = reader.parseFromString(xmlContent);
            if (parseResult.isError())
            {
                return Err<void>(parseResult.error().getCode(), "Failed to parse workbook.xml: " + parseResult.error().getMessage());
            }
            
            // 解析 sheets 节点，填充 context.sheets
            auto sheetNodesResult = reader.findNodes("//sheets/sheet");
            if (sheetNodesResult.isError())
            {
                return Err<void>(sheetNodesResult.error().getCode(), "Failed to find sheet nodes: " + sheetNodesResult.error().getMessage());
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

            return Ok();
        }

        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override
        {
            XmlNodeBuilder workbook("workbook");
            workbook.addAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                    .addAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            // 添加工作簿保护信息
            auto& workbookProtectionManager = context.workbookProtectionManager;
            if (workbookProtectionManager.isWorkbookProtected()) {
                const auto& protection = workbookProtectionManager.getWorkbookProtection();
                XmlNodeBuilder workbookProtection("workbookProtection");

                // 添加现代Excel的SHA-512密码保护属性
                if (!protection.passwordHash.empty()) {
                    workbookProtection.addAttribute("workbookAlgorithmName", protection.algorithmName);
                    workbookProtection.addAttribute("workbookHashValue", protection.passwordHash);
                    workbookProtection.addAttribute("workbookSaltValue", protection.saltValue);
                    workbookProtection.addAttribute("workbookSpinCount", std::to_string(protection.spinCount));
                }

                // 添加保护选项属性
                if (protection.lockStructure) {
                    workbookProtection.addAttribute("lockStructure", "1");
                }
                if (protection.lockWindows) {
                    workbookProtection.addAttribute("lockWindows", "1");
                }
                if (protection.lockRevision) {
                    workbookProtection.addAttribute("lockRevision", "1");
                }

                workbook.addChild(workbookProtection);
            }

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
                return Err<void>(setRootResult.error().getCode(), "Failed to set root node: " + setRootResult.error().getMessage());
            }
            
            auto xmlContentResult = xmlWriter.generateXmlString();
            if (xmlContentResult.isError())
            {
                return Err<void>(xmlContentResult.error().getCode(), "Failed to generate XML: " + xmlContentResult.error().getMessage());
            }
            
            std::vector<uint8_t> xmlData(xmlContentResult.value().begin(), xmlContentResult.value().end());
            auto writeResult = zipWriter.write(std::string(partName()), xmlData);
            if (writeResult.isError())
            {
                return Err<void>(writeResult.error().getCode(), "Failed to write " + std::string(partName()) + ": " + writeResult.error().getMessage());
            }
            return Ok();
        }


        std::string partName() const override
        {
            return "xl/workbook.xml";
        }
    };
}
