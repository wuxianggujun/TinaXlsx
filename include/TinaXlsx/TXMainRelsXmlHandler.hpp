//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TXXmlHandler.hpp"
#include "TXXmlReader.hpp"
#include "TXXmlWriter.hpp"

namespace TinaXlsx
{
    class TXMainRelsXmlHandler : public TXXmlHandler
    {
    public:
        bool load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override
        {
            // _rels/.rels 通常不需要加载
            return true;
        }

        bool save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override
        {
            XmlNodeBuilder relationships("Relationships");
            relationships.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

            // 工作簿关系
            relationships.addChild(XmlNodeBuilder("Relationship")
                                   .addAttribute("Id", "rId1")
                                   .addAttribute(
                                       "Type",
                                       "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
                                   .addAttribute("Target", "xl/workbook.xml"));

            // 文档属性（如果启用）
            if (context.componentManager.hasComponent(ExcelComponent::DocumentProperties))
            {
                relationships.addChild(XmlNodeBuilder("Relationship")
                                       .addAttribute("Id", "rId2")
                                       .addAttribute(
                                           "Type",
                                           "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
                                       .addAttribute("Target", "docProps/core.xml"));
                relationships.addChild(XmlNodeBuilder("Relationship")
                                       .addAttribute("Id", "rId3")
                                       .addAttribute(
                                           "Type",
                                           "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
                                       .addAttribute("Target", "docProps/app.xml"));
            }

            TXXmlWriter writer;
            auto setRootResult = writer.setRootNode(relationships);
            if (setRootResult.isError()) {
                m_lastError = "Failed to set root node: " + setRootResult.error().getMessage();
                return false;
            }
            
            auto xmlContentResult = writer.generateXmlString();
            if (xmlContentResult.isError()) {
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
            return "_rels/.rels";
        }
    };
}
