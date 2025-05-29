//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TinaXlsx/TXXmlHandler.hpp"
#include "TinaXlsx/TXXmlReader.hpp"
#include "TinaXlsx/TXXmlWriter.hpp"

namespace TinaXlsx {
    class TXWorkbookRelsXmlHandler : public TXXmlHandler {
    public:
        bool load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override {
            // xl/_rels/workbook.xml.rels 通常不需要加载
            return true;
        }

        bool save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override {
            XmlNodeBuilder relationships("Relationships");
            relationships.addAttribute("xmlns", "http://schemas.openxmlformats.org/relationships");

            size_t rid = 1;

            // 工作表关系
            for (size_t i = 0; i < context.sheets.size(); ++i) {
                relationships.addChild(XmlNodeBuilder("Relationship")
                                       .addAttribute("Id", "rId" + std::to_string(rid))
                                       .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
                                       .addAttribute("Target", "worksheets/sheet" + std::to_string(i + 1) + ".xml"));
                ++rid;
            }

            // 样式关系（如果启用）
            if (context.componentManager.hasComponent(ExcelComponent::Styles)) {
                relationships.addChild(XmlNodeBuilder("Relationship")
                                       .addAttribute("Id", "rId" + std::to_string(rid))
                                       .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
                                       .addAttribute("Target", "styles.xml"));
                ++rid;
            }

            // 共享字符串关系（如果启用）
            if (context.componentManager.hasComponent(ExcelComponent::SharedStrings)) {
                relationships.addChild(XmlNodeBuilder("Relationship")
                                       .addAttribute("Id", "rId" + std::to_string(rid))
                                       .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
                                       .addAttribute("Target", "sharedStrings.xml"));
            }

            TXXmlWriter writer;
            writer.setRootNode(relationships);
            std::string xmlContent = writer.generateXmlString();
            std::vector<uint8_t> xmlData(xmlContent.begin(), xmlContent.end());
            if (!zipWriter.write(std::string(partName()), xmlData)) {
                m_lastError = "Failed to write " + std::string(partName());
                return false;
            }
            return true;
        }

        std::string partName() const override {
            return "xl/_rels/workbook.xml.rels";
        }
    };
}
