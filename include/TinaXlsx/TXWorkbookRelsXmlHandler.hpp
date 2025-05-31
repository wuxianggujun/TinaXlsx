//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TXXmlHandler.hpp"
#include "TXXmlReader.hpp"
#include "TXXmlWriter.hpp"

namespace TinaXlsx {
    class TXWorkbookRelsXmlHandler : public TXXmlHandler {
    public:
        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override {
            // xl/_rels/workbook.xml.rels 通常不需要加载
            return Ok();
        }

        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override {
            XmlNodeBuilder relationships("Relationships");
            relationships.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

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
            auto setRootResult = writer.setRootNode(relationships);
            if (setRootResult.isError()) {
                return Err<void>(setRootResult.error().getCode(), "Failed to set root node: " + setRootResult.error().getMessage());
            }
            
            auto xmlContentResult = writer.generateXmlString();
            if (xmlContentResult.isError()) {
                return Err<void>(xmlContentResult.error().getCode(), "Failed to generate XML: " + xmlContentResult.error().getMessage());
            }
            
            std::vector<uint8_t> xmlData(xmlContentResult.value().begin(), xmlContentResult.value().end());
            auto writeResult = zipWriter.write(std::string(partName()), xmlData);
            if (writeResult.isError()) {
                return Err<void>(writeResult.error().getCode(), "Failed to write " + std::string(partName()) + ": " + writeResult.error().getMessage());
            }
            return Ok();
        }

        std::string partName() const override {
            return "xl/_rels/workbook.xml.rels";
        }
    };
}
