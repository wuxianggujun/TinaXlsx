//
// Created by wuxianggujun on 2025/1/15.
//

#include "TinaXlsx/TXWorkbookRelsXmlHandler.hpp"
#include "TinaXlsx/TXXmlWriter.hpp"
#include "TinaXlsx/TXWorkbookContext.hpp"

namespace TinaXlsx {

    void TXWorkbookRelsXmlHandler::setAllPivotTables(const std::unordered_map<std::string, std::vector<std::shared_ptr<TXPivotTable>>>& allPivotTables) {
        m_allPivotTables = allPivotTables;
    }

    TXResult<void> TXWorkbookRelsXmlHandler::save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) {
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
            ++rid;
        }

        // 透视表缓存关系（如果有透视表）
        if (context.componentManager.hasComponent(ExcelComponent::PivotTables)) {
            // 计算总的透视表数量
            size_t totalPivotTables = 0;
            for (const auto& pair : m_allPivotTables) {
                totalPivotTables += pair.second.size();
            }

            // 为每个透视表缓存添加关系
            for (size_t i = 1; i <= totalPivotTables; ++i) {
                relationships.addChild(XmlNodeBuilder("Relationship")
                                       .addAttribute("Id", "rId" + std::to_string(rid))
                                       .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition")
                                       .addAttribute("Target", "pivotCache/pivotCacheDefinition" + std::to_string(i) + ".xml"));
                ++rid;
            }
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

} // namespace TinaXlsx
