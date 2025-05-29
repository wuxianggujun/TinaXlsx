//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TinaXlsx/TXXmlHandler.hpp"
#include "TinaXlsx/TXXmlReader.hpp"
#include "TinaXlsx/TXXmlWriter.hpp"

namespace TinaXlsx
{
    class TXContentTypesXmlHandler : public TXXmlHandler
    {
        public:
        bool load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override
        {
            return true;
        }

        bool save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override
        {
            XmlNodeBuilder types("Types");

            types.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/content-types");

            types.addChild(XmlNodeBuilder("Default")
                .addAttribute("Extension","rels")
                .addAttribute("ContentType","application/vnd.openxmlformats-package.relationships+xml"));

            types.addChild(XmlNodeBuilder("Default")
                .addAttribute("Extension","xml")
                .addAttribute("ContentType","application/xml"));

            types.addChild(XmlNodeBuilder("Override")
                .addAttribute("PartName","/xl/workbook.xml")
                .addAttribute("ContentType","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"));

            // 工作表
            for (u64 i = 0; i < context.sheets.size(); ++i)
            {
                types.addChild(XmlNodeBuilder("Override")
                    .addAttribute("PartName", "/xl/worksheets/sheet" + std::to_string(i + 1) + ".xml")
                    .addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"));
            }

            // 样式（如果启用）
            if (context.componentManager.hasComponent(ExcelComponent::Styles)) {
                types.addChild(XmlNodeBuilder("Override")
                               .addAttribute("PartName", "/xl/styles.xml")
                               .addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"));
            }

            // 共享字符串（如果启用）
            if (context.componentManager.hasComponent(ExcelComponent::SharedStrings)) {
                types.addChild(XmlNodeBuilder("Override")
                               .addAttribute("PartName", "/xl/sharedStrings.xml")
                               .addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"));
            }

            // 文档属性（如果启用）
            if (context.componentManager.hasComponent(ExcelComponent::DocumentProperties)) {
                types.addChild(XmlNodeBuilder("Override")
                               .addAttribute("PartName", "/docProps/core.xml")
                               .addAttribute("ContentType", "application/vnd.openxmlformats-package.core-properties+xml"));
                types.addChild(XmlNodeBuilder("Override")
                               .addAttribute("PartName", "/docProps/app.xml")
                               .addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml"));
            }

            TXXmlWriter writer;
            writer.setRootNode(types);
            std::string xmlContent = writer.generateXmlString();
            std::vector<uint8_t> xmlData(xmlContent.begin(), xmlContent.end());
            if (!zipWriter.write(std::string(partName()), xmlData)) {
                m_lastError = "Failed to write " + std::string(partName());
                return false;
            }

            return true;
        }

        std::string partName() const override
        {
            return "[Content_Types].xml";
        }
        
    };
}
