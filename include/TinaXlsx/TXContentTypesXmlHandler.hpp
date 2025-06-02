//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TXXmlHandler.hpp"
#include "TXXmlReader.hpp"
#include "TXXmlWriter.hpp"

namespace TinaXlsx
{
    class TXContentTypesXmlHandler : public TXXmlHandler
    {
        public:
        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override
        {
            return Ok();
        }

        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override
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

            // 透视表内容类型（如果启用）
            if (context.componentManager.hasComponent(ExcelComponent::PivotTables)) {
                // 计算透视表总数
                u32 pivotTableCount = 0;
                // 这里需要从工作簿上下文中获取透视表信息
                // 暂时假设每个工作表最多有一个透视表
                for (u64 i = 0; i < context.sheets.size(); ++i) {
                    // 假设有透视表（实际应该从工作簿中获取）
                    ++pivotTableCount;

                    // 透视表定义内容类型
                    types.addChild(XmlNodeBuilder("Override")
                                   .addAttribute("PartName", "/xl/pivotTables/pivotTable" + std::to_string(pivotTableCount) + ".xml")
                                   .addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"));

                    // 透视表缓存定义内容类型
                    types.addChild(XmlNodeBuilder("Override")
                                   .addAttribute("PartName", "/xl/pivotCache/pivotCacheDefinition" + std::to_string(pivotTableCount) + ".xml")
                                   .addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"));

                    // 透视表缓存记录内容类型
                    types.addChild(XmlNodeBuilder("Override")
                                   .addAttribute("PartName", "/xl/pivotCache/pivotCacheRecords" + std::to_string(pivotTableCount) + ".xml")
                                   .addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"));
                }
            }

            // 图表和绘图内容类型
            u32 chartCount = 0;
            for (u64 i = 0; i < context.sheets.size(); ++i) {
                const TXSheet* sheet = context.sheets[i].get();
                if (sheet->getChartCount() > 0) {
                    // 绘图内容类型
                    types.addChild(XmlNodeBuilder("Override")
                                   .addAttribute("PartName", "/xl/drawings/drawing" + std::to_string(i + 1) + ".xml")
                                   .addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml"));

                    // 图表内容类型
                    for (size_t j = 0; j < sheet->getChartCount(); ++j) {
                        types.addChild(XmlNodeBuilder("Override")
                                       .addAttribute("PartName", "/xl/charts/chart" + std::to_string(chartCount + 1) + ".xml")
                                       .addAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"));
                        ++chartCount;
                    }
                }
            }

            TXXmlWriter writer;
            auto setRootResult = writer.setRootNode(types);
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

        std::string partName() const override
        {
            return "[Content_Types].xml";
        }
        
    };
}
