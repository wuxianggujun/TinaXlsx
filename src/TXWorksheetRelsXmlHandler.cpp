//
// Created by wuxianggujun on 2025/1/15.
//

#include "TinaXlsx/TXWorksheetRelsXmlHandler.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXError.hpp"
#include "TinaXlsx/TXXmlWriter.hpp"

namespace TinaXlsx
{
    // ==================== TXWorksheetRelsXmlHandler 实现 ====================

    TXWorksheetRelsXmlHandler::TXWorksheetRelsXmlHandler(u32 sheetIndex)
        : m_sheetIndex(sheetIndex)
    {
    }

    TXResult<void> TXWorksheetRelsXmlHandler::load(TXZipArchiveReader& /*zipReader*/, TXWorkbookContext& /*context*/)
    {
        // 暂不实现加载功能
        return Err<void>(TXErrorCode::InvalidArgument, "Worksheet rels loading not implemented");
    }

    TXResult<void> TXWorksheetRelsXmlHandler::save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context)
    {
        if (m_sheetIndex >= context.sheets.size()) {
            return Err<void>(TXErrorCode::InvalidArgument, "Invalid sheet index");
        }

        const TXSheet* sheet = context.sheets[m_sheetIndex].get();

        // 检查是否有图表或透视表
        bool hasCharts = sheet->getChartCount() > 0;
        // 检查透视表（通过工作表关系处理器的设置）
        bool hasPivotTables = !m_pivotTables.empty();

        // 如果没有图表和透视表，不需要生成关系文件
        if (!hasCharts && !hasPivotTables) {
            return Ok();
        }

        XmlNodeBuilder relationships("Relationships");
        relationships.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

        u32 relationshipId = 1;

        // 添加透视表关系（如果有）
        if (hasPivotTables) {
            for (size_t i = 0; i < m_pivotTables.size(); ++i) {
                relationships.addChild(XmlNodeBuilder("Relationship")
                                      .addAttribute("Id", "rId" + std::to_string(relationshipId++))
                                      .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable")
                                      .addAttribute("Target", "../pivotTables/pivotTable" + std::to_string(i + 1) + ".xml"));
            }
        }

        // 添加绘图关系（如果有图表）
        if (hasCharts) {
            relationships.addChild(XmlNodeBuilder("Relationship")
                                  .addAttribute("Id", "rId" + std::to_string(relationshipId))
                                  .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing")
                                  .addAttribute("Target", "../drawings/drawing" + std::to_string(m_sheetIndex + 1) + ".xml"));
        }

        TXXmlWriter writer;
        auto setRootResult = writer.setRootNode(relationships);
        if (setRootResult.isError()) {
            return Err<void>(setRootResult.error().getCode(), 
                           "Failed to set root node: " + setRootResult.error().getMessage());
        }

        auto xmlContentResult = writer.generateXmlString();
        if (xmlContentResult.isError()) {
            return Err<void>(xmlContentResult.error().getCode(), 
                           "Failed to generate XML: " + xmlContentResult.error().getMessage());
        }

        std::vector<uint8_t> xmlData(xmlContentResult.value().begin(), xmlContentResult.value().end());
        auto writeResult = zipWriter.write(partName(), xmlData);
        if (writeResult.isError()) {
            return Err<void>(writeResult.error().getCode(), 
                           "Failed to write " + partName() + ": " + writeResult.error().getMessage());
        }

        return Ok();
    }

    void TXWorksheetRelsXmlHandler::setPivotTables(const std::vector<std::shared_ptr<TXPivotTable>>& pivotTables) {
        m_pivotTables = pivotTables;
    }

    std::string TXWorksheetRelsXmlHandler::partName() const
    {
        return "xl/worksheets/_rels/sheet" + std::to_string(m_sheetIndex + 1) + ".xml.rels";
    }

    // ==================== TXDrawingRelsXmlHandler 实现 ====================

    TXDrawingRelsXmlHandler::TXDrawingRelsXmlHandler(u32 sheetIndex)
        : m_sheetIndex(sheetIndex)
    {
    }

    TXResult<void> TXDrawingRelsXmlHandler::load(TXZipArchiveReader& /*zipReader*/, TXWorkbookContext& /*context*/)
    {
        // 暂不实现加载功能
        return Err<void>(TXErrorCode::InvalidArgument, "Drawing rels loading not implemented");
    }

    TXResult<void> TXDrawingRelsXmlHandler::save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context)
    {
        if (m_sheetIndex >= context.sheets.size()) {
            return Err<void>(TXErrorCode::InvalidArgument, "Invalid sheet index");
        }

        const TXSheet* sheet = context.sheets[m_sheetIndex].get();
        auto charts = sheet->getAllCharts();

        if (charts.empty()) {
            // 如果没有图表，不需要生成关系文件
            return Ok();
        }

        XmlNodeBuilder relationships("Relationships");
        relationships.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

        // 为每个图表添加关系
        for (size_t i = 0; i < charts.size(); ++i) {
            relationships.addChild(XmlNodeBuilder("Relationship")
                                  .addAttribute("Id", "rId" + std::to_string(i + 1))
                                  .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart")
                                  .addAttribute("Target", "../charts/chart" + std::to_string(i + 1) + ".xml"));
        }

        TXXmlWriter writer;
        auto setRootResult = writer.setRootNode(relationships);
        if (setRootResult.isError()) {
            return Err<void>(setRootResult.error().getCode(), 
                           "Failed to set root node: " + setRootResult.error().getMessage());
        }

        auto xmlContentResult = writer.generateXmlString();
        if (xmlContentResult.isError()) {
            return Err<void>(xmlContentResult.error().getCode(), 
                           "Failed to generate XML: " + xmlContentResult.error().getMessage());
        }

        std::vector<uint8_t> xmlData(xmlContentResult.value().begin(), xmlContentResult.value().end());
        auto writeResult = zipWriter.write(partName(), xmlData);
        if (writeResult.isError()) {
            return Err<void>(writeResult.error().getCode(), 
                           "Failed to write " + partName() + ": " + writeResult.error().getMessage());
        }

        return Ok();
    }

    std::string TXDrawingRelsXmlHandler::partName() const
    {
        return "xl/drawings/_rels/drawing" + std::to_string(m_sheetIndex + 1) + ".xml.rels";
    }

} // namespace TinaXlsx
