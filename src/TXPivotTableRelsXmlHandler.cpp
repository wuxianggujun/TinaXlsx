//
// Created by wuxianggujun on 2025/1/15.
//

#include "TinaXlsx/TXPivotTableRelsXmlHandler.hpp"
#include "TinaXlsx/TXError.hpp"
#include "TinaXlsx/TXXmlWriter.hpp"

namespace TinaXlsx
{
    TXPivotTableRelsXmlHandler::TXPivotTableRelsXmlHandler(int pivotTableId)
        : m_pivotTableId(pivotTableId)
    {
    }

    TXResult<void> TXPivotTableRelsXmlHandler::load(TXZipArchiveReader& /*zipReader*/, TXWorkbookContext& /*context*/)
    {
        // 暂不实现加载功能
        return Err<void>(TXErrorCode::InvalidArgument, "Pivot table rels loading not implemented");
    }

    TXResult<void> TXPivotTableRelsXmlHandler::save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& /*context*/)
    {
        XmlNodeBuilder relationships("Relationships");
        relationships.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

        // 添加透视表缓存定义关系
        // 透视表通过工作簿关系引用缓存定义，关系ID计算：
        // 工作表数量 + 共享字符串(1) + 透视表缓存ID
        int workbookRId = 5 + 1 + m_pivotTableId; // 5个工作表 + 1个共享字符串 + 缓存ID
        relationships.addChild(XmlNodeBuilder("Relationship")
                              .addAttribute("Id", "rId1")
                              .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition")
                              .addAttribute("Target", "../_rels/workbook.xml.rels#rId" + std::to_string(workbookRId)));

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

    std::string TXPivotTableRelsXmlHandler::partName() const
    {
        return "xl/pivotTables/_rels/pivotTable" + std::to_string(m_pivotTableId) + ".xml.rels";
    }

} // namespace TinaXlsx
