//
// Created by wuxianggujun on 2025/1/15.
//

#include "TinaXlsx/TXPivotCacheRelsXmlHandler.hpp"
#include "TinaXlsx/TXError.hpp"
#include "TinaXlsx/TXXmlWriter.hpp"

namespace TinaXlsx
{
    TXPivotCacheRelsXmlHandler::TXPivotCacheRelsXmlHandler(int cacheId)
        : m_cacheId(cacheId)
    {
    }

    TXResult<void> TXPivotCacheRelsXmlHandler::load(TXZipArchiveReader& /*zipReader*/, TXWorkbookContext& /*context*/)
    {
        // 暂不实现加载功能
        return Err<void>(TXErrorCode::InvalidArgument, "Pivot cache rels loading not implemented");
    }

    TXResult<void> TXPivotCacheRelsXmlHandler::save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& /*context*/)
    {
        XmlNodeBuilder relationships("Relationships");
        relationships.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

        // 添加透视表缓存记录关系
        relationships.addChild(XmlNodeBuilder("Relationship")
                              .addAttribute("Id", "rId1")
                              .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords")
                              .addAttribute("Target", "pivotCacheRecords" + std::to_string(m_cacheId) + ".xml"));

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

    std::string TXPivotCacheRelsXmlHandler::partName() const
    {
        return "xl/pivotCache/_rels/pivotCacheDefinition" + std::to_string(m_cacheId) + ".xml.rels";
    }

} // namespace TinaXlsx
