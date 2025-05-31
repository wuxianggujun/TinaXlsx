//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TXSharedStringsPool.hpp"
#include "TXXmlHandler.hpp"
#include "TXXmlReader.hpp"
#include "TXXmlWriter.hpp"
#include "TXComponentManager.hpp"

namespace TinaXlsx
{
    class TXSharedStringsXmlHandler : public TXXmlHandler
    {
    public:
        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override
        {
            auto xmlData = zipReader.read(partName());
            if (xmlData.isError())
            {
                return Err<void>(xmlData.error().getCode(), "Failed to read " + partName());
            }
            
            std::string xmlContent(xmlData.value().begin(), xmlData.value().end());
            TXXmlReader reader;
            auto parseResult = reader.parseFromString(xmlContent);
            if (parseResult.isError())
            {
                return Err<void>(parseResult.error().getCode(), "Failed to parse sharedStrings.xml: " + parseResult.error().getMessage());
            }

            // 解析共享字符串
            auto siNodesResult = reader.findNodes("//si/t");
            if (siNodesResult.isError())
            {
                return Err<void>(siNodesResult.error().getCode(), "Failed to find shared string nodes: " + siNodesResult.error().getMessage());
            }
            
            std::vector<std::string> sharedStrings;
            sharedStrings.reserve(siNodesResult.value().size());
            for (const auto& siNode : siNodesResult.value())
            {
                sharedStrings.push_back(siNode.value);
            }
            // 假设 TXStyleManager 或其他组件需要存储共享字符串
            return Ok();
        }

        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override
        {
            // 获取共享字符串池中的字符串
            const auto& strings = context.sharedStringsPool.getStrings();
            
            // 如果没有字符串或共享字符串池不是dirty状态，则不生成文件
            if (strings.empty() || !context.sharedStringsPool.isDirty()) {
                return Ok();  // 跳过空池或未修改的池
            }
            
            // 生成 XML
            XmlNodeBuilder sst("sst");
            sst.addAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
               .addAttribute("count", std::to_string(strings.size()))
               .addAttribute("uniqueCount", std::to_string(strings.size()));

            for (const auto& str : strings)
            {
                XmlNodeBuilder si("si");
                si.addChild(XmlNodeBuilder("t").setText(str));
                sst.addChild(si);
            }

            TXXmlWriter writer;
            auto setRootResult = writer.setRootNode(sst);
            if (setRootResult.isError())
            {
                return Err<void>(setRootResult.error().getCode(), "Failed to set root node: " + setRootResult.error().getMessage());
            }
            
            auto xmlContentResult = writer.generateXmlString();
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
            return "xl/sharedStrings.xml";
        }
    };
}
