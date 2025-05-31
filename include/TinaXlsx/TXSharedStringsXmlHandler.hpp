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
        bool load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override
        {
            auto xmlData = zipReader.read(partName());
            if (xmlData.isError())
            {
                m_lastError = "Failed to read " + partName();
                return false;
            }
            
            std::string xmlContent(xmlData.value().begin(), xmlData.value().end());
            TXXmlReader reader;
            auto parseResult = reader.parseFromString(xmlContent);
            if (parseResult.isError())
            {
                m_lastError = "Failed to parse sharedStrings.xml: " + parseResult.error().getMessage();
                return false;
            }

            // 解析共享字符串
            auto siNodesResult = reader.findNodes("//si/t");
            if (siNodesResult.isError())
            {
                m_lastError = "Failed to find shared string nodes: " + siNodesResult.error().getMessage();
                return false;
            }
            
            std::vector<std::string> sharedStrings;
            sharedStrings.reserve(siNodesResult.value().size());
            for (const auto& siNode : siNodesResult.value())
            {
                sharedStrings.push_back(siNode.value);
            }
            // 假设 TXStyleManager 或其他组件需要存储共享字符串
            return true;
        }

        bool save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override
        {
            // 获取共享字符串池中的字符串
            const auto& strings = context.sharedStringsPool.getStrings();
            
            // 如果没有字符串或共享字符串池不是dirty状态，则不生成文件
            if (strings.empty() || !context.sharedStringsPool.isDirty()) {
                return true;  // 跳过空池或未修改的池
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
                m_lastError = "Failed to set root node: " + setRootResult.error().getMessage();
                return false;
            }
            
            auto xmlContentResult = writer.generateXmlString();
            if (xmlContentResult.isError())
            {
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
            return "xl/sharedStrings.xml";
        }
    };
}
