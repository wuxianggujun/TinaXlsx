//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TXSharedStringsPool.hpp"
#include "TXXmlHandler.hpp"
#include "TXXmlReader.hpp"
#include "TXXmlWriter.hpp"
#include "TXComponentManager.hpp"
#include "TXSharedStringsStreamWriter.hpp"

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

            // 统一使用流式写入器（高性能）
            return saveWithStreamWriter(zipWriter, context);
        }

    private:
        /**
         * @brief 使用流式写入器保存（高性能版本）
         */
        TXResult<void> saveWithStreamWriter(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context)
        {
            const auto& strings = context.sharedStringsPool.getStrings();

            // 创建流式写入器
            auto writer = TXSharedStringsWriterFactory::createWriter(strings.size());

            // 开始写入文档
            writer->startDocument(strings.size());

            // 逐个写入字符串
            for (const auto& str : strings) {
                // 检查是否需要保留空格（简单启发式：包含前导/尾随空格）
                bool preserveSpace = !str.empty() && (str.front() == ' ' || str.back() == ' ');
                writer->writeString(str, preserveSpace);
            }

            // 写入到ZIP文件
            auto writeResult = writer->writeToZip(zipWriter, std::string(partName()));
            if (writeResult.isError()) {
                return writeResult;
            }

            return Ok();
        }



        std::string partName() const override
        {
            return "xl/sharedStrings.xml";
        }
    };
}
