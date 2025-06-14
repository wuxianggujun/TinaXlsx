//
// Created by wuxianggujun on 2025/5/29.
//

// StylesXmlHandler.hpp
#pragma once
#include "TXXmlHandler.hpp"
#include "TXXmlReader.hpp"
#include "TXXmlWriter.hpp"
#include "TXStyleManager.hpp"

namespace TinaXlsx {
    class StylesXmlHandler : public TXXmlHandler {
    public:
        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override {
            auto xmlDataResult = zipReader.read(std::string(partName())); // Returns TXResult<std::vector<uint8_t>>
            if (xmlDataResult.isError()) {
                return Err<void>(xmlDataResult.error().getCode(), "Failed to read " + std::string(partName()) + " from zip: " + xmlDataResult.error().getMessage());
            }
            const std::vector<uint8_t>& fileBytes = xmlDataResult.value(); // Get the actual std::vector<uint8_t>

            if (fileBytes.empty()) {
                return Err<void>(TXErrorCode::InvalidFileFormat, std::string(partName()) + " is empty (no content).");
            }
            
            std::string xmlContent(fileBytes.begin(), fileBytes.end());
            TXXmlReader reader;
            auto parseResult = reader.parseFromString(xmlContent);
            if (parseResult.isError()) {
                return Err<void>(parseResult.error().getCode(), "Failed to parse " + std::string(partName()) + ": " + parseResult.error().getMessage());
            }
            
            // TODO: 解析 cellXfs 节点，填充 context.styleManager
            // 解析样式（示例，需根据 TXStyleManager 实现）
            auto xfNodesResult = reader.findNodes("//cellXfs/xf");
            if (xfNodesResult.isError()) {
                return Err<void>(xfNodesResult.error().getCode(), "Failed to find cellXfs nodes: " + xfNodesResult.error().getMessage());
            }
            // 目前暂时返回成功，等待后续实现样式加载逻辑
            return Ok();
        }

        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override {
            // 使用styleManager的createStylesXmlNode方法生成样式XML
            XmlNodeBuilder styleSheet = context.styleManager.createStylesXmlNode();

            TXXmlWriter writer;
            auto setRootResult = writer.setRootNode(styleSheet);
            if (setRootResult.isError()) {
                return Err<void>(setRootResult.error().getCode(), "Failed to set root node: " + setRootResult.error().getMessage());
            }
            
            auto xmlContentResult = writer.generateXmlString();
            if (xmlContentResult.isError()) {
                return Err<void>(xmlContentResult.error().getCode(), "Failed to generate XML string: " + xmlContentResult.error().getMessage());
            }
            
            std::vector<uint8_t> xmlData(xmlContentResult.value().begin(), xmlContentResult.value().end());
            auto writeResult = zipWriter.write(std::string(partName()), xmlData);
            if (writeResult.isError()) {
                return Err<void>(writeResult.error().getCode(), "Failed to write " + std::string(partName()) + ": " + writeResult.error().getMessage());
            }
            return Ok();
        }

        std::string partName() const override {
            return "xl/styles.xml";
        }
    };
}
