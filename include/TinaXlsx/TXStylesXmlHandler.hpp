//
// Created by wuxianggujun on 2025/5/29.
//

// StylesXmlHandler.hpp
#pragma once
#include "TinaXlsx/TXXmlHandler.hpp"
#include "TinaXlsx/TXXmlReader.hpp"
#include "TinaXlsx/TXXmlWriter.hpp"
#include "TinaXlsx/TXStyleManager.hpp"

namespace TinaXlsx {
    class StylesXmlHandler : public TXXmlHandler {
    public:
        bool load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override {
            auto xmlData = zipReader.read(std::string(partName()));
            if (xmlData.empty()) {
                m_lastError = "Failed to read " + std::string(partName());
                return false;
            }
            std::string xmlContent(xmlData.begin(), xmlData.end());
            TXXmlReader reader;
            if (!reader.parseFromString(xmlContent)) {
                m_lastError = "Failed to parse styles.xml: " + reader.getLastError();
                return false;
            }
            
            // TODO: 解析 cellXfs 节点，填充 context.styleManager
            // 目前暂时返回true，等待后续实现样式加载逻辑
            return true;
        }

        bool save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override {
            // 使用styleManager的createStylesXmlNode方法生成样式XML
            XmlNodeBuilder styleSheet = context.styleManager.createStylesXmlNode();

            TXXmlWriter writer;
            writer.setRootNode(styleSheet);
            std::string xmlContent = writer.generateXmlString();
            std::vector<uint8_t> xmlData(xmlContent.begin(), xmlContent.end());
            if (!zipWriter.write(std::string(partName()), xmlData)) {
                m_lastError = "Failed to write " + std::string(partName());
                return false;
            }
            return true;
        }

        std::string partName() const override {
            return "xl/styles.xml";
        }
    };
}
