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
            // 解析 cellXfs 节点，填充 context.styleManager
            auto styleNodes = reader.findNodes("//cellXfs/xf");
            for (const auto& styleNode : styleNodes) {
                // 假设 styleManager 有 addStyle 接口
                // context.styleManager.addStyle(parseStyle(styleNode));
            }
            return true;
        }

        bool save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override {
            XmlNodeBuilder styleSheet("styleSheet");
            styleSheet.addAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            XmlNodeBuilder cellXfs("cellXfs");
            // 示例：假设 styleManager 提供样式迭代
            // for (const auto& style : context.styleManager.styles()) {
            //     XmlNodeBuilder xf("xf");
            //     xf.addAttribute("numFmtId", style.numFmtId());
            //     cellXfs.addChild(xf);
            // }
            styleSheet.addChild(cellXfs);

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

        std::string_view partName() const override {
            return "xl/styles.xml";
        }
    };
}
