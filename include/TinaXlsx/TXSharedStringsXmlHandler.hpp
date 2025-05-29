//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TinaXlsx/TXXmlHandler.hpp"
#include "TinaXlsx/TXXmlReader.hpp"
#include "TinaXlsx/TXXmlWriter.hpp"
#include "TinaXlsx/TXComponentManager.hpp"

namespace TinaXlsx {
    class TXSharedStringsXmlHandler : public TXXmlHandler {
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
                m_lastError = "Failed to parse sharedStrings.xml: " + reader.getLastError();
                return false;
            }

            // 解析共享字符串
            auto siNodes = reader.findNodes("//si/t");
            std::vector<std::string> sharedStrings;
            for (const auto& siNode : siNodes) {
                sharedStrings.push_back(siNode.value);
            }
            // 假设 TXStyleManager 或其他组件需要存储共享字符串
            return true;
        }

        bool save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override {
            if (!context.componentManager.hasComponent(ExcelComponent::SharedStrings)) {
                return true; // 如果未启用，直接返回
            }

            // 收集所有工作表中的字符串
            std::vector<std::string> strings;
            for (const auto& sheet : context.sheets) {
                auto usedRange = sheet->getUsedRange();
                if (usedRange.isValid()) {
                    for (row_t row = usedRange.getStart().getRow(); row <= usedRange.getEnd().getRow(); ++row) {
                        for (column_t col = usedRange.getStart().getCol(); col <= usedRange.getEnd().getCol(); ++col) {
                            const TXCell* cell = sheet->getCell(row, col);
                            if (cell && std::holds_alternative<std::string>(cell->getValue())) {
                                strings.push_back(std::get<std::string>(cell->getValue()));
                            }
                        }
                    }
                }
            }

            // 生成 XML
            XmlNodeBuilder sst("sst");
            sst.addAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
               .addAttribute("count", std::to_string(strings.size()))
               .addAttribute("uniqueCount", std::to_string(strings.size()));

            for (const auto& str : strings) {
                XmlNodeBuilder si("si");
                si.addChild(XmlNodeBuilder("t").setText(str));
                sst.addChild(si);
            }

            TXXmlWriter writer;
            writer.setRootNode(sst);
            std::string xmlContent = writer.generateXmlString();
            std::vector<uint8_t> xmlData(xmlContent.begin(), xmlContent.end());
            if (!zipWriter.write(std::string(partName()), xmlData)) {
                m_lastError = "Failed to write " + std::string(partName());
                return false;
            }
            return true;
        }

        std::string partName() const override {
            return "xl/sharedStrings.xml";
        }
    };
}
