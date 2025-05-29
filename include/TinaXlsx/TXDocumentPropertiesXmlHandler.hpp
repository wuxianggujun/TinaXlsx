//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once
#include "TinaXlsx/TXXmlHandler.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include "TinaXlsx/TXWorkbookContext.hpp"
#include "TinaXlsx/TXComponentManager.hpp"

namespace TinaXlsx
{

class TXDocumentPropertiesXmlHandler : public TXXmlHandler {
public:
    bool save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override {
        // 生成 XML 内容
        auto [core_xml, app_xml] = ComponentGenerator::generateDocumentProperties();

        // 转换为字节数据
        std::vector<uint8_t> core_data(core_xml.begin(), core_xml.end());
        std::vector<uint8_t> app_data(app_xml.begin(), app_xml.end());

        // 写入 ZIP 文件
        if (!zipWriter.write("docProps/core.xml", core_data) ||
            !zipWriter.write("docProps/app.xml", app_data)) {
            m_lastError = "写入文档属性 XML 文件失败";
            return false;
            }
        return true;
    }

    // 如果不需要读取，可以简单实现
    bool load(TXZipArchiveReader&, TXWorkbookContext&) override {
        // TODO: 读取文档属性 XML 文件
        return true; // 或者根据需要实现
    }

    [[nodiscard]] std::string partName() const override {
        return "docProps/"; // 定义路径前缀
    }
};
}
