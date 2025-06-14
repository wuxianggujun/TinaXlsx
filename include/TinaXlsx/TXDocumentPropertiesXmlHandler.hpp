//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once
#include "TXXmlHandler.hpp"
#include "TXZipArchive.hpp"
#include "TXWorkbookContext.hpp"
#include "TXComponentManager.hpp"

namespace TinaXlsx
{
    class TXDocumentPropertiesXmlHandler : public TXXmlHandler
    {
    public:
        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override
        {
            // 生成 core.xml
            XmlNodeBuilder coreProps("cp:coreProperties");
            coreProps.addAttribute(
                         "xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
                     .addAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
                     .addAttribute("xmlns:dcterms", "http://purl.org/dc/terms/")
                     .addAttribute("xmlns:dcmitype", "http://purl.org/dc/dcmitype/")
                     .addAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");
            coreProps.addChild(XmlNodeBuilder("dc:creator").setText("TinaXlsx"));
            coreProps.addChild(XmlNodeBuilder("cp:lastModifiedBy").setText("TinaXlsx"));
            coreProps.addChild(XmlNodeBuilder("dcterms:created")
                               .addAttribute("xsi:type", "dcterms:W3CDTF")
                               .setText("2025-05-29T00:00:00Z"));
            coreProps.addChild(XmlNodeBuilder("dcterms:modified")
                               .addAttribute("xsi:type", "dcterms:W3CDTF")
                               .setText("2025-05-29T00:00:00Z"));

            TXXmlWriter coreWriter;
            auto setCoreRootResult = coreWriter.setRootNode(coreProps);
            if (setCoreRootResult.isError()) {
                return Err<void>(setCoreRootResult.error().getCode(), "Failed to set core root node: " + setCoreRootResult.error().getMessage());
            }
            
            auto coreContentResult = coreWriter.generateXmlString();
            if (coreContentResult.isError()) {
                return Err<void>(coreContentResult.error().getCode(), "Failed to generate core XML string: " + coreContentResult.error().getMessage());
            }
            
            std::vector<uint8_t> coreData(coreContentResult.value().begin(), coreContentResult.value().end());
            auto writeCoreResult = zipWriter.write("docProps/core.xml", coreData);
            if (writeCoreResult.isError())
            {
                return Err<void>(writeCoreResult.error().getCode(), "Failed to write docProps/core.xml: " + writeCoreResult.error().getMessage());
            }

            // 生成 app.xml
            XmlNodeBuilder appProps("Properties");
            appProps.addAttribute("xmlns", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")
                    .addAttribute("xmlns:vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            appProps.addChild(XmlNodeBuilder("Application").setText("TinaXlsx"));
            appProps.addChild(XmlNodeBuilder("DocSecurity").setText("0"));
            appProps.addChild(XmlNodeBuilder("ScaleCrop").setText("false"));
            appProps.addChild(XmlNodeBuilder("SharedDoc").setText("false"));
            appProps.addChild(XmlNodeBuilder("HyperlinksChanged").setText("false"));
            appProps.addChild(XmlNodeBuilder("AppVersion").setText("16.0300"));

            TXXmlWriter appWriter;
            auto setAppRootResult = appWriter.setRootNode(appProps);
            if (setAppRootResult.isError()) {
                return Err<void>(setAppRootResult.error().getCode(), "Failed to set app root node: " + setAppRootResult.error().getMessage());
            }
            
            auto appContentResult = appWriter.generateXmlString();
            if (appContentResult.isError()) {
                return Err<void>(appContentResult.error().getCode(), "Failed to generate app XML string: " + appContentResult.error().getMessage());
            }
            
            std::vector<uint8_t> appData(appContentResult.value().begin(), appContentResult.value().end());
            auto writeAppResult = zipWriter.write("docProps/app.xml", appData);
            if (writeAppResult.isError())
            {
                return Err<void>(writeAppResult.error().getCode(), "Failed to write docProps/app.xml: " + writeAppResult.error().getMessage());
            }

            return Ok();
        }

        // 如果不需要读取，可以简单实现
        TXResult<void> load(TXZipArchiveReader&, TXWorkbookContext&) override
        {
            // TODO: 读取文档属性 XML 文件
            return Ok(); // 或者根据需要实现
        }

        [[nodiscard]] std::string partName() const override
        {
            return "docProps/"; // 定义路径前缀
        }
    };
}
