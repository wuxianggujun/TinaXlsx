//
// @file TXUnifiedXmlHandler.cpp
// @brief 统一的简单XML处理器实现
//

#include "TinaXlsx/TXUnifiedXmlHandler.hpp"
#include "TinaXlsx/TXComponentManager.hpp"
#include "TinaXlsx/TXPivotTable.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXPugiStreamWriter.hpp"  // 流式写入器
#include <sstream>
#include <iomanip>
#include <cstring>  // for strlen
#include <ctime>    // for time functions

namespace TinaXlsx {

// ==================== TXUnifiedXmlHandler 实现 ====================

TXUnifiedXmlHandler::TXUnifiedXmlHandler(HandlerType type, u32 index)
    : type_(type), index_(index) {
}

TXResult<void> TXUnifiedXmlHandler::load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) {
    // 大部分简单处理器不需要加载逻辑
    return Ok();
}

TXResult<void> TXUnifiedXmlHandler::save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) {
    // 对于工作表关系，如果没有图表和透视表，不需要生成文件
    if (type_ == HandlerType::WorksheetRels) {
        if (index_ >= context.sheets.size()) {
            return Err<void>(TXErrorCode::InvalidArgument, "Invalid sheet index");
        }

        const TXSheet* sheet = context.sheets[index_].get();
        bool hasCharts = sheet->getChartCount() > 0;
        bool hasPivotTables = !pivot_tables_.empty();

        // 如果没有图表和透视表，不需要生成关系文件
        if (!hasCharts && !hasPivotTables) {
            return Ok();
        }
    }

    // 使用流式写入器生成XML
    switch (type_) {
        case HandlerType::MainRels:
            return generateMainRelsStream(zipWriter, context);
        case HandlerType::WorkbookRels:
            return generateWorkbookRelsStream(zipWriter, context);
        case HandlerType::WorksheetRels:
            return generateWorksheetRelsStream(zipWriter, context);
        case HandlerType::DocumentProperties:
            return generateDocumentPropertiesStream(zipWriter, context);
        case HandlerType::PivotTableRels:
            return generatePivotTableRelsStream(zipWriter, context);
        case HandlerType::PivotCacheRels:
            return generatePivotCacheRelsStream(zipWriter, context);
        default:
            return Err<void>(TXErrorCode::InvalidArgument, "Unknown handler type");
    }
}

std::string TXUnifiedXmlHandler::partName() const {
    return getPartName();
}

void TXUnifiedXmlHandler::setPivotTables(const std::vector<std::shared_ptr<TXPivotTable>>& pivotTables) {
    pivot_tables_ = pivotTables;
}

void TXUnifiedXmlHandler::setAllPivotTables(const std::unordered_map<std::string, std::vector<std::shared_ptr<TXPivotTable>>>& allPivotTables) {
    all_pivot_tables_ = allPivotTables;
}

// ==================== 流式生成方法 ====================

TXResult<void> TXUnifiedXmlHandler::generateMainRelsStream(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) const {
    TXBufferedXmlWriter writer(4096); // 4KB缓冲区，关系文件很小

    // 写入XML声明
    const char* xmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
    writer.write(xmlDecl, strlen(xmlDecl));

    // 写入根元素开始标签
    const char* rootStart = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
    writer.write(rootStart, strlen(rootStart));

    // 工作簿关系
    const char* workbookRel = "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>";
    writer.write(workbookRel, strlen(workbookRel));

    // 文档属性（如果启用）
    if (context.componentManager.hasComponent(ExcelComponent::DocumentProperties)) {
        const char* coreRel = "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>";
        writer.write(coreRel, strlen(coreRel));

        const char* appRel = "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>";
        writer.write(appRel, strlen(appRel));
    }

    // 写入根元素结束标签
    const char* rootEnd = "</Relationships>";
    writer.write(rootEnd, strlen(rootEnd));

    // 写入到ZIP文件
    auto writeResult = zipWriter.write(getPartName(), writer.getBuffer());
    if (writeResult.isError()) {
        return Err<void>(writeResult.error().getCode(),
                       "Failed to write " + getPartName() + ": " + writeResult.error().getMessage());
    }

    return Ok();
}

TXResult<void> TXUnifiedXmlHandler::generateWorkbookRelsStream(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) const {
    TXBufferedXmlWriter writer(8192); // 8KB缓冲区，工作簿关系文件稍大

    // 写入XML声明
    const char* xmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
    writer.write(xmlDecl, strlen(xmlDecl));

    // 写入根元素开始标签
    const char* rootStart = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
    writer.write(rootStart, strlen(rootStart));

    u32 relationshipId = 1;

    // 工作表关系
    for (u64 i = 0; i < context.sheets.size(); ++i) {
        std::string worksheetRel = "<Relationship Id=\"rId" + std::to_string(relationshipId++) +
                                  "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"" +
                                  " Target=\"worksheets/sheet" + std::to_string(i + 1) + ".xml\"/>";
        writer.write(worksheetRel.c_str(), worksheetRel.length());
    }

    // 样式关系（如果启用）
    if (context.componentManager.hasComponent(ExcelComponent::Styles)) {
        std::string stylesRel = "<Relationship Id=\"rId" + std::to_string(relationshipId++) +
                               "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\"" +
                               " Target=\"styles.xml\"/>";
        writer.write(stylesRel.c_str(), stylesRel.length());
    }

    // 共享字符串关系（如果启用）
    if (context.componentManager.hasComponent(ExcelComponent::SharedStrings)) {
        std::string sharedStringsRel = "<Relationship Id=\"rId" + std::to_string(relationshipId++) +
                                      "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\"" +
                                      " Target=\"sharedStrings.xml\"/>";
        writer.write(sharedStringsRel.c_str(), sharedStringsRel.length());
    }

    // 透视表缓存关系（如果有透视表）
    if (!all_pivot_tables_.empty()) {
        u32 cacheId = 1;
        for (const auto& [sheetName, pivotTables] : all_pivot_tables_) {
            for (const auto& pivotTable : pivotTables) {
                std::string pivotCacheRel = "<Relationship Id=\"rId" + std::to_string(relationshipId++) +
                                           "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\"" +
                                           " Target=\"pivotCache/pivotCacheDefinition" + std::to_string(cacheId++) + ".xml\"/>";
                writer.write(pivotCacheRel.c_str(), pivotCacheRel.length());
            }
        }
    }

    // 写入根元素结束标签
    const char* rootEnd = "</Relationships>";
    writer.write(rootEnd, strlen(rootEnd));

    // 写入到ZIP文件
    auto writeResult = zipWriter.write(getPartName(), writer.getBuffer());
    if (writeResult.isError()) {
        return Err<void>(writeResult.error().getCode(),
                       "Failed to write " + getPartName() + ": " + writeResult.error().getMessage());
    }

    return Ok();
}



TXResult<void> TXUnifiedXmlHandler::generateWorksheetRelsStream(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) const {
    TXBufferedXmlWriter writer(4096); // 4KB缓冲区

    // 写入XML声明
    const char* xmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
    writer.write(xmlDecl, strlen(xmlDecl));

    // 写入根元素开始标签
    const char* rootStart = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
    writer.write(rootStart, strlen(rootStart));

    if (index_ < context.sheets.size()) {
        const TXSheet* sheet = context.sheets[index_].get();
        bool hasCharts = sheet->getChartCount() > 0;
        bool hasPivotTables = !pivot_tables_.empty();

        u32 relationshipId = 1;

        // 先添加透视表关系（如果有透视表）
        if (hasPivotTables) {
            for (size_t i = 0; i < pivot_tables_.size(); ++i) {
                std::string pivotTableRel = "<Relationship Id=\"rId" + std::to_string(relationshipId++) +
                                           "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable\"" +
                                           " Target=\"../pivotTables/pivotTable" + std::to_string(i + 1) + ".xml\"/>";
                writer.write(pivotTableRel.c_str(), pivotTableRel.length());
            }
        }

        // 再添加绘图关系（如果有图表）
        if (hasCharts) {
            std::string drawingRel = "<Relationship Id=\"rId" + std::to_string(relationshipId) +
                                    "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\"" +
                                    " Target=\"../drawings/drawing" + std::to_string(index_ + 1) + ".xml\"/>";
            writer.write(drawingRel.c_str(), drawingRel.length());
        }
    }

    // 写入根元素结束标签
    const char* rootEnd = "</Relationships>";
    writer.write(rootEnd, strlen(rootEnd));

    // 写入到ZIP文件
    auto writeResult = zipWriter.write(getPartName(), writer.getBuffer());
    if (writeResult.isError()) {
        return Err<void>(writeResult.error().getCode(),
                       "Failed to write " + getPartName() + ": " + writeResult.error().getMessage());
    }

    return Ok();
}

TXResult<void> TXUnifiedXmlHandler::generatePivotTableRelsStream(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) const {
    TXBufferedXmlWriter writer(2048); // 2KB缓冲区，透视表关系文件很小

    // 写入XML声明
    const char* xmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
    writer.write(xmlDecl, strlen(xmlDecl));

    // 写入根元素开始标签
    const char* rootStart = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
    writer.write(rootStart, strlen(rootStart));

    // 透视表到缓存定义的关系
    int workbookRId = static_cast<int>(context.sheets.size()) + 1 + 1 + index_; // 工作表 + 样式 + 共享字符串 + 缓存ID
    std::string cacheRel = "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition\"" +
                          std::string(" Target=\"../_rels/workbook.xml.rels#rId") + std::to_string(workbookRId) + "\"/>";
    writer.write(cacheRel.c_str(), cacheRel.length());

    // 写入根元素结束标签
    const char* rootEnd = "</Relationships>";
    writer.write(rootEnd, strlen(rootEnd));

    // 写入到ZIP文件
    auto writeResult = zipWriter.write(getPartName(), writer.getBuffer());
    if (writeResult.isError()) {
        return Err<void>(writeResult.error().getCode(),
                       "Failed to write " + getPartName() + ": " + writeResult.error().getMessage());
    }

    return Ok();
}

TXResult<void> TXUnifiedXmlHandler::generatePivotCacheRelsStream(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) const {
    TXBufferedXmlWriter writer(2048); // 2KB缓冲区

    // 写入XML声明
    const char* xmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
    writer.write(xmlDecl, strlen(xmlDecl));

    // 写入根元素开始标签
    const char* rootStart = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
    writer.write(rootStart, strlen(rootStart));

    // 缓存定义到缓存记录的关系
    std::string recordsRel = "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords\"" +
                            std::string(" Target=\"pivotCacheRecords") + std::to_string(index_) + ".xml\"/>";
    writer.write(recordsRel.c_str(), recordsRel.length());

    // 写入根元素结束标签
    const char* rootEnd = "</Relationships>";
    writer.write(rootEnd, strlen(rootEnd));

    // 写入到ZIP文件
    auto writeResult = zipWriter.write(getPartName(), writer.getBuffer());
    if (writeResult.isError()) {
        return Err<void>(writeResult.error().getCode(),
                       "Failed to write " + getPartName() + ": " + writeResult.error().getMessage());
    }

    return Ok();
}

TXResult<void> TXUnifiedXmlHandler::generateDocumentPropertiesStream(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) const {
    // 生成 core.xml (核心属性)
    {
        TXBufferedXmlWriter writer(4096);

        // 写入XML声明
        const char* xmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
        writer.write(xmlDecl, strlen(xmlDecl));

        // 写入核心属性根元素
        const char* coreStart = "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
                               "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" "
                               "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">";
        writer.write(coreStart, strlen(coreStart));

        // 基本属性
        const char* creator = "<dc:creator>TinaXlsx</dc:creator>";
        writer.write(creator, strlen(creator));

        const char* lastModifiedBy = "<cp:lastModifiedBy>TinaXlsx</cp:lastModifiedBy>";
        writer.write(lastModifiedBy, strlen(lastModifiedBy));

        // 当前时间
        auto now = std::time(nullptr);
        auto tm = *std::gmtime(&now);
        std::ostringstream timeStream;
        timeStream << std::put_time(&tm, "%Y-%m-%dT%H:%M:%SZ");
        std::string timeStr = timeStream.str();

        std::string created = "<dcterms:created xsi:type=\"dcterms:W3CDTF\">" + timeStr + "</dcterms:created>";
        writer.write(created.c_str(), created.length());

        std::string modified = "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">" + timeStr + "</dcterms:modified>";
        writer.write(modified.c_str(), modified.length());

        const char* coreEnd = "</cp:coreProperties>";
        writer.write(coreEnd, strlen(coreEnd));

        // 写入到ZIP文件
        auto writeResult = zipWriter.write("docProps/core.xml", writer.getBuffer());
        if (writeResult.isError()) {
            return Err<void>(writeResult.error().getCode(),
                           "Failed to write docProps/core.xml: " + writeResult.error().getMessage());
        }
    }

    // 生成 app.xml (应用属性)
    {
        TXBufferedXmlWriter writer(4096);

        // 写入XML声明
        const char* xmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
        writer.write(xmlDecl, strlen(xmlDecl));

        // 写入应用属性根元素
        const char* appStart = "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" "
                              "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">";
        writer.write(appStart, strlen(appStart));

        const char* application = "<Application>TinaXlsx</Application>";
        writer.write(application, strlen(application));

        const char* docSecurity = "<DocSecurity>0</DocSecurity>";
        writer.write(docSecurity, strlen(docSecurity));

        const char* scaleCrop = "<ScaleCrop>false</ScaleCrop>";
        writer.write(scaleCrop, strlen(scaleCrop));

        const char* linksUpToDate = "<LinksUpToDate>false</LinksUpToDate>";
        writer.write(linksUpToDate, strlen(linksUpToDate));

        const char* sharedDoc = "<SharedDoc>false</SharedDoc>";
        writer.write(sharedDoc, strlen(sharedDoc));

        const char* hyperlinksChanged = "<HyperlinksChanged>false</HyperlinksChanged>";
        writer.write(hyperlinksChanged, strlen(hyperlinksChanged));

        const char* appVersion = "<AppVersion>16.0300</AppVersion>";
        writer.write(appVersion, strlen(appVersion));

        const char* appEnd = "</Properties>";
        writer.write(appEnd, strlen(appEnd));

        // 写入到ZIP文件
        auto writeResult = zipWriter.write("docProps/app.xml", writer.getBuffer());
        if (writeResult.isError()) {
            return Err<void>(writeResult.error().getCode(),
                           "Failed to write docProps/app.xml: " + writeResult.error().getMessage());
        }
    }

    return Ok();
}

// ==================== 具体生成方法（保留用于兼容性） ===================="

XmlNodeBuilder TXUnifiedXmlHandler::generateMainRels(const TXWorkbookContext& context) const {
    XmlNodeBuilder relationships("Relationships");
    relationships.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

    // 工作簿关系
    relationships.addChild(XmlNodeBuilder("Relationship")
                          .addAttribute("Id", "rId1")
                          .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
                          .addAttribute("Target", "xl/workbook.xml"));

    // 文档属性（如果启用）
    if (context.componentManager.hasComponent(ExcelComponent::DocumentProperties)) {
        relationships.addChild(XmlNodeBuilder("Relationship")
                              .addAttribute("Id", "rId2")
                              .addAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties")
                              .addAttribute("Target", "docProps/core.xml"));
        relationships.addChild(XmlNodeBuilder("Relationship")
                              .addAttribute("Id", "rId3")
                              .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties")
                              .addAttribute("Target", "docProps/app.xml"));
    }

    return relationships;
}

XmlNodeBuilder TXUnifiedXmlHandler::generateWorkbookRels(const TXWorkbookContext& context) const {
    XmlNodeBuilder relationships("Relationships");
    relationships.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

    u32 relationshipId = 1;

    // 工作表关系
    for (u64 i = 0; i < context.sheets.size(); ++i) {
        relationships.addChild(XmlNodeBuilder("Relationship")
                              .addAttribute("Id", "rId" + std::to_string(relationshipId++))
                              .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
                              .addAttribute("Target", "worksheets/sheet" + std::to_string(i + 1) + ".xml"));
    }

    // 样式关系（如果启用）
    if (context.componentManager.hasComponent(ExcelComponent::Styles)) {
        relationships.addChild(XmlNodeBuilder("Relationship")
                              .addAttribute("Id", "rId" + std::to_string(relationshipId++))
                              .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
                              .addAttribute("Target", "styles.xml"));
    }

    // 共享字符串关系（如果启用）
    if (context.componentManager.hasComponent(ExcelComponent::SharedStrings)) {
        relationships.addChild(XmlNodeBuilder("Relationship")
                              .addAttribute("Id", "rId" + std::to_string(relationshipId++))
                              .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings")
                              .addAttribute("Target", "sharedStrings.xml"));
    }

    // 透视表缓存关系（如果有透视表）
    if (!all_pivot_tables_.empty()) {
        u32 cacheId = 1;
        for (const auto& [sheetName, pivotTables] : all_pivot_tables_) {
            for (const auto& pivotTable : pivotTables) {
                relationships.addChild(XmlNodeBuilder("Relationship")
                                      .addAttribute("Id", "rId" + std::to_string(relationshipId++))
                                      .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition")
                                      .addAttribute("Target", "pivotCache/pivotCacheDefinition" + std::to_string(cacheId++) + ".xml"));
            }
        }
    }

    return relationships;
}

XmlNodeBuilder TXUnifiedXmlHandler::generateWorksheetRels(const TXWorkbookContext& context) const {
    XmlNodeBuilder relationships("Relationships");
    relationships.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

    if (index_ >= context.sheets.size()) {
        return relationships; // 返回空关系
    }

    const TXSheet* sheet = context.sheets[index_].get();

    // 检查是否有图表或透视表
    bool hasCharts = sheet->getChartCount() > 0;
    bool hasPivotTables = !pivot_tables_.empty();

    // 如果没有图表和透视表，返回空关系（这种情况下不应该生成文件）
    if (!hasCharts && !hasPivotTables) {
        return relationships;
    }

    u32 relationshipId = 1;

    // 先添加透视表关系（如果有透视表）
    if (hasPivotTables) {
        for (size_t i = 0; i < pivot_tables_.size(); ++i) {
            relationships.addChild(XmlNodeBuilder("Relationship")
                                  .addAttribute("Id", "rId" + std::to_string(relationshipId++))
                                  .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable")
                                  .addAttribute("Target", "../pivotTables/pivotTable" + std::to_string(i + 1) + ".xml"));
        }
    }

    // 再添加绘图关系（如果有图表）
    if (hasCharts) {
        relationships.addChild(XmlNodeBuilder("Relationship")
                              .addAttribute("Id", "rId" + std::to_string(relationshipId))
                              .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing")
                              .addAttribute("Target", "../drawings/drawing" + std::to_string(index_ + 1) + ".xml"));
    }

    return relationships;
}

TXResult<void> TXUnifiedXmlHandler::generateDocumentProperties(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) const {
    // 生成 core.xml
    XmlNodeBuilder coreProps("cp:coreProperties");
    coreProps.addAttribute("xmlns:cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
             .addAttribute("xmlns:dc", "http://purl.org/dc/elements/1.1/")
             .addAttribute("xmlns:dcterms", "http://purl.org/dc/terms/")
             .addAttribute("xmlns:dcmitype", "http://purl.org/dc/dcmitype/")
             .addAttribute("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance");

    coreProps.addChild(XmlNodeBuilder("dc:creator").setText("TinaXlsx"));
    coreProps.addChild(XmlNodeBuilder("cp:lastModifiedBy").setText("TinaXlsx"));

    // 添加创建和修改时间
    auto now = std::chrono::system_clock::now();
    auto time_t = std::chrono::system_clock::to_time_t(now);
    std::stringstream ss;
    ss << std::put_time(std::gmtime(&time_t), "%Y-%m-%dT%H:%M:%SZ");
    std::string timestamp = ss.str();

    coreProps.addChild(XmlNodeBuilder("dcterms:created")
                      .addAttribute("xsi:type", "dcterms:W3CDTF")
                      .setText(timestamp));
    coreProps.addChild(XmlNodeBuilder("dcterms:modified")
                      .addAttribute("xsi:type", "dcterms:W3CDTF")
                      .setText(timestamp));

    // 写入 core.xml
    TXXmlWriter coreWriter;
    auto setCoreRootResult = coreWriter.setRootNode(coreProps);
    if (setCoreRootResult.isError()) {
        return Err<void>(setCoreRootResult.error().getCode(), "Failed to set core root node: " + setCoreRootResult.error().getMessage());
    }

    auto coreContentResult = coreWriter.generateXmlString();
    if (coreContentResult.isError()) {
        return Err<void>(coreContentResult.error().getCode(), "Failed to generate core XML: " + coreContentResult.error().getMessage());
    }

    std::vector<uint8_t> coreData(coreContentResult.value().begin(), coreContentResult.value().end());
    auto writeCoreResult = zipWriter.write("docProps/core.xml", coreData);
    if (writeCoreResult.isError()) {
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
    appProps.addChild(XmlNodeBuilder("LinksUpToDate").setText("false"));
    appProps.addChild(XmlNodeBuilder("HyperlinksChanged").setText("false"));
    appProps.addChild(XmlNodeBuilder("AppVersion").setText("16.0300"));

    // 写入 app.xml
    TXXmlWriter appWriter;
    auto setAppRootResult = appWriter.setRootNode(appProps);
    if (setAppRootResult.isError()) {
        return Err<void>(setAppRootResult.error().getCode(), "Failed to set app root node: " + setAppRootResult.error().getMessage());
    }

    auto appContentResult = appWriter.generateXmlString();
    if (appContentResult.isError()) {
        return Err<void>(appContentResult.error().getCode(), "Failed to generate app XML: " + appContentResult.error().getMessage());
    }

    std::vector<uint8_t> appData(appContentResult.value().begin(), appContentResult.value().end());
    auto writeAppResult = zipWriter.write("docProps/app.xml", appData);
    if (writeAppResult.isError()) {
        return Err<void>(writeAppResult.error().getCode(), "Failed to write docProps/app.xml: " + writeAppResult.error().getMessage());
    }

    return Ok();
}

XmlNodeBuilder TXUnifiedXmlHandler::generatePivotTableRels(const TXWorkbookContext& context) const {
    XmlNodeBuilder relationships("Relationships");
    relationships.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

    // 透视表到缓存定义的关系
    // 注意：这里的Target应该引用工作簿关系中的缓存定义
    // 透视表通过工作簿关系引用缓存定义，关系ID计算：
    // 工作表数量 + 样式(1) + 共享字符串(1) + 透视表缓存ID
    int workbookRId = static_cast<int>(context.sheets.size()) + 1 + 1 + index_; // 工作表 + 样式 + 共享字符串 + 缓存ID
    relationships.addChild(XmlNodeBuilder("Relationship")
                          .addAttribute("Id", "rId1")
                          .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition")
                          .addAttribute("Target", "../_rels/workbook.xml.rels#rId" + std::to_string(workbookRId)));

    return relationships;
}

XmlNodeBuilder TXUnifiedXmlHandler::generatePivotCacheRels(const TXWorkbookContext& context) const {
    XmlNodeBuilder relationships("Relationships");
    relationships.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

    // 缓存定义到缓存记录的关系
    relationships.addChild(XmlNodeBuilder("Relationship")
                          .addAttribute("Id", "rId1")
                          .addAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords")
                          .addAttribute("Target", "pivotCacheRecords" + std::to_string(index_) + ".xml"));

    return relationships;
}

// ==================== 辅助方法 ====================

std::string TXUnifiedXmlHandler::getPartName() const {
    switch (type_) {
        case HandlerType::MainRels:
            return "_rels/.rels";
        case HandlerType::WorkbookRels:
            return "xl/_rels/workbook.xml.rels";
        case HandlerType::WorksheetRels:
            return "xl/worksheets/_rels/sheet" + std::to_string(index_ + 1) + ".xml.rels";
        case HandlerType::DocumentProperties:
            return "docProps/"; // 文档属性处理器处理两个文件，返回目录前缀
        case HandlerType::PivotTableRels:
            return "xl/pivotTables/_rels/pivotTable" + std::to_string(index_) + ".xml.rels";
        case HandlerType::PivotCacheRels:
            return "xl/pivotCache/_rels/pivotCacheDefinition" + std::to_string(index_) + ".xml.rels";
        default:
            return "";
    }
}

TXResult<void> TXUnifiedXmlHandler::writeXmlToZip(TXZipArchiveWriter& zipWriter, const XmlNodeBuilder& rootNode) const {
    TXXmlWriter writer;
    auto setRootResult = writer.setRootNode(rootNode);
    if (setRootResult.isError()) {
        return Err<void>(setRootResult.error().getCode(), "Failed to set root node: " + setRootResult.error().getMessage());
    }
    
    auto xmlContentResult = writer.generateXmlString();
    if (xmlContentResult.isError()) {
        return Err<void>(xmlContentResult.error().getCode(), "Failed to generate XML: " + xmlContentResult.error().getMessage());
    }
    
    std::vector<uint8_t> xmlData(xmlContentResult.value().begin(), xmlContentResult.value().end());
    auto writeResult = zipWriter.write(getPartName(), xmlData);
    if (writeResult.isError()) {
        return Err<void>(writeResult.error().getCode(), "Failed to write " + getPartName() + ": " + writeResult.error().getMessage());
    }
    
    return Ok();
}

bool TXUnifiedXmlHandler::needsPivotTableProcessing() const {
    return type_ == HandlerType::WorksheetRels || 
           type_ == HandlerType::WorkbookRels || 
           type_ == HandlerType::PivotTableRels || 
           type_ == HandlerType::PivotCacheRels;
}

// ==================== TXUnifiedXmlHandlerFactory 实现 ====================

std::unique_ptr<TXUnifiedXmlHandler> TXUnifiedXmlHandlerFactory::createMainRelsHandler() {
    return std::make_unique<TXUnifiedXmlHandler>(TXUnifiedXmlHandler::HandlerType::MainRels);
}

std::unique_ptr<TXUnifiedXmlHandler> TXUnifiedXmlHandlerFactory::createWorkbookRelsHandler() {
    return std::make_unique<TXUnifiedXmlHandler>(TXUnifiedXmlHandler::HandlerType::WorkbookRels);
}

std::unique_ptr<TXUnifiedXmlHandler> TXUnifiedXmlHandlerFactory::createWorksheetRelsHandler(u32 sheetIndex) {
    return std::make_unique<TXUnifiedXmlHandler>(TXUnifiedXmlHandler::HandlerType::WorksheetRels, sheetIndex);
}

std::unique_ptr<TXUnifiedXmlHandler> TXUnifiedXmlHandlerFactory::createDocumentPropertiesHandler() {
    return std::make_unique<TXUnifiedXmlHandler>(TXUnifiedXmlHandler::HandlerType::DocumentProperties, 0);
}

std::unique_ptr<TXUnifiedXmlHandler> TXUnifiedXmlHandlerFactory::createPivotTableRelsHandler(u32 pivotTableId) {
    return std::make_unique<TXUnifiedXmlHandler>(TXUnifiedXmlHandler::HandlerType::PivotTableRels, pivotTableId);
}

std::unique_ptr<TXUnifiedXmlHandler> TXUnifiedXmlHandlerFactory::createPivotCacheRelsHandler(u32 cacheId) {
    return std::make_unique<TXUnifiedXmlHandler>(TXUnifiedXmlHandler::HandlerType::PivotCacheRels, cacheId);
}

} // namespace TinaXlsx
