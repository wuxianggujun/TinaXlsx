#include "TinaXlsx/TXComponentManager.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include <algorithm>

namespace TinaXlsx {

// ==================== ComponentManager 实现 ====================

void ComponentManager::registerComponent(ExcelComponent component) {
    registered_components_.insert(component);
}

bool ComponentManager::hasComponent(ExcelComponent component) const {
    return registered_components_.find(component) != registered_components_.end();
}

const std::unordered_set<ExcelComponent>& ComponentManager::getComponents() const {
    return registered_components_;
}

void ComponentManager::autoDetectComponents(const TXSheet* sheet) {
    if (!sheet) return;
    
    // 检测合并单元格
    if (sheet->getMergeCount() > 0) {
        registerComponent(ExcelComponent::MergedCells);
    }
    
    // 检测工作表中的功能
    auto used_range = sheet->getUsedRange();
    if (used_range.isValid()) {
        // 遍历使用的范围检测功能
        for (row_t row = used_range.getStart().getRow(); row <= used_range.getEnd().getRow(); ++row) {
            for (column_t col = used_range.getStart().getCol(); col <= used_range.getEnd().getCol(); ++col) {
                const TXCell* cell = sheet->getCell(row, col);
                if (cell && !cell->isEmpty()) {
                    // 检测公式
                    if (cell->isFormula()) {
                        registerComponent(ExcelComponent::Formulas);
                    }
                    
                    // 检测字符串值（用于共享字符串）
                    auto value = cell->getValue();
                    if (std::holds_alternative<std::string>(value)) {
                        registerComponent(ExcelComponent::SharedStrings);
                    }
                    
                    // 注意：样式组件将在TXWorkbook中根据需要注册
                }
            }
        }
    }
}

void ComponentManager::reset() {
    registered_components_.clear();
    // 基础工作簿组件始终存在
    registerComponent(ExcelComponent::BasicWorkbook);
}

// ==================== ComponentGenerator 实现 ====================

std::string ComponentGenerator::generateContentTypes(const std::unordered_set<ExcelComponent>& components, size_t sheet_count) {
    std::string content_types = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>)";

    // 根据组件添加相应的ContentType
    if (components.find(ExcelComponent::Styles) != components.end()) {
        content_types += R"(
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>)";
    }
    
    if (components.find(ExcelComponent::SharedStrings) != components.end()) {
        content_types += R"(
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>)";
    }
    
    if (components.find(ExcelComponent::DocumentProperties) != components.end()) {
        content_types += R"(
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>)";
    }

    // 添加工作表
    for (std::size_t i = 0; i < sheet_count; ++i) {
        content_types += R"(
<Override PartName="/xl/worksheets/sheet)" + std::to_string(i + 1) + 
                       R"(.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>)";
    }

    content_types += "\n</Types>";
    return content_types;
}

std::string ComponentGenerator::generateMainRelationships(const std::unordered_set<ExcelComponent>& components) {
    std::string main_rels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>)";

    int rel_id = 2;
    if (components.find(ExcelComponent::DocumentProperties) != components.end()) {
        main_rels += R"(
<Relationship Id="rId)" + std::to_string(rel_id) + R"(" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>)";
        rel_id++;
        main_rels += R"(
<Relationship Id="rId)" + std::to_string(rel_id) + R"(" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>)";
        rel_id++;
    }

    main_rels += "\n</Relationships>";
    return main_rels;
}

std::string ComponentGenerator::generateWorkbookRelationships(const std::unordered_set<ExcelComponent>& components, size_t sheet_count) {
    std::string workbook_rels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)";

    int rel_id = 1;
    
    if (components.find(ExcelComponent::Styles) != components.end()) {
        workbook_rels += R"(
<Relationship Id="rId)" + std::to_string(rel_id) + R"(" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>)";
        rel_id++;
    }
    
    if (components.find(ExcelComponent::SharedStrings) != components.end()) {
        workbook_rels += R"(
<Relationship Id="rId)" + std::to_string(rel_id) + R"(" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>)";
        rel_id++;
    }

    // 添加工作表关系
    for (std::size_t i = 0; i < sheet_count; ++i) {
        workbook_rels += R"(
<Relationship Id="rId)" + std::to_string(rel_id) + R"(" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet)" + 
                       std::to_string(i + 1) + R"(.xml"/>)";
        rel_id++;
    }

    workbook_rels += "\n</Relationships>";
    return workbook_rels;
}

std::string ComponentGenerator::generateStyles() {
    return R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="0"/>
<fonts count="1">
<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
</fonts>
<fills count="2">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
</fills>
<borders count="1">
<border><left/><right/><top/><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
</styleSheet>)";
}

std::string ComponentGenerator::generateSharedStrings(const std::vector<std::string>& strings) {
    if (strings.empty()) {
        return R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>)";
    }
    
    std::string shared_strings = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count=")" + 
                               std::to_string(strings.size()) + R"(" uniqueCount=")" + 
                               std::to_string(strings.size()) + R"(">)";
    
    for (const auto& str : strings) {
        shared_strings += R"(
<si><t>)" + str + R"(</t></si>)";
    }
    
    shared_strings += "\n</sst>";
    return shared_strings;
}

std::pair<std::string, std::string> ComponentGenerator::generateDocumentProperties() {
    std::string core_xml = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>TinaXlsx</dc:creator>
<cp:lastModifiedBy>TinaXlsx</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">2025-01-27T00:00:00Z</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">2025-01-27T00:00:00Z</dcterms:modified>
</cp:coreProperties>)";

    std::string app_xml = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>TinaXlsx</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>16.0300</AppVersion>
</Properties>)";

    return std::make_pair(core_xml, app_xml);
}

} // namespace TinaXlsx 