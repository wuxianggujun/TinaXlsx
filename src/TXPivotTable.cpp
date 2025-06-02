/*
 * TinaXlsx - Modern C++ Excel File Processing Library
 * 
 * Copyright (c) 2025 wuxianggujun
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

#include "TinaXlsx/TXPivotTable.hpp"
#include "TinaXlsx/TXPivotTableXmlHandler.hpp"
#include <algorithm>
#include <stdexcept>

namespace TinaXlsx {

    // ==================== TXPivotField 实现 ====================

    TXPivotField::TXPivotField(const std::string& name, PivotFieldType type)
        : name_(name)
        , type_(type)
        , aggregateFunction_(PivotAggregateFunction::Sum)
        , sortAscending_(true) {
        if (name.empty()) {
            throw std::invalid_argument("Field name cannot be empty");
        }
    }

    // ==================== TXPivotCache 实现 ====================

    TXPivotCache::TXPivotCache(const TXRange& sourceRange)
        : sourceRange_(sourceRange)
        , needsRefresh_(true)
        , sourceSheet_(nullptr) {
        // TODO: 验证源数据范围的有效性
    }

    void TXPivotCache::refresh() {
        // 如果字段名称已经手动设置，不要重新刷新
        if (!needsRefresh_) {
            return;
        }

        // 实现数据刷新逻辑
        // 1. 从源数据范围读取数据
        // 2. 解析字段名称
        // 3. 构建数据缓存

        fieldNames_.clear();

        // TODO: 这里应该从实际的工作表中读取数据
        // 目前的实现是临时的，实际应该：
        // 1. 获取工作表引用
        // 2. 从sourceRange_的第一行读取字段名称
        // 3. 从数据范围读取实际数据

        if (sourceSheet_) {
            // 从实际工作表读取字段名称（第一行）
            auto start = sourceRange_.getStart();
            auto end = sourceRange_.getEnd();
            row_t startRow = start.getRow();
            column_t startCol = start.getCol();
            column_t endCol = end.getCol();

            for (column_t col = startCol; col <= endCol; ++col) {
                auto cellValue = sourceSheet_->getCellValue(startRow, col);
                std::string fieldName;

                // 从CellValue中提取字符串
                if (std::holds_alternative<std::string>(cellValue)) {
                    fieldName = std::get<std::string>(cellValue);
                } else if (std::holds_alternative<f64>(cellValue)) {
                    fieldName = std::to_string(std::get<f64>(cellValue));
                } else if (std::holds_alternative<i64>(cellValue)) {
                    fieldName = std::to_string(std::get<i64>(cellValue));
                } else if (std::holds_alternative<bool>(cellValue)) {
                    fieldName = std::get<bool>(cellValue) ? "TRUE" : "FALSE";
                } else {
                    fieldName = "";
                }

                if (fieldName.empty()) {
                    fieldName = "字段" + std::to_string(static_cast<int>(col.index()) - static_cast<int>(startCol.index()) + 1);
                }
                fieldNames_.push_back(fieldName);
            }
        } else {
            // 如果没有工作表引用，生成通用字段名称
            auto start = sourceRange_.getStart();
            auto end = sourceRange_.getEnd();
            column_t columnCount = end.getCol() - start.getCol() + 1;

            for (int i = 0; i < static_cast<int>(columnCount.index()); ++i) {
                fieldNames_.push_back("字段" + std::to_string(i + 1));
            }
        }

        needsRefresh_ = false;
    }

    std::vector<std::string> TXPivotCache::getFieldNames() const {
        return fieldNames_;
    }

    void TXPivotCache::setSourceSheet(const TXSheet* sheet) {
        sourceSheet_ = sheet;
        needsRefresh_ = true;  // 需要重新刷新数据
    }

    void TXPivotCache::setFieldNames(const std::vector<std::string>& fieldNames) {
        fieldNames_ = fieldNames;
        needsRefresh_ = false;  // 手动设置字段名称，不需要刷新
    }

    // ==================== TXPivotTable 实现 ====================

    TXPivotTable::TXPivotTable(const TXRange& sourceRange, const std::string& targetCell)
        : name_("PivotTable1")
        , targetCell_(targetCell)
        , cache_(std::make_unique<TXPivotCache>(sourceRange)) {
        if (targetCell.empty()) {
            throw std::invalid_argument("Target cell cannot be empty");
        }
    }

    bool TXPivotTable::addRowField(const std::string& fieldName) {
        if (!validateField(fieldName)) {
            return false;
        }

        auto field = std::make_shared<TXPivotField>(fieldName, PivotFieldType::Row);
        fields_.push_back(field);
        fieldMap_[fieldName] = field;
        return true;
    }

    bool TXPivotTable::addColumnField(const std::string& fieldName) {
        if (!validateField(fieldName)) {
            return false;
        }

        auto field = std::make_shared<TXPivotField>(fieldName, PivotFieldType::Column);
        fields_.push_back(field);
        fieldMap_[fieldName] = field;
        return true;
    }

    bool TXPivotTable::addDataField(const std::string& fieldName, PivotAggregateFunction aggregateFunc) {
        if (!validateField(fieldName)) {
            return false;
        }

        auto field = std::make_shared<TXPivotField>(fieldName, PivotFieldType::Data);
        field->setAggregateFunction(aggregateFunc);
        fields_.push_back(field);
        fieldMap_[fieldName] = field;
        return true;
    }

    bool TXPivotTable::addFilterField(const std::string& fieldName) {
        if (!validateField(fieldName)) {
            return false;
        }

        auto field = std::make_shared<TXPivotField>(fieldName, PivotFieldType::Filter);
        fields_.push_back(field);
        fieldMap_[fieldName] = field;
        return true;
    }

    bool TXPivotTable::removeField(const std::string& fieldName) {
        auto it = fieldMap_.find(fieldName);
        if (it == fieldMap_.end()) {
            return false;
        }

        // 从字段列表中移除
        auto fieldIt = std::find_if(fields_.begin(), fields_.end(),
            [&fieldName](const std::shared_ptr<TXPivotField>& field) {
                return field->getName() == fieldName;
            });

        if (fieldIt != fields_.end()) {
            fields_.erase(fieldIt);
        }

        // 从映射中移除
        fieldMap_.erase(it);
        return true;
    }

    std::shared_ptr<TXPivotField> TXPivotTable::getField(const std::string& fieldName) const {
        auto it = fieldMap_.find(fieldName);
        return (it != fieldMap_.end()) ? it->second : nullptr;
    }

    std::vector<std::shared_ptr<TXPivotField>> TXPivotTable::getFieldsByType(PivotFieldType type) const {
        std::vector<std::shared_ptr<TXPivotField>> result;
        for (const auto& field : fields_) {
            if (field->getType() == type) {
                result.push_back(field);
            }
        }
        return result;
    }

    bool TXPivotTable::generate() {
        try {
            // 1. 刷新数据缓存
            cache_->refresh();

            // 2. 验证字段配置
            if (fields_.empty()) {
                throw std::runtime_error("No fields configured for pivot table");
            }

            // 3. 计算聚合数据
            calculateAggregates();

            // 4. 生成透视表XML
            generatePivotTableXML();

            return true;
        }
        catch (const std::exception&) {
            // TODO: 记录错误日志
            return false;
        }
    }

    bool TXPivotTable::refresh() {
        // 刷新数据并重新生成透视表
        return generate();
    }

    bool TXPivotTable::integrateToWorkbook(TXZipArchiveWriter& zipWriter, const std::string& sheetName) {
        try {
            // 1. 生成透视表XML文件
            TXPivotTableXmlHandler pivotHandler(this, 1);

            // 创建临时的空引用（用于测试）
            std::vector<std::unique_ptr<TXSheet>> emptySheets;
            TXStyleManager emptyStyleManager;
            ComponentManager emptyComponentManager;
            TXSharedStringsPool emptyStringsPool;
            TXWorkbookProtectionManager emptyProtectionManager;

            TXWorkbookContext context(emptySheets, emptyStyleManager, emptyComponentManager,
                                    emptyStringsPool, emptyProtectionManager);

            auto pivotResult = pivotHandler.save(zipWriter, context);
            if (pivotResult.isError()) {
                return false;
            }

            // 2. 生成透视表缓存XML文件
            TXPivotCacheXmlHandler cacheHandler(this, 1);
            auto cacheResult = cacheHandler.save(zipWriter, context);
            if (cacheResult.isError()) {
                return false;
            }

            // 3. 生成透视表缓存记录XML文件
            std::string cacheRecordsXml = generatePivotCacheRecordsXml();
            std::vector<uint8_t> cacheRecordsData(cacheRecordsXml.begin(), cacheRecordsXml.end());
            auto recordsResult = zipWriter.write("xl/pivotCache/pivotCacheRecords1.xml", cacheRecordsData);
            if (recordsResult.isError()) {
                return false;
            }

            // 4. 更新工作表关系文件
            std::string sheetRelsXml = generateSheetRelationshipsXml();
            std::vector<uint8_t> sheetRelsData(sheetRelsXml.begin(), sheetRelsXml.end());
            auto sheetRelsResult = zipWriter.write("xl/worksheets/_rels/sheet1.xml.rels", sheetRelsData);
            if (sheetRelsResult.isError()) {
                return false;
            }

            // 5. 更新工作簿关系文件
            std::string workbookRelsXml = generateWorkbookRelationshipsXml();
            std::vector<uint8_t> workbookRelsData(workbookRelsXml.begin(), workbookRelsXml.end());
            auto workbookRelsResult = zipWriter.write("xl/_rels/workbook.xml.rels", workbookRelsData);
            if (workbookRelsResult.isError()) {
                return false;
            }

            return true;
        }
        catch (const std::exception&) {
            return false;
        }
    }

    bool TXPivotTable::validateField(const std::string& fieldName) const {
        if (fieldName.empty()) {
            return false;
        }

        // 检查字段是否已存在
        if (fieldMap_.find(fieldName) != fieldMap_.end()) {
            return false;
        }

        // TODO: 实际的库应该提供API让用户设置可用字段，或者从实际工作表数据中读取
        // 目前为了让库能正常工作，我们允许用户添加任意字段名称
        // 在实际使用中，应该验证字段是否在数据源中存在

        return true;  // 暂时允许任意字段名称
    }

    void TXPivotTable::generatePivotTableXML() {
        // 实现透视表XML生成
        // 1. 创建透视表定义XML
        generatePivotTableDefinitionXML();

        // 2. 创建透视表缓存XML
        generatePivotCacheDefinitionXML();

        // 3. 更新工作表关系
        updateWorksheetRelationships();
    }

    void TXPivotTable::calculateAggregates() {
        // 实现聚合计算
        // 1. 根据行字段和列字段分组数据
        // 2. 对每个分组应用聚合函数
        // 3. 构建结果矩阵

        // 这里是聚合计算的核心逻辑
        // 实际实现需要读取源数据并进行分组聚合
    }

    void TXPivotTable::generatePivotTableDefinitionXML() {
        // 生成透视表定义XML
        // 这个XML定义了透视表的结构、字段配置等

        // 实际实现需要生成类似这样的XML结构：
        /*
        <pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                             name="PivotTable1" cacheId="1" applyNumberFormats="0"
                             applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0"
                             applyAlignmentFormats="0" applyWidthHeightFormats="1"
                             dataCaption="数据" updatedVersion="6" minRefreshableVersion="6"
                             useAutoFormatting="1" itemPrintTitles="1" createdVersion="6"
                             indent="0" outline="1" outlineData="1" multipleFieldFilters="0"
                             compact="0" compactData="0">
            <location ref="F1:H15" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/>
            <pivotFields count="5">
                <!-- 字段定义 -->
            </pivotFields>
            <rowFields count="1">
                <field x="0"/>
            </rowFields>
            <colFields count="1">
                <field x="2"/>
            </colFields>
            <dataFields count="1">
                <dataField name="求和项:销售额" fld="3" subtotal="sum"/>
            </dataFields>
        </pivotTableDefinition>
        */
    }

    void TXPivotTable::generatePivotCacheDefinitionXML() {
        // 生成透视表缓存定义XML
        // 这个XML定义了数据源和缓存信息

        // 实际实现需要生成类似这样的XML结构：
        /*
        <pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                             r:id="rId1" invalid="0" refreshOnLoad="1" refreshedBy="User"
                             refreshedDate="44927.123456" createdVersion="6" refreshedVersion="6"
                             minRefreshableVersion="6" recordCount="10">
            <cacheSource type="worksheet">
                <worksheetSource ref="A1:E11" sheet="销售数据"/>
            </cacheSource>
            <cacheFields count="5">
                <cacheField name="产品类别" numFmtId="0">
                    <sharedItems count="3">
                        <s v="电子产品"/>
                        <s v="服装"/>
                        <s v="家具"/>
                    </sharedItems>
                </cacheField>
                <!-- 更多字段定义 -->
            </cacheFields>
        </pivotCacheDefinition>
        */
    }

    void TXPivotTable::updateWorksheetRelationships() {
        // 更新工作表关系
        // 需要在工作表XML中添加透视表引用
        // 并更新关系文件

        // 实际实现需要：
        // 1. 在工作表XML中添加 <pivotTables> 节点
        // 2. 更新 _rels/sheet1.xml.rels 文件
        // 3. 添加透视表文件到ZIP包中
    }

    std::string TXPivotTable::generatePivotCacheRecordsXml() const {
        // TODO: 实现从实际工作表数据中读取并生成缓存记录
        // 当前返回空的缓存记录，让Excel从数据源重新计算
        return R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0">
</pivotCacheRecords>)";
    }

    std::string TXPivotTable::generateSheetRelationshipsXml() const {
        return R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>)";
    }

    std::string TXPivotTable::generateWorkbookRelationshipsXml() const {
        return R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>)";
    }

} // namespace TinaXlsx
