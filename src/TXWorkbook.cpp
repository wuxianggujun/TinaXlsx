#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include "TinaXlsx/TXXmlHandler.hpp"
#include "TinaXlsx/TXWorksheetWriter.hpp"
#include "TinaXlsx/TXWorksheetReader.hpp"
#include "TinaXlsx/TXComponentManager.hpp"
#include <vector>
#include <unordered_map>
#include <unordered_set>
#include <algorithm>
#include <iostream>
#include <regex>

// XML特殊字符编码
std::string encodeXmlEntities(const std::string& input) {
    std::string result = input;
    result = std::regex_replace(result, std::regex("&"), "&amp;");
    result = std::regex_replace(result, std::regex("<"), "&lt;");
    result = std::regex_replace(result, std::regex(">"), "&gt;");
    result = std::regex_replace(result, std::regex("\""), "&quot;");
    result = std::regex_replace(result, std::regex("'"), "&apos;");
    return result;
}

namespace TinaXlsx {

class TXWorkbook::Impl {
public:
    Impl() : active_sheet_index_(0), last_error_(""), auto_component_detection_(true) {
        component_manager_.registerComponent(ExcelComponent::BasicWorkbook);
    }

    ~Impl() = default;

    bool loadFromFile(const std::string& filename) {
        // 清空现有数据
        clear();
        
        TXZipArchiveReader zip;
        if (!zip.open(filename)) {
            last_error_ = "Failed to open XLSX file: " + zip.lastError();
            return false;
        }

        // 检查是否是有效的XLSX文件
        if (!zip.has("[Content_Types].xml")) {
            last_error_ = "Invalid XLSX file: missing [Content_Types].xml";
            return false;
        }

        // 读取工作簿结构（简化版）
        if (zip.has("xl/workbook.xml")) {
            std::string workbook_xml = zip.readString("xl/workbook.xml");
            if (workbook_xml.empty()) {
                last_error_ = "Failed to read workbook.xml";
                return false;
            }

            // 解析工作簿XML获取工作表信息
            if (!parseWorkbook(workbook_xml)) {
                return false;
            }
        }

        // 如果没有工作表，创建一个默认的
        if (sheets_.empty()) {
            addSheet("Sheet1");
        }

        last_error_.clear();
        return true;
    }

    bool saveToFile(const std::string& filename) {
        if (sheets_.empty()) {
            last_error_ = "No sheets to save";
            return false;
        }

        // 智能检测组件
        if (auto_component_detection_) {
            performAutoComponentDetection();
        }

        TXZipArchiveWriter zip;
        if (!zip.open(filename, false)) {
            last_error_ = "Failed to create XLSX file: " + zip.lastError();
            return false;
        }

        // 使用组件化方式写入XLSX结构
        if (!writeXlsxStructureWithComponents(zip)) {
            return false;
        }

        zip.close();
        last_error_.clear();
        return true;
    }

    TXSheet* addSheet(const std::string& name) {
        // 检查名称是否已存在
        if (hasSheet(name)) {
            last_error_ = "Sheet name already exists: " + name;
            return nullptr;
        }

        sheets_.emplace_back(std::make_unique<TXSheet>(name));
        last_error_.clear();
        return sheets_.back().get();
    }

    TXSheet* getSheet(const std::string& name) {
        for (auto& sheet : sheets_) {
            if (sheet->getName() == name) {
                return sheet.get();
            }
        }
        return nullptr;
    }

    TXSheet* getSheet(std::size_t index) {
        if (index < sheets_.size()) {
            return sheets_[index].get();
        }
        return nullptr;
    }

    bool removeSheet(const std::string& name) {
        auto it = std::find_if(sheets_.begin(), sheets_.end(),
            [&name](const std::unique_ptr<TXSheet>& sheet) {
                return sheet->getName() == name;
            });

        if (it != sheets_.end()) {
            sheets_.erase(it);
            // 调整活动工作表索引
            if (active_sheet_index_ >= sheets_.size() && !sheets_.empty()) {
                active_sheet_index_ = sheets_.size() - 1;
            }
            last_error_.clear();
            return true;
        }

        last_error_ = "Sheet not found: " + name;
        return false;
    }

    std::size_t getSheetCount() const {
        return sheets_.size();
    }

    std::vector<std::string> getSheetNames() const {
        std::vector<std::string> names;
        names.reserve(sheets_.size());
        for (const auto& sheet : sheets_) {
            names.push_back(sheet->getName());
        }
        return names;
    }

    bool hasSheet(const std::string& name) const {
        return getSheet(name) != nullptr;
    }

    bool renameSheet(const std::string& oldName, const std::string& newName) {
        if (hasSheet(newName)) {
            last_error_ = "New sheet name already exists: " + newName;
            return false;
        }

        TXSheet* sheet = getSheet(oldName);
        if (!sheet) {
            last_error_ = "Sheet not found: " + oldName;
            return false;
        }

        sheet->setName(newName);
        last_error_.clear();
        return true;
    }

    TXSheet* getActiveSheet() {
        if (active_sheet_index_ < sheets_.size()) {
            return sheets_[active_sheet_index_].get();
        }
        return nullptr;
    }

    bool setActiveSheet(const std::string& name) {
        for (std::size_t i = 0; i < sheets_.size(); ++i) {
            if (sheets_[i]->getName() == name) {
                active_sheet_index_ = i;
                last_error_.clear();
                return true;
            }
        }

        last_error_ = "Sheet not found: " + name;
        return false;
    }

    void clear() {
        sheets_.clear();
        active_sheet_index_ = 0;
        last_error_.clear();
    }

    bool isEmpty() const {
        return sheets_.empty();
    }

    const std::string& getLastError() const {
        return last_error_;
    }

    ComponentManager& getComponentManager() {
        return component_manager_;
    }

    const ComponentManager& getComponentManager() const {
        return component_manager_;
    }

    void setAutoComponentDetection(bool enable) {
        auto_component_detection_ = enable;
    }

    void registerComponent(ExcelComponent component) {
        component_manager_.registerComponent(component);
    }

private:
    TXSheet* getSheet(const std::string& name) const {
        for (const auto& sheet : sheets_) {
            if (sheet->getName() == name) {
                return sheet.get();
            }
        }
        return nullptr;
    }

    bool parseWorkbook(const std::string& workbook_xml) {
        TXXmlHandler xml;
        if (!xml.parseFromString(workbook_xml)) {
            last_error_ = "Failed to parse workbook XML: " + xml.getLastError();
            return false;
        }

        // 查找工作表信息（简化版，只提取名称）
        auto sheet_nodes = xml.findNodes("//sheet");
        for (const auto& node : sheet_nodes) {
            auto name_it = node.attributes.find("name");
            if (name_it != node.attributes.end()) {
                addSheet(name_it->second);
            }
        }

        return true;
    }

    bool writeXlsxStructureWithComponents(TXZipArchiveWriter& zip) {
        auto& components = component_manager_.getComponents();
        
        // 生成Content-Types.xml
        std::string content_types = ComponentGenerator::generateContentTypes(components, sheets_.size());
        std::vector<uint8_t> content_types_data(content_types.begin(), content_types.end());
        if (!zip.write("[Content_Types].xml", content_types_data)) {
            last_error_ = "Failed to write Content Types: " + zip.lastError();
            return false;
        }

        // 生成主关系文件
        std::string main_rels = ComponentGenerator::generateMainRelationships(components);
        std::vector<uint8_t> main_rels_data(main_rels.begin(), main_rels.end());
        if (!zip.write("_rels/.rels", main_rels_data)) {
            last_error_ = "Failed to write main relationships: " + zip.lastError();
            return false;
        }

        // 写入工作簿文件
        if (!writeWorkbook(zip)) {
            return false;
        }

        // 生成工作簿关系文件
        std::string workbook_rels = ComponentGenerator::generateWorkbookRelationships(components, sheets_.size());
        std::vector<uint8_t> workbook_rels_data(workbook_rels.begin(), workbook_rels.end());
        if (!zip.write("xl/_rels/workbook.xml.rels", workbook_rels_data)) {
            last_error_ = "Failed to write workbook relationships: " + zip.lastError();
            return false;
        }

        // 按需写入组件文件
        if (component_manager_.hasComponent(ExcelComponent::Styles)) {
            std::string styles_xml = ComponentGenerator::generateStyles();
            std::vector<uint8_t> styles_data(styles_xml.begin(), styles_xml.end());
            if (!zip.write("xl/styles.xml", styles_data)) {
                last_error_ = "Failed to write styles: " + zip.lastError();
                return false;
            }
        }

        if (component_manager_.hasComponent(ExcelComponent::SharedStrings)) {
            // 收集实际的字符串数据
            std::vector<std::string> strings = collectAllStrings();
            std::string shared_strings_xml = ComponentGenerator::generateSharedStrings(strings);
            std::vector<uint8_t> shared_strings_data(shared_strings_xml.begin(), shared_strings_xml.end());
            if (!zip.write("xl/sharedStrings.xml", shared_strings_data)) {
                last_error_ = "Failed to write shared strings: " + zip.lastError();
                return false;
            }
        }

        if (component_manager_.hasComponent(ExcelComponent::DocumentProperties)) {
            auto doc_props = ComponentGenerator::generateDocumentProperties();
            std::vector<uint8_t> core_data(doc_props.first.begin(), doc_props.first.end());
            std::vector<uint8_t> app_data(doc_props.second.begin(), doc_props.second.end());
            if (!zip.write("docProps/core.xml", core_data)) {
                last_error_ = "Failed to write core properties: " + zip.lastError();
                return false;
            }
            if (!zip.write("docProps/app.xml", app_data)) {
                last_error_ = "Failed to write app properties: " + zip.lastError();
                return false;
            }
        }

        // 写入工作表文件
        for (std::size_t i = 0; i < sheets_.size(); ++i) {
            if (!writeWorksheet(zip, i)) {
                return false;
            }
        }

        return true;
    }

    bool writeWorkbook(TXZipArchiveWriter& zip) {
        std::string workbook_xml = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>)";

        for (std::size_t i = 0; i < sheets_.size(); ++i) {
            workbook_xml += R"(<sheet name=")" + sheets_[i]->getName() + 
                          R"(" sheetId=")" + std::to_string(i + 1) + 
                          R"(" r:id="rId)" + std::to_string(i + 1) + R"("/>)";
        }

        workbook_xml += "</sheets></workbook>";

        std::vector<uint8_t> workbook_data(workbook_xml.begin(), workbook_xml.end());
        if (!zip.write("xl/workbook.xml", workbook_data)) {
            last_error_ = "Failed to write workbook: " + zip.lastError();
            return false;
        }

        return true;
    }

    bool writeWorksheet(TXZipArchiveWriter& zip, std::size_t sheet_index) {
        if (sheet_index >= sheets_.size()) {
            last_error_ = "Invalid sheet index";
            return false;
        }

        const auto& sheet = sheets_[sheet_index];
        
        // 使用新的工作表写入器
        TXWorksheetWriter writer;
        if (!writer.writeToZip(zip, sheet.get(), sheet_index + 1)) {
            last_error_ = "Failed to write worksheet: " + writer.getLastError();
            return false;
        }
        
        return true;
    }

    void performAutoComponentDetection() {
        // 重置组件，保留基础组件
        component_manager_.reset();
        
        // 检测每个工作表的功能
        for (const auto& sheet : sheets_) {
            component_manager_.autoDetectComponents(sheet.get());
        }
        
        // 如果有任何数据，注册基础样式和文档属性
        bool has_data = false;
        for (const auto& sheet : sheets_) {
            if (sheet->getUsedRange().isValid()) {
                has_data = true;
                break;
            }
        }
        
        if (has_data) {
            component_manager_.registerComponent(ExcelComponent::Styles);
            component_manager_.registerComponent(ExcelComponent::DocumentProperties);
        }
    }

    std::vector<std::string> collectAllStrings() {
        std::vector<std::string> strings;
        std::unordered_set<std::string> unique_strings;
        
        for (const auto& sheet : sheets_) {
            auto used_range = sheet->getUsedRange();
            if (used_range.isValid()) {
                for (row_t row = used_range.getStart().getRow(); row <= used_range.getEnd().getRow(); ++row) {
                    for (column_t col = used_range.getStart().getCol(); col <= used_range.getEnd().getCol(); ++col) {
                        const TXCell* cell = sheet->getCell(row, col);
                        if (cell && !cell->isEmpty()) {
                            const auto& value = cell->getValue();
                            if (std::holds_alternative<std::string>(value)) {
                                const std::string& str_value = std::get<std::string>(value);
                                if (unique_strings.find(str_value) == unique_strings.end()) {
                                    unique_strings.insert(str_value);
                                    strings.push_back(str_value);
                                }
                            }
                        }
                    }
                }
            }
        }
        
        return strings;
    }

private:
    std::vector<std::unique_ptr<TXSheet>> sheets_;
    std::size_t active_sheet_index_;
    mutable std::string last_error_;
    bool auto_component_detection_;
    ComponentManager component_manager_;
};

// TXWorkbook 实现
TXWorkbook::TXWorkbook() : pImpl(std::make_unique<Impl>()) {}

TXWorkbook::~TXWorkbook() = default;

TXWorkbook::TXWorkbook(TXWorkbook&& other) noexcept : pImpl(std::move(other.pImpl)) {}

TXWorkbook& TXWorkbook::operator=(TXWorkbook&& other) noexcept {
    if (this != &other) {
        pImpl = std::move(other.pImpl);
    }
    return *this;
}

bool TXWorkbook::loadFromFile(const std::string& filename) {
    return pImpl->loadFromFile(filename);
}

bool TXWorkbook::saveToFile(const std::string& filename) {
    return pImpl->saveToFile(filename);
}

TXSheet* TXWorkbook::addSheet(const std::string& name) {
    return pImpl->addSheet(name);
}

TXSheet* TXWorkbook::getSheet(const std::string& name) {
    return pImpl->getSheet(name);
}

TXSheet* TXWorkbook::getSheet(std::size_t index) {
    return pImpl->getSheet(index);
}

bool TXWorkbook::removeSheet(const std::string& name) {
    return pImpl->removeSheet(name);
}

std::size_t TXWorkbook::getSheetCount() const {
    return pImpl->getSheetCount();
}

std::vector<std::string> TXWorkbook::getSheetNames() const {
    return pImpl->getSheetNames();
}

bool TXWorkbook::hasSheet(const std::string& name) const {
    return pImpl->hasSheet(name);
}

bool TXWorkbook::renameSheet(const std::string& oldName, const std::string& newName) {
    return pImpl->renameSheet(oldName, newName);
}

TXSheet* TXWorkbook::getActiveSheet() {
    return pImpl->getActiveSheet();
}

bool TXWorkbook::setActiveSheet(const std::string& name) {
    return pImpl->setActiveSheet(name);
}

const std::string& TXWorkbook::getLastError() const {
    return pImpl->getLastError();
}

void TXWorkbook::clear() {
    pImpl->clear();
}

bool TXWorkbook::isEmpty() const {
    return pImpl->isEmpty();
}

ComponentManager& TXWorkbook::getComponentManager() {
    return pImpl->getComponentManager();
}

const ComponentManager& TXWorkbook::getComponentManager() const {
    return pImpl->getComponentManager();
}

void TXWorkbook::setAutoComponentDetection(bool enable) {
    pImpl->setAutoComponentDetection(enable);
}

void TXWorkbook::registerComponent(ExcelComponent component) {
    pImpl->registerComponent(component);
}

} // namespace TinaXlsx 
