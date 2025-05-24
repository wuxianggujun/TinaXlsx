/**
 * @file ExcelStructureManager.cpp
 * @brief Excel文件结构管理器实现
 */

#include "TinaXlsx/ExcelStructureManager.hpp"
#include "TinaXlsx/Exception.hpp"

namespace TinaXlsx {

ExcelStructureManager::ExcelStructureManager(std::unique_ptr<ExcelZipReader> zipReader)
    : zipReader_(std::move(zipReader)) {
    
    if (!zipReader_) {
        throw InvalidArgumentException("ZIP reader cannot be null");
    }
    
    if (!zipReader_->isValidExcelFile()) {
        throw FileException("Invalid Excel file format");
    }
}

void ExcelStructureManager::parseStructure() {
    if (initialized_) {
        return;
    }
    
    try {
        parseSharedStrings();    // 首先解析共享字符串
        parseRelationships();    // 然后解析关系
        parseWorkbook();         // 最后解析工作簿信息
        
        initialized_ = true;
    } catch (const std::exception& e) {
        throw ParseException("Failed to parse Excel structure: " + std::string(e.what()));
    }
}

void ExcelStructureManager::parseWorkbook() {
    auto workbookContent = zipReader_->readWorkbook();
    if (!workbookContent) {
        throw ParseException("Failed to read workbook.xml");
    }
    
    WorkbookXmlParser parser;
    auto sheetsInfo = parser.parseWorkbook(*workbookContent);
    
    sheets_.clear();
    for (const auto& sheetXmlInfo : sheetsInfo) {
        SheetInfo sheet;
        sheet.name = sheetXmlInfo.name;
        sheet.relationId = sheetXmlInfo.relationId;
        sheet.sheetId = sheetXmlInfo.sheetId;
        
        // 从关系中查找对应的文件路径
        auto it = relationships_.find(sheet.relationId);
        if (it != relationships_.end()) {
            sheet.filePath = it->second;
        } else {
            // 如果没有找到关系，尝试使用默认路径
            sheet.filePath = "xl/worksheets/sheet" + std::to_string(sheet.sheetId) + ".xml";
        }
        
        sheets_.push_back(sheet);
    }
}

void ExcelStructureManager::parseRelationships() {
    auto relsContent = zipReader_->readWorkbookRelationships();
    if (!relsContent) {
        // 关系文件不存在是可以的，某些简单的Excel文件可能没有
        return;
    }
    
    RelationshipsXmlParser parser;
    auto relationships = parser.parseRelationships(*relsContent);
    
    relationships_.clear();
    for (const auto& rel : relationships) {
        // 只关心工作表关系
        if (rel.type.find("worksheet") != std::string::npos) {
            std::string target = rel.target;
            // 如果是相对路径，添加xl/前缀
            if (!target.empty() && target[0] != '/') {
                target = "xl/" + target;
            } else if (!target.empty() && target[0] == '/') {
                target = target.substr(1); // 移除开头的斜杠
            }
            relationships_[rel.id] = target;
        }
    }
}

void ExcelStructureManager::parseSharedStrings() {
    auto sharedStringsContent = zipReader_->readSharedStrings();
    if (!sharedStringsContent) {
        // 共享字符串文件不存在是可以的
        sharedStrings_.clear();
        return;
    }
    
    SharedStringsXmlParser parser;
    sharedStrings_ = parser.parseSharedStrings(*sharedStringsContent);
}

ExcelZipReader* ExcelStructureManager::getZipReader() const {
    return zipReader_.get();
}

const std::vector<std::string>& ExcelStructureManager::getSharedStrings() const {
    return sharedStrings_;
}

std::vector<std::string> ExcelStructureManager::getSheetNames() const {
    ensureInitialized();
    std::vector<std::string> names;
    names.reserve(sheets_.size());
    for (const auto& sheet : sheets_) {
        names.push_back(sheet.name);
    }
    return names;
}

std::optional<ExcelStructureManager::SheetInfo> ExcelStructureManager::findSheetByName(const std::string& name) const {
    ensureInitialized();
    for (const auto& sheet : sheets_) {
        if (sheet.name == name) {
            return sheet;
        }
    }
    return std::nullopt;
}

std::optional<ExcelStructureManager::SheetInfo> ExcelStructureManager::getSheetByIndex(size_t index) const {
    ensureInitialized();
    if (index < sheets_.size()) {
        return sheets_[index];
    }
    return std::nullopt;
}

void ExcelStructureManager::ensureInitialized() const {
    if (!initialized_) {
        const_cast<ExcelStructureManager*>(this)->parseStructure();
    }
}

const std::vector<ExcelStructureManager::SheetInfo>& ExcelStructureManager::getSheets() const {
    ensureInitialized();
    return sheets_;
}

size_t ExcelStructureManager::getSheetCount() const {
    ensureInitialized();
    return sheets_.size();
}

bool ExcelStructureManager::isValidExcelFile() const {
    return zipReader_ && zipReader_->isValidExcelFile();
}

} // namespace TinaXlsx 