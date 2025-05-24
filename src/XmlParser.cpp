/**
 * @file XmlParser.cpp
 * @brief XML解析器基础设施实现
 */

#include "TinaXlsx/XmlParser.hpp"
#include "TinaXlsx/Exception.hpp"
#include <expat.h>
#include <cstring>
#include <algorithm>
#include <cctype>
#include <sstream>

namespace TinaXlsx {

// =============================================================================
// XmlParser 基类实现
// =============================================================================

XmlParser::XmlParser() {
    initializeParser();
}

XmlParser::~XmlParser() {
    cleanupParser();
}

XmlParser::XmlParser(XmlParser&& other) noexcept
    : parser_(other.parser_), context_(std::move(other.context_)),
      startCallback_(std::move(other.startCallback_)),
      endCallback_(std::move(other.endCallback_)),
      dataCallback_(std::move(other.dataCallback_)) {
    other.parser_ = nullptr;
    
    // 重新设置用户数据指针
    if (parser_) {
        XML_SetUserData(parser_, this);
    }
}

XmlParser& XmlParser::operator=(XmlParser&& other) noexcept {
    if (this != &other) {
        cleanupParser();
        
        parser_ = other.parser_;
        context_ = std::move(other.context_);
        startCallback_ = std::move(other.startCallback_);
        endCallback_ = std::move(other.endCallback_);
        dataCallback_ = std::move(other.dataCallback_);
        
        other.parser_ = nullptr;
        
        // 重新设置用户数据指针
        if (parser_) {
            XML_SetUserData(parser_, this);
        }
    }
    return *this;
}

bool XmlParser::parse(const std::string& content, bool isFinal) {
    if (!parser_) {
        return false;
    }
    
    XML_Status status = XML_Parse(parser_, content.c_str(), 
                                  static_cast<int>(content.length()), 
                                  isFinal ? XML_TRUE : XML_FALSE);
    
    return status != XML_STATUS_ERROR;
}

void XmlParser::reset() {
    context_.reset();
    if (parser_) {
        XML_ParserReset(parser_, nullptr);
        XML_SetUserData(parser_, this);
        XML_SetElementHandler(parser_, startElementHandler, endElementHandler);
        XML_SetCharacterDataHandler(parser_, characterDataHandler);
    }
}

void XmlParser::initializeParser() {
    parser_ = XML_ParserCreate(nullptr);
    if (parser_) {
        XML_SetUserData(parser_, this);
        XML_SetElementHandler(parser_, startElementHandler, endElementHandler);
        XML_SetCharacterDataHandler(parser_, characterDataHandler);
    }
}

void XmlParser::cleanupParser() {
    if (parser_) {
        XML_ParserFree(parser_);
        parser_ = nullptr;
    }
}

void XmlParser::startElementHandler(void* userData, const XML_Char* name, const XML_Char** atts) {
    auto* parser = static_cast<XmlParser*>(userData);
    if (!parser || !parser->startCallback_) {
        return;
    }
    
    // 转换属性为vector
    std::vector<std::pair<std::string, std::string>> attributes;
    if (atts) {
        for (int i = 0; atts[i]; i += 2) {
            attributes.emplace_back(atts[i], atts[i + 1]);
        }
    }
    
    parser->startCallback_(name, attributes);
}

void XmlParser::endElementHandler(void* userData, const XML_Char* name) {
    auto* parser = static_cast<XmlParser*>(userData);
    if (!parser || !parser->endCallback_) {
        return;
    }
    
    parser->endCallback_(name);
}

void XmlParser::characterDataHandler(void* userData, const XML_Char* s, int len) {
    auto* parser = static_cast<XmlParser*>(userData);
    if (!parser || !parser->dataCallback_) {
        return;
    }
    
    parser->dataCallback_(std::string(s, len));
}

// =============================================================================
// WorkbookXmlParser 实现
// =============================================================================

WorkbookXmlParser::WorkbookXmlParser() {
    setElementStartCallback([this](const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes) {
        handleElementStart(name, attributes);
    });
    
    setElementEndCallback([this](const std::string& name) {
        handleElementEnd(name);
    });
}

std::vector<WorkbookXmlParser::SheetInfo> WorkbookXmlParser::parseWorkbook(const std::string& content) {
    sheets_.clear();
    reset();
    
    if (!parse(content)) {
        throw ParseException("Failed to parse workbook XML");
    }
    
    return sheets_;
}

void WorkbookXmlParser::handleElementStart(const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes) {
    if (name == "sheet") {
        SheetInfo sheet;
        
        for (const auto& [attrName, attrValue] : attributes) {
            if (attrName == "name") {
                sheet.name = attrValue;
            } else if (attrName == "r:id") {
                sheet.relationId = attrValue;
            } else if (attrName == "sheetId") {
                sheet.sheetId = std::strtoul(attrValue.c_str(), nullptr, 10);
            }
        }
        
        if (!sheet.name.empty() && !sheet.relationId.empty()) {
            sheets_.push_back(sheet);
        }
    }
}

void WorkbookXmlParser::handleElementEnd(const std::string& name) {
    // 工作簿解析通常不需要处理元素结束
}

// =============================================================================
// SharedStringsXmlParser 实现
// =============================================================================

SharedStringsXmlParser::SharedStringsXmlParser() {
    setElementStartCallback([this](const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes) {
        handleElementStart(name, attributes);
    });
    
    setElementEndCallback([this](const std::string& name) {
        handleElementEnd(name);
    });
    
    setCharacterDataCallback([this](const std::string& data) {
        handleCharacterData(data);
    });
}

std::vector<std::string> SharedStringsXmlParser::parseSharedStrings(const std::string& content) {
    sharedStrings_.clear();
    currentString_.clear();
    reset();
    
    if (!parse(content)) {
        throw ParseException("Failed to parse shared strings XML");
    }
    
    return sharedStrings_;
}

void SharedStringsXmlParser::handleElementStart(const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes) {
    if (name == "si") {
        context_.state = XmlParseState::SharedString;
        currentString_.clear();
    } else if (name == "t" && context_.state == XmlParseState::SharedString) {
        context_.state = XmlParseState::Value;
    }
}

void SharedStringsXmlParser::handleElementEnd(const std::string& name) {
    if (name == "si") {
        sharedStrings_.push_back(currentString_);
        currentString_.clear();
        context_.state = XmlParseState::None;
    } else if (name == "t") {
        context_.state = XmlParseState::SharedString;
    }
}

void SharedStringsXmlParser::handleCharacterData(const std::string& data) {
    if (context_.state == XmlParseState::Value) {
        currentString_ += data;
    }
}

// =============================================================================
// RelationshipsXmlParser 实现
// =============================================================================

RelationshipsXmlParser::RelationshipsXmlParser() {
    setElementStartCallback([this](const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes) {
        handleElementStart(name, attributes);
    });
}

std::vector<RelationshipsXmlParser::Relationship> RelationshipsXmlParser::parseRelationships(const std::string& content) {
    relationships_.clear();
    reset();
    
    if (!parse(content)) {
        throw ParseException("Failed to parse relationships XML");
    }
    
    return relationships_;
}

void RelationshipsXmlParser::handleElementStart(const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes) {
    if (name == "Relationship") {
        Relationship rel;
        
        for (const auto& [attrName, attrValue] : attributes) {
            if (attrName == "Id") {
                rel.id = attrValue;
            } else if (attrName == "Type") {
                rel.type = attrValue;
            } else if (attrName == "Target") {
                rel.target = attrValue;
            }
        }
        
        if (!rel.id.empty() && !rel.target.empty()) {
            relationships_.push_back(rel);
        }
    }
}

// =============================================================================
// WorksheetXmlParser 实现
// =============================================================================

WorksheetXmlParser::WorksheetXmlParser() {
    setElementStartCallback([this](const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes) {
        handleElementStart(name, attributes);
    });
    
    setElementEndCallback([this](const std::string& name) {
        handleElementEnd(name);
    });
    
    setCharacterDataCallback([this](const std::string& data) {
        handleCharacterData(data);
    });
}

void WorksheetXmlParser::parseWorksheet(const std::string& content) {
    reset();
    currentRow_.clear();
    currentRowIndex_ = 0;
    
    if (!parse(content)) {
        throw ParseException("Failed to parse worksheet XML");
    }
}

void WorksheetXmlParser::handleElementStart(const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes) {
    if (name == "sheetData") {
        context_.state = XmlParseState::SheetData;
    } else if (name == "row" && context_.state == XmlParseState::SheetData) {
        context_.state = XmlParseState::Row;
        currentRow_.clear();
        
        // 查找行号
        for (const auto& [attrName, attrValue] : attributes) {
            if (attrName == "r") {
                currentRowIndex_ = std::strtoul(attrValue.c_str(), nullptr, 10) - 1; // 转为0基索引
                break;
            }
        }
    } else if (name == "c" && context_.state == XmlParseState::Row) {
        context_.state = XmlParseState::Cell;
        context_.currentValue.clear();
        
        // 解析单元格属性
        std::string cellRef;
        std::string cellType;
        
        for (const auto& [attrName, attrValue] : attributes) {
            if (attrName == "r") {
                cellRef = attrValue;
            } else if (attrName == "t") {
                cellType = attrValue;
            }
        }
        
        // 这里可以存储单元格信息以供后续使用
        context_.currentElementName = cellRef + "|" + cellType;
        
    } else if (name == "v" && context_.state == XmlParseState::Cell) {
        context_.state = XmlParseState::Value;
        context_.currentValue.clear();
    }
}

void WorksheetXmlParser::handleElementEnd(const std::string& name) {
    if (name == "sheetData") {
        context_.state = XmlParseState::None;
    } else if (name == "row") {
        // 行结束，调用行回调
        if (rowCallback_) {
            rowCallback_(currentRowIndex_, currentRow_);
        }
        context_.state = XmlParseState::SheetData;
    } else if (name == "c") {
        // 单元格结束，处理值
        if (!context_.currentElementName.empty()) {
            size_t pos = context_.currentElementName.find('|');
            std::string cellRef = context_.currentElementName.substr(0, pos);
            std::string cellType = (pos != std::string::npos) ? context_.currentElementName.substr(pos + 1) : "";
            
            CellPosition position = parseCellReference(cellRef);
            CellValue value;
            
            if (cellType == "s") {
                // 共享字符串
                value = parseSharedString(context_.currentValue);
            } else {
                // 普通值
                value = context_.currentValue.empty() ? CellValue(std::monostate{}) : CellValue(context_.currentValue);
            }
            
            // 确保行有足够的列
            while (currentRow_.size() <= position.column) {
                currentRow_.push_back(CellValue(std::monostate{}));
            }
            currentRow_[position.column] = value;
            
            // 调用单元格回调
            if (cellCallback_) {
                cellCallback_(position, value);
            }
        }
        
        context_.state = XmlParseState::Row;
    } else if (name == "v") {
        context_.state = XmlParseState::Cell;
    }
}

void WorksheetXmlParser::handleCharacterData(const std::string& data) {
    if (context_.state == XmlParseState::Value) {
        context_.currentValue += data;
    }
}

CellPosition WorksheetXmlParser::parseCellReference(const std::string& cellRef) {
    if (cellRef.empty()) {
        return {0, 0};
    }
    
    size_t i = 0;
    ColumnIndex col = 0;
    
    // 解析列部分（字母）
    while (i < cellRef.length() && std::isalpha(cellRef[i])) {
        col = col * 26 + (std::toupper(cellRef[i]) - 'A' + 1);
        i++;
    }
    col--; // 转换为0基于的索引
    
    // 解析行部分（数字）
    RowIndex row = 0;
    if (i < cellRef.length()) {
        row = std::strtoul(cellRef.c_str() + i, nullptr, 10);
        if (row > 0) row--; // 转换为0基于的索引
    }
    
    return {row, col};
}

CellValue WorksheetXmlParser::parseSharedString(const std::string& index) {
    if (!sharedStrings_ || index.empty()) {
        return CellValue(std::monostate{});
    }
    
    try {
        size_t idx = std::stoul(index);
        if (idx < sharedStrings_->size()) {
            return CellValue((*sharedStrings_)[idx]);
        }
    } catch (...) {
        // 索引无效，返回原始值
    }
    
    return CellValue(index);
}

} // namespace TinaXlsx 