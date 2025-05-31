//
// @file TXXmlReader.cpp
// @brief XML 读取器实现
//

#include "TinaXlsx/TXXmlReader.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include <pugixml.hpp>
#include <sstream>

namespace TinaXlsx {

TXXmlReader::TXXmlReader() : doc_(std::make_unique<pugi::xml_document>()) {}

TXXmlReader::TXXmlReader(const XmlParseOptions& options) 
    : doc_(std::make_unique<pugi::xml_document>()), options_(options) {}

TXXmlReader::~TXXmlReader() = default;

TXXmlReader::TXXmlReader(TXXmlReader&& other) noexcept 
    : doc_(std::move(other.doc_)), options_(other.options_), isValid_(other.isValid_) {
    other.isValid_ = false;
}

TXXmlReader& TXXmlReader::operator=(TXXmlReader&& other) noexcept {
    if (this != &other) {
        doc_ = std::move(other.doc_);
        options_ = other.options_;
        isValid_ = other.isValid_;
        other.isValid_ = false;
    }
    return *this;
}

TXResult<void> TXXmlReader::readFromZip(TXZipArchiveReader& zipReader, const std::string& xmlPath) {
    auto xmlData = zipReader.read(xmlPath);
    if (xmlData.isError()) {
        return Err<void>(TXErrorCode::XmlParseError, "Failed to read XML from ZIP: " + xmlData.error().getMessage());
    }
    
    const std::vector<uint8_t>& fileBytes = xmlData.value();
    std::string xmlContent(fileBytes.begin(), fileBytes.end());
    return parseFromString(xmlContent);
}

TXResult<void> TXXmlReader::parseFromString(const std::string& xmlContent) {
    auto parseResult = doc_->load_string(xmlContent.c_str(), getParseFlags());
    
    if (!parseResult) {
        std::ostringstream oss;
        oss << "XML parse error: " << parseResult.description() 
            << " at offset " << parseResult.offset;
        isValid_ = false;
        return Err<void>(TXErrorCode::XmlParseError, oss.str());
    }

    isValid_ = true;
    return Ok();
}

TXResult<std::vector<XmlNodeInfo>> TXXmlReader::findNodes(const std::string& xpath) const {
    if (!isValid_) {
        return Err<std::vector<XmlNodeInfo>>(TXErrorCode::XmlInvalidState, "XML document is not valid");
    }

    try {
        std::vector<XmlNodeInfo> results;
        auto nodeSet = doc_->select_nodes(xpath.c_str());
        for (const auto& nodeWrapper : nodeSet) {
            results.push_back(convertNode(nodeWrapper.node()));
        }
        return Ok(std::move(results));
    } catch (const pugi::xpath_exception& e) {
        return Err<std::vector<XmlNodeInfo>>(TXErrorCode::XmlXpathError, "XPath error: " + std::string(e.what()));
    }
}

TXResult<XmlNodeInfo> TXXmlReader::getRootNode() const {
    if (!isValid_) {
        return Err<XmlNodeInfo>(TXErrorCode::XmlInvalidState, "XML document is not valid");
    }

    auto root = doc_->document_element();
    if (!root) {
        return Err<XmlNodeInfo>(TXErrorCode::XmlNoRoot, "No root element found");
    }
    
    return Ok(convertNode(root));
}

TXResult<std::string> TXXmlReader::getNodeText(const std::string& xpath) const {
    if (!isValid_) {
        return Err<std::string>(TXErrorCode::XmlInvalidState, "XML document is not valid");
    }

    try {
        auto node = doc_->select_node(xpath.c_str()).node();
        if (!node) {
            return Err<std::string>(TXErrorCode::XmlNodeNotFound, "Node not found: " + xpath);
        }
        return Ok(std::string(node.text().as_string()));
    } catch (const pugi::xpath_exception& e) {
        return Err<std::string>(TXErrorCode::XmlXpathError, "XPath error: " + std::string(e.what()));
    }
}

TXResult<std::string> TXXmlReader::getNodeAttribute(const std::string& xpath, const std::string& attributeName) const {
    if (!isValid_) {
        return Err<std::string>(TXErrorCode::XmlInvalidState, "XML document is not valid");
    }

    try {
        auto node = doc_->select_node(xpath.c_str()).node();
        if (!node) {
            return Err<std::string>(TXErrorCode::XmlNodeNotFound, "Node not found: " + xpath);
        }
        
        auto attr = node.attribute(attributeName.c_str());
        if (!attr) {
            return Err<std::string>(TXErrorCode::XmlAttributeNotFound, "Attribute not found: " + attributeName);
        }
        
        return Ok(std::string(attr.as_string()));
    } catch (const pugi::xpath_exception& e) {
        return Err<std::string>(TXErrorCode::XmlXpathError, "XPath error: " + std::string(e.what()));
    }
}

TXResult<std::vector<std::string>> TXXmlReader::getAllNodeTexts(const std::string& xpath) const {
    if (!isValid_) {
        return Err<std::vector<std::string>>(TXErrorCode::XmlInvalidState, "XML document is not valid");
    }

    try {
        std::vector<std::string> results;
        auto nodeSet = doc_->select_nodes(xpath.c_str());
        for (const auto& nodeWrapper : nodeSet) {
            results.push_back(nodeWrapper.node().text().as_string());
        }
        return Ok(std::move(results));
    } catch (const pugi::xpath_exception& e) {
        return Err<std::vector<std::string>>(TXErrorCode::XmlXpathError, "XPath error: " + std::string(e.what()));
    }
}

bool TXXmlReader::isValid() const {
    return isValid_;
}

void TXXmlReader::reset() {
    doc_->reset();
    isValid_ = false;
}

XmlNodeInfo TXXmlReader::convertNode(const pugi::xml_node& node) const {
    XmlNodeInfo info;
    info.name = node.name();
    info.value = node.text().as_string();

    // 转换属性
    for (const auto& attr : node.attributes()) {
        info.attributes[attr.name()] = attr.value();
    }

    // 转换子节点
    for (const auto& child : node.children()) {
        if (child.type() == pugi::node_element) {
            info.children.push_back(convertNode(child));
        }
    }

    return info;
}

unsigned int TXXmlReader::getParseFlags() const {
    unsigned int flags = pugi::parse_default;
    
    if (!options_.preserve_whitespace) {
        flags |= pugi::parse_trim_pcdata;
    }
    if (!options_.merge_pcdata) {
        flags &= ~pugi::parse_merge_pcdata;
    }
    if (options_.trim_pcdata) {
        flags |= pugi::parse_trim_pcdata;
    }
    
    return flags;
}

} // namespace TinaXlsx
