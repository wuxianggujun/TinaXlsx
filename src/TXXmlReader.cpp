//
// @file TXXmlReader.cpp
// @brief XML 读取器实现
//

#include "TinaXlsx/TXXmlReader.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include <pugixml.hpp>
#include <sstream>

namespace TinaXlsx {

// TXXmlReader::Impl 实现
class TXXmlReader::Impl {
public:
    pugi::xml_document doc_;
    XmlParseOptions options_;
    std::string lastError_;
    bool isValid_ = false;

    // 将 pugi::xml_node 转换为 XmlNodeInfo
    XmlNodeInfo convertNode(const pugi::xml_node& node) const {
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

    // 设置解析选项
    unsigned int getParseFlags() const {
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
};

TXXmlReader::TXXmlReader() : pImpl_(std::make_unique<Impl>()) {}

TXXmlReader::TXXmlReader(const XmlParseOptions& options) 
    : pImpl_(std::make_unique<Impl>()) {
    pImpl_->options_ = options;
}

TXXmlReader::~TXXmlReader() = default;

TXXmlReader::TXXmlReader(TXXmlReader&& other) noexcept 
    : pImpl_(std::move(other.pImpl_)) {}

TXXmlReader& TXXmlReader::operator=(TXXmlReader&& other) noexcept {
    if (this != &other) {
        pImpl_ = std::move(other.pImpl_);
    }
    return *this;
}

bool TXXmlReader::readFromZip(TXZipArchiveReader& zipReader, const std::string& xmlPath) {
    auto xmlData = zipReader.read(xmlPath);
    if (xmlData.empty()) {
        pImpl_->lastError_ = "Failed to read XML from ZIP: " + zipReader.lastError();
        pImpl_->isValid_ = false;
        return false;
    }

    std::string xmlContent(xmlData.begin(), xmlData.end());
    return parseFromString(xmlContent);
}

bool TXXmlReader::parseFromString(const std::string& xmlContent) {
    auto parseResult = pImpl_->doc_.load_string(xmlContent.c_str(), pImpl_->getParseFlags());
    
    if (!parseResult) {
        std::ostringstream oss;
        oss << "XML parse error: " << parseResult.description() 
            << " at offset " << parseResult.offset;
        pImpl_->lastError_ = oss.str();
        pImpl_->isValid_ = false;
        return false;
    }

    pImpl_->isValid_ = true;
    pImpl_->lastError_.clear();
    return true;
}

std::vector<XmlNodeInfo> TXXmlReader::findNodes(const std::string& xpath) const {
    std::vector<XmlNodeInfo> results;
    
    if (!pImpl_->isValid_) {
        return results;
    }

    try {
        auto nodeSet = pImpl_->doc_.select_nodes(xpath.c_str());
        for (const auto& nodeWrapper : nodeSet) {
            results.push_back(pImpl_->convertNode(nodeWrapper.node()));
        }
    } catch (const pugi::xpath_exception& e) {
        pImpl_->lastError_ = "XPath error: " + std::string(e.what());
    }

    return results;
}

XmlNodeInfo TXXmlReader::getRootNode() const {
    if (!pImpl_->isValid_) {
        return {};
    }

    auto root = pImpl_->doc_.document_element();
    return pImpl_->convertNode(root);
}

std::string TXXmlReader::getNodeText(const std::string& xpath) const {
    if (!pImpl_->isValid_) {
        return "";
    }

    try {
        auto node = pImpl_->doc_.select_node(xpath.c_str()).node();
        return node.text().as_string();
    } catch (const pugi::xpath_exception&) {
        return "";
    }
}

std::string TXXmlReader::getNodeAttribute(const std::string& xpath, const std::string& attributeName) const {
    if (!pImpl_->isValid_) {
        return "";
    }

    try {
        auto node = pImpl_->doc_.select_node(xpath.c_str()).node();
        return node.attribute(attributeName.c_str()).as_string();
    } catch (const pugi::xpath_exception&) {
        return "";
    }
}

std::vector<std::string> TXXmlReader::getAllNodeTexts(const std::string& xpath) const {
    std::vector<std::string> results;
    
    if (!pImpl_->isValid_) {
        return results;
    }

    try {
        auto nodeSet = pImpl_->doc_.select_nodes(xpath.c_str());
        for (const auto& nodeWrapper : nodeSet) {
            results.push_back(nodeWrapper.node().text().as_string());
        }
    } catch (const pugi::xpath_exception&) {
        // Ignore XPath errors for this method
    }

    return results;
}

bool TXXmlReader::isValid() const {
    return pImpl_->isValid_;
}

const std::string& TXXmlReader::getLastError() const {
    return pImpl_->lastError_;
}

void TXXmlReader::reset() {
    pImpl_->doc_.reset();
    pImpl_->isValid_ = false;
    pImpl_->lastError_.clear();
}

} // namespace TinaXlsx