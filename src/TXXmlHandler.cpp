#include "TinaXlsx/TXXmlHandler.hpp"
#include <pugixml.hpp>
#include <algorithm>
#include <sstream>

namespace TinaXlsx {

class TXXmlHandler::Impl {
public:
    Impl() : document_(), options_(), last_error_("") {}
    
    explicit Impl(const TXXmlHandler::ParseOptions& options) 
        : document_(), options_(options), last_error_("") {}

    ~Impl() = default;

    bool parseFromString(const std::string& xml_content) {
        document_.reset();
        
        unsigned int parse_flags = pugi::parse_default;
        if (options_.preserve_whitespace) {
            parse_flags |= pugi::parse_ws_pcdata;
        }
        if (options_.merge_pcdata) {
            parse_flags |= pugi::parse_merge_pcdata;
        }
        if (options_.trim_pcdata) {
            parse_flags |= pugi::parse_trim_pcdata;
        }

        pugi::xml_parse_result result = document_.load_string(xml_content.c_str(), parse_flags);
        if (!result) {
            last_error_ = "XML parse error: " + std::string(result.description()) + 
                         " at offset " + std::to_string(result.offset);
            return false;
        }

        last_error_.clear();
        return true;
    }

    bool parseFromFile(const std::string& filename) {
        document_.reset();
        
        unsigned int parse_flags = pugi::parse_default;
        if (options_.preserve_whitespace) {
            parse_flags |= pugi::parse_ws_pcdata;
        }
        if (options_.merge_pcdata) {
            parse_flags |= pugi::parse_merge_pcdata;
        }
        if (options_.trim_pcdata) {
            parse_flags |= pugi::parse_trim_pcdata;
        }

        pugi::xml_parse_result result = document_.load_file(filename.c_str(), parse_flags);
        if (!result) {
            last_error_ = "XML parse error: " + std::string(result.description()) + 
                         " at offset " + std::to_string(result.offset);
            return false;
        }

        last_error_.clear();
        return true;
    }

    std::string saveToString(bool formatted) const {
        if (!isValid()) {
            return "";
        }

        std::ostringstream oss;
        unsigned int save_flags = formatted ? pugi::format_indent : pugi::format_raw;
        document_.save(oss, "  ", save_flags, pugi::encoding_utf8);
        return oss.str();
    }

    bool saveToFile(const std::string& filename, bool formatted) const {
        if (!isValid()) {
            const_cast<Impl*>(this)->last_error_ = "Invalid XML document";
            return false;
        }

        unsigned int save_flags = formatted ? pugi::format_indent : pugi::format_raw;
        bool result = document_.save_file(filename.c_str(), "  ", save_flags, pugi::encoding_utf8);
        if (!result) {
            const_cast<Impl*>(this)->last_error_ = "Failed to save XML file: " + filename;
            return false;
        }

        const_cast<Impl*>(this)->last_error_.clear();
        return true;
    }

    bool isValid() const {
        return !document_.empty();
    }

    std::string getRootName() const {
        if (!isValid()) {
            return "";
        }
        pugi::xml_node root = document_.document_element();
        return root ? root.name() : "";
    }

    TXXmlHandler::XmlNodeInfo convertNodeToInfo(const pugi::xml_node& node) const {
        TXXmlHandler::XmlNodeInfo info;
        info.name = node.name();
        info.value = node.child_value();

        // 收集属性
        for (const auto& attr : node.attributes()) {
            info.attributes[attr.name()] = attr.value();
        }

        // 收集子节点
        for (const auto& child : node.children()) {
            if (child.type() == pugi::node_element) {
                info.children.push_back(convertNodeToInfo(child));
            }
        }

        return info;
    }

    std::vector<TXXmlHandler::XmlNodeInfo> findNodes(const std::string& xpath) const {
        std::vector<TXXmlHandler::XmlNodeInfo> results;
        if (!isValid()) {
            return results;
        }

        try {
            pugi::xpath_query query(xpath.c_str());
            pugi::xpath_node_set nodes = query.evaluate_node_set(document_);
            
            for (const auto& node : nodes) {
                if (node.node()) {
                    results.push_back(convertNodeToInfo(node.node()));
                }
            }
        } catch (const pugi::xpath_exception& e) {
            const_cast<Impl*>(this)->last_error_ = "XPath error: " + std::string(e.what());
        }

        return results;
    }

    TXXmlHandler::XmlNodeInfo findNode(const std::string& xpath) const {
        auto nodes = findNodes(xpath);
        return nodes.empty() ? TXXmlHandler::XmlNodeInfo{} : nodes[0];
    }

    std::string getNodeText(const std::string& xpath) const {
        if (!isValid()) {
            return "";
        }

        try {
            pugi::xpath_query query(xpath.c_str());
            pugi::xpath_node_set nodes = query.evaluate_node_set(document_);
            
            if (!nodes.empty() && nodes[0].node()) {
                return nodes[0].node().child_value();
            }
        } catch (const pugi::xpath_exception& e) {
            const_cast<Impl*>(this)->last_error_ = "XPath error: " + std::string(e.what());
        }

        return "";
    }

    std::string getNodeAttribute(const std::string& xpath, const std::string& attr_name) const {
        if (!isValid()) {
            return "";
        }

        try {
            pugi::xpath_query query(xpath.c_str());
            pugi::xpath_node_set nodes = query.evaluate_node_set(document_);
            
            if (!nodes.empty() && nodes[0].node()) {
                return nodes[0].node().attribute(attr_name.c_str()).as_string();
            }
        } catch (const pugi::xpath_exception& e) {
            const_cast<Impl*>(this)->last_error_ = "XPath error: " + std::string(e.what());
        }

        return "";
    }

    bool setNodeText(const std::string& xpath, const std::string& text) {
        if (!isValid()) {
            last_error_ = "Invalid XML document";
            return false;
        }

        try {
            pugi::xpath_query query(xpath.c_str());
            pugi::xpath_node_set nodes = query.evaluate_node_set(document_);
            
            if (!nodes.empty() && nodes[0].node()) {
                nodes[0].node().text().set(text.c_str());
                last_error_.clear();
                return true;
            }
        } catch (const pugi::xpath_exception& e) {
            last_error_ = "XPath error: " + std::string(e.what());
        }

        return false;
    }

    bool setNodeAttribute(const std::string& xpath, const std::string& attr_name, const std::string& attr_value) {
        if (!isValid()) {
            last_error_ = "Invalid XML document";
            return false;
        }

        try {
            pugi::xpath_query query(xpath.c_str());
            pugi::xpath_node_set nodes = query.evaluate_node_set(document_);
            
            if (!nodes.empty() && nodes[0].node()) {
                pugi::xml_attribute attr = nodes[0].node().attribute(attr_name.c_str());
                if (!attr) {
                    attr = nodes[0].node().append_attribute(attr_name.c_str());
                }
                attr.set_value(attr_value.c_str());
                last_error_.clear();
                return true;
            }
        } catch (const pugi::xpath_exception& e) {
            last_error_ = "XPath error: " + std::string(e.what());
        }

        return false;
    }

    bool addChildNode(const std::string& parent_xpath, const std::string& node_name, const std::string& node_text) {
        if (!isValid()) {
            last_error_ = "Invalid XML document";
            return false;
        }

        try {
            pugi::xpath_query query(parent_xpath.c_str());
            pugi::xpath_node_set nodes = query.evaluate_node_set(document_);
            
            if (!nodes.empty() && nodes[0].node()) {
                pugi::xml_node new_node = nodes[0].node().append_child(node_name.c_str());
                if (!node_text.empty()) {
                    new_node.text().set(node_text.c_str());
                }
                last_error_.clear();
                return true;
            }
        } catch (const pugi::xpath_exception& e) {
            last_error_ = "XPath error: " + std::string(e.what());
        }

        return false;
    }

    std::size_t removeNodes(const std::string& xpath) {
        if (!isValid()) {
            last_error_ = "Invalid XML document";
            return 0;
        }

        std::size_t removed_count = 0;
        try {
            pugi::xpath_query query(xpath.c_str());
            pugi::xpath_node_set nodes = query.evaluate_node_set(document_);
            
            for (const auto& node : nodes) {
                if (node.node() && node.parent()) {
                    if (node.parent().remove_child(node.node())) {
                        ++removed_count;
                    }
                }
            }
            last_error_.clear();
        } catch (const pugi::xpath_exception& e) {
            last_error_ = "XPath error: " + std::string(e.what());
        }

        return removed_count;
    }

    std::unordered_map<std::string, std::vector<TXXmlHandler::XmlNodeInfo>> 
    batchFindNodes(const std::vector<std::string>& xpaths) const {
        std::unordered_map<std::string, std::vector<TXXmlHandler::XmlNodeInfo>> results;
        
        for (const auto& xpath : xpaths) {
            results[xpath] = findNodes(xpath);
        }
        
        return results;
    }

    std::size_t batchSetNodeText(const std::unordered_map<std::string, std::string>& xpath_to_text) {
        std::size_t success_count = 0;
        
        for (const auto& pair : xpath_to_text) {
            if (setNodeText(pair.first, pair.second)) {
                ++success_count;
            }
        }
        
        return success_count;
    }

    void reset() {
        document_.reset();
        last_error_.clear();
    }

    bool createDocument(const std::string& root_name, const std::string& encoding) {
        document_.reset();
        
        pugi::xml_node decl = document_.prepend_child(pugi::node_declaration);
        decl.append_attribute("version") = "1.0";
        decl.append_attribute("encoding") = encoding.c_str();
        
        pugi::xml_node root = document_.append_child(root_name.c_str());
        if (!root) {
            last_error_ = "Failed to create root node: " + root_name;
            return false;
        }
        
        last_error_.clear();
        return true;
    }

    TXXmlHandler::DocumentStats getDocumentStats() const {
        TXXmlHandler::DocumentStats stats;
        if (!isValid()) {
            return stats;
        }

        // 计算统计信息
        std::function<void(const pugi::xml_node&, std::size_t)> countNodes = 
            [&](const pugi::xml_node& node, std::size_t depth) {
                if (node.type() == pugi::node_element) {
                    ++stats.total_nodes;
                    stats.max_depth = std::max(stats.max_depth, depth);
                    
                    // 计算属性数量
                    for (const auto& attr : node.attributes()) {
                        ++stats.total_attributes;
                        (void)attr; // 避免未使用变量警告
                    }
                    
                    // 递归处理子节点
                    for (const auto& child : node.children()) {
                        countNodes(child, depth + 1);
                    }
                }
            };

        countNodes(document_.document_element(), 1);
        
        // 计算文档大小
        std::string doc_str = saveToString(false);
        stats.document_size = doc_str.length();
        
        return stats;
    }

    const std::string& getLastError() const {
        return last_error_;
    }

private:
    pugi::xml_document document_;
    TXXmlHandler::ParseOptions options_;
    std::string last_error_;
};

// TXXmlHandler 实现
TXXmlHandler::TXXmlHandler() : pImpl(std::make_unique<Impl>()) {}

TXXmlHandler::TXXmlHandler(const ParseOptions& options) 
    : pImpl(std::make_unique<Impl>(options)) {}

TXXmlHandler::~TXXmlHandler() = default;

TXXmlHandler::TXXmlHandler(TXXmlHandler&& other) noexcept : pImpl(std::move(other.pImpl)) {}

TXXmlHandler& TXXmlHandler::operator=(TXXmlHandler&& other) noexcept {
    if (this != &other) {
        pImpl = std::move(other.pImpl);
    }
    return *this;
}

bool TXXmlHandler::parseFromString(const std::string& xml_content) {
    return pImpl->parseFromString(xml_content);
}

bool TXXmlHandler::parseFromFile(const std::string& filename) {
    return pImpl->parseFromFile(filename);
}

std::string TXXmlHandler::saveToString(bool formatted) const {
    return pImpl->saveToString(formatted);
}

bool TXXmlHandler::saveToFile(const std::string& filename, bool formatted) const {
    return pImpl->saveToFile(filename, formatted);
}

bool TXXmlHandler::isValid() const {
    return pImpl->isValid();
}

std::string TXXmlHandler::getRootName() const {
    return pImpl->getRootName();
}

std::vector<TXXmlHandler::XmlNodeInfo> TXXmlHandler::findNodes(const std::string& xpath) const {
    return pImpl->findNodes(xpath);
}

TXXmlHandler::XmlNodeInfo TXXmlHandler::findNode(const std::string& xpath) const {
    return pImpl->findNode(xpath);
}

std::string TXXmlHandler::getNodeText(const std::string& xpath) const {
    return pImpl->getNodeText(xpath);
}

std::string TXXmlHandler::getNodeAttribute(const std::string& xpath, const std::string& attr_name) const {
    return pImpl->getNodeAttribute(xpath, attr_name);
}

bool TXXmlHandler::setNodeText(const std::string& xpath, const std::string& text) {
    return pImpl->setNodeText(xpath, text);
}

bool TXXmlHandler::setNodeAttribute(const std::string& xpath, const std::string& attr_name, const std::string& attr_value) {
    return pImpl->setNodeAttribute(xpath, attr_name, attr_value);
}

bool TXXmlHandler::addChildNode(const std::string& parent_xpath, const std::string& node_name, const std::string& node_text) {
    return pImpl->addChildNode(parent_xpath, node_name, node_text);
}

std::size_t TXXmlHandler::removeNodes(const std::string& xpath) {
    return pImpl->removeNodes(xpath);
}

std::unordered_map<std::string, std::vector<TXXmlHandler::XmlNodeInfo>> 
TXXmlHandler::batchFindNodes(const std::vector<std::string>& xpaths) const {
    return pImpl->batchFindNodes(xpaths);
}

std::size_t TXXmlHandler::batchSetNodeText(const std::unordered_map<std::string, std::string>& xpath_to_text) {
    return pImpl->batchSetNodeText(xpath_to_text);
}

const std::string& TXXmlHandler::getLastError() const {
    return pImpl->getLastError();
}

void TXXmlHandler::reset() {
    pImpl->reset();
}

bool TXXmlHandler::createDocument(const std::string& root_name, const std::string& encoding) {
    return pImpl->createDocument(root_name, encoding);
}

TXXmlHandler::DocumentStats TXXmlHandler::getDocumentStats() const {
    return pImpl->getDocumentStats();
}

} // namespace TinaXlsx 