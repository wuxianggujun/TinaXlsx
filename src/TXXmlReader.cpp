#include "TinaXlsx/TXXmlReader.hpp"
#include <pugixml.hpp>
#include <algorithm>

namespace TinaXlsx {

// TXXmlNode 方法实现
std::string TXXmlNode::getAttribute(const std::string& attrName, const std::string& defaultValue) const {
    auto it = std::find_if(attributes.begin(), attributes.end(),
        [&attrName](const std::pair<std::string, std::string>& attr) {
            return attr.first == attrName;
        });
    
    return (it != attributes.end()) ? it->second : defaultValue;
}

bool TXXmlNode::hasAttribute(const std::string& attrName) const {
    return std::any_of(attributes.begin(), attributes.end(),
        [&attrName](const std::pair<std::string, std::string>& attr) {
            return attr.first == attrName;
        });
}

const TXXmlNode* TXXmlNode::findChild(const std::string& childName) const {
    auto it = std::find_if(children.begin(), children.end(),
        [&childName](const TXXmlNode& child) {
            return child.name == childName;
        });
    
    return (it != children.end()) ? &(*it) : nullptr;
}

std::vector<const TXXmlNode*> TXXmlNode::findChildren(const std::string& childName) const {
    std::vector<const TXXmlNode*> result;
    
    for (const auto& child : children) {
        if (child.name == childName) {
            result.push_back(&child);
        }
    }
    
    return result;
}

class TXXmlReader::Impl {
public:
    Impl() = default;
    
    bool loadFromString(const std::string& xmlString) {
        reset();
        
        pugi::xml_parse_result result = doc_.load_string(xmlString.c_str());
        if (!result) {
            lastError_ = "XML parse error: " + std::string(result.description()) +
                        " at offset " + std::to_string(result.offset);
            return false;
        }
        
        // 构建内部节点结构
        buildNodeTree();
        
        lastError_.clear();
        return true;
    }
    
    bool loadFromFile(const std::string& filename) {
        reset();
        
        pugi::xml_parse_result result = doc_.load_file(filename.c_str());
        if (!result) {
            lastError_ = "XML parse error: " + std::string(result.description()) +
                        " in file " + filename;
            return false;
        }
        
        // 构建内部节点结构
        buildNodeTree();
        
        lastError_.clear();
        return true;
    }
    
    const TXXmlNode* getRootNode() const {
        return rootNode_.name.empty() ? nullptr : &rootNode_;
    }
    
    std::vector<const TXXmlNode*> findNodesByXPath(const std::string& xpath) const {
        std::vector<const TXXmlNode*> result;
        
        try {
            pugi::xpath_node_set nodes = doc_.select_nodes(xpath.c_str());
            
            for (const auto& node : nodes) {
                if (node.node()) {
                    // 在我们的节点树中找到对应的节点
                    const TXXmlNode* found = findNodeInTree(&rootNode_, node.node());
                    if (found) {
                        result.push_back(found);
                    }
                }
            }
        } catch (const pugi::xpath_exception& e) {
            lastError_ = "XPath error: " + std::string(e.what());
        }
        
        return result;
    }
    
    const TXXmlNode* findFirstNodeByXPath(const std::string& xpath) const {
        auto nodes = findNodesByXPath(xpath);
        return nodes.empty() ? nullptr : nodes[0];
    }
    
    void traverse(const std::function<void(const TXXmlNode*)>& visitor) const {
        if (!rootNode_.name.empty()) {
            traverseNode(&rootNode_, visitor);
        }
    }
    
    const std::string& getLastError() const {
        return lastError_;
    }
    
    void reset() {
        doc_.reset();
        rootNode_ = TXXmlNode{};
        lastError_.clear();
    }

private:
    void buildNodeTree() {
        auto root = doc_.document_element();
        if (root) {
            rootNode_ = convertNode(root);
        }
    }
    
    TXXmlNode convertNode(const pugi::xml_node& pugiNode) {
        TXXmlNode node;
        node.name = pugiNode.name();
        node.text = pugiNode.text().get();
        
        // 转换属性
        for (const auto& attr : pugiNode.attributes()) {
            node.attributes.emplace_back(attr.name(), attr.value());
        }
        
        // 递归转换子节点
        for (const auto& child : pugiNode.children()) {
            if (child.type() == pugi::node_element) {
                node.children.push_back(convertNode(child));
            }
        }
        
        return node;
    }
    
    const TXXmlNode* findNodeInTree(const TXXmlNode* node, const pugi::xml_node& pugiNode) const {
        if (!node) return nullptr;
        
        // 简单的名称匹配（这里可以改进为更精确的匹配）
        if (node->name == pugiNode.name()) {
            return node;
        }
        
        // 递归查找子节点
        for (const auto& child : node->children) {
            const TXXmlNode* found = findNodeInTree(&child, pugiNode);
            if (found) {
                return found;
            }
        }
        
        return nullptr;
    }
    
    void traverseNode(const TXXmlNode* node, const std::function<void(const TXXmlNode*)>& visitor) const {
        if (!node) return;
        
        visitor(node);
        
        for (const auto& child : node->children) {
            traverseNode(&child, visitor);
        }
    }

private:
    pugi::xml_document doc_;
    TXXmlNode rootNode_;
    mutable std::string lastError_;
};

// TXXmlReader 实现
TXXmlReader::TXXmlReader() : pImpl(std::make_unique<Impl>()) {}

TXXmlReader::~TXXmlReader() = default;

bool TXXmlReader::loadFromString(const std::string& xmlString) {
    return pImpl->loadFromString(xmlString);
}

bool TXXmlReader::loadFromFile(const std::string& filename) {
    return pImpl->loadFromFile(filename);
}

const TXXmlNode* TXXmlReader::getRootNode() const {
    return pImpl->getRootNode();
}

std::vector<const TXXmlNode*> TXXmlReader::findNodesByXPath(const std::string& xpath) const {
    return pImpl->findNodesByXPath(xpath);
}

const TXXmlNode* TXXmlReader::findFirstNodeByXPath(const std::string& xpath) const {
    return pImpl->findFirstNodeByXPath(xpath);
}

void TXXmlReader::traverse(const std::function<void(const TXXmlNode*)>& visitor) const {
    pImpl->traverse(visitor);
}

const std::string& TXXmlReader::getLastError() const {
    return pImpl->getLastError();
}

void TXXmlReader::reset() {
    pImpl->reset();
}

} // namespace TinaXlsx 