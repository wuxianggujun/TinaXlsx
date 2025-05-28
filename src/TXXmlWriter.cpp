#include "TinaXlsx/TXXmlWriter.hpp"
#include <pugixml.hpp>
#include <sstream>
#include <stack>
#include <regex>

namespace TinaXlsx {

// XML实体编码函数
std::string encodeXmlEntities(const std::string& input) {
    std::string result = input;
    result = std::regex_replace(result, std::regex("&"), "&amp;");
    result = std::regex_replace(result, std::regex("<"), "&lt;");
    result = std::regex_replace(result, std::regex(">"), "&gt;");
    result = std::regex_replace(result, std::regex("\""), "&quot;");
    result = std::regex_replace(result, std::regex("'"), "&apos;");
    return result;
}

class TXXmlWriter::Impl {
public:
    Impl() {
        reset();
    }
    
    void startDocument(const std::string& encoding, bool standalone) {
        reset();
        
        // 创建XML声明
        auto declaration = doc_.prepend_child(pugi::node_declaration);
        declaration.append_attribute("version") = "1.0";
        declaration.append_attribute("encoding") = encoding.c_str();
        if (standalone) {
            declaration.append_attribute("standalone") = "yes";
        }
        
        currentNode_ = &doc_;
    }
    
    void startElement(const std::string& name) {
        if (!currentNode_) {
            // 如果没有调用startDocument，创建一个默认的根节点
            currentNode_ = &doc_;
        }
        
        pugi::xml_node newNode = currentNode_->append_child(name.c_str());
        nodeStack_.push(currentNode_);
        currentNode_ = &newNode;
    }
    
    void startElement(const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes) {
        startElement(name);
        
        for (const auto& attr : attributes) {
            addAttribute(attr.first, attr.second);
        }
    }
    
    void endElement() {
        if (!nodeStack_.empty()) {
            currentNode_ = nodeStack_.top();
            nodeStack_.pop();
        }
    }
    
    void addAttribute(const std::string& name, const std::string& value) {
        if (currentNode_ && currentNode_->type() == pugi::node_element) {
            currentNode_->append_attribute(name.c_str()) = value.c_str();
        }
    }
    
    void addText(const std::string& text, bool encode) {
        if (currentNode_ && currentNode_->type() == pugi::node_element) {
            std::string textToAdd = encode ? encodeXmlEntities(text) : text;
            currentNode_->text().set(textToAdd.c_str());
        }
    }
    
    void addCData(const std::string& data) {
        if (currentNode_ && currentNode_->type() == pugi::node_element) {
            auto cdataNode = currentNode_->append_child(pugi::node_cdata);
            cdataNode.set_value(data.c_str());
        }
    }
    
    std::string toString() const {
        std::ostringstream oss;
        doc_.save(oss, "  ", pugi::format_default, pugi::encoding_utf8);
        return oss.str();
    }
    
    void reset() {
        doc_.reset();
        while (!nodeStack_.empty()) {
            nodeStack_.pop();
        }
        currentNode_ = nullptr;
    }

private:
    pugi::xml_document doc_;
    pugi::xml_node* currentNode_;
    std::stack<pugi::xml_node*> nodeStack_;
};

// TXXmlWriter 实现
TXXmlWriter::TXXmlWriter() : pImpl(std::make_unique<Impl>()) {}

TXXmlWriter::~TXXmlWriter() = default;

void TXXmlWriter::startDocument(const std::string& encoding, bool standalone) {
    pImpl->startDocument(encoding, standalone);
}

void TXXmlWriter::startElement(const std::string& name) {
    pImpl->startElement(name);
}

void TXXmlWriter::startElement(const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes) {
    pImpl->startElement(name, attributes);
}

void TXXmlWriter::endElement() {
    pImpl->endElement();
}

void TXXmlWriter::addAttribute(const std::string& name, const std::string& value) {
    pImpl->addAttribute(name, value);
}

void TXXmlWriter::addText(const std::string& text, bool encode) {
    pImpl->addText(text, encode);
}

void TXXmlWriter::addCData(const std::string& data) {
    pImpl->addCData(data);
}

std::string TXXmlWriter::toString() const {
    return pImpl->toString();
}

void TXXmlWriter::reset() {
    pImpl->reset();
}

} // namespace TinaXlsx 