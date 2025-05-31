//
// @file TXXmlWriter.cpp
// @brief XML 写入器实现
//

#include "TinaXlsx/TXXmlWriter.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include <pugixml.hpp>
#include <sstream>

namespace TinaXlsx
{
    // XmlNodeBuilder 实现
    XmlNodeBuilder::XmlNodeBuilder(const std::string& name) : name_(name)
    {
    }

    XmlNodeBuilder& XmlNodeBuilder::setText(const std::string& text)
    {
        text_ = text;
        return *this;
    }

    XmlNodeBuilder& XmlNodeBuilder::addAttribute(const std::string& name, const std::string& value)
    {
        attributes_[name] = value;
        return *this;
    }

    XmlNodeBuilder& XmlNodeBuilder::addChild(const XmlNodeBuilder& child)
    {
        children_.push_back(child);
        return *this;
    }

    const std::string& XmlNodeBuilder::getName() const { return name_; }
    const std::string& XmlNodeBuilder::getText() const { return text_; }
    const std::unordered_map<std::string, std::string>& XmlNodeBuilder::getAttributes() const { return attributes_; }
    const std::vector<XmlNodeBuilder>& XmlNodeBuilder::getChildren() const { return children_; }

    TXXmlWriter::TXXmlWriter() : doc_(std::make_unique<pugi::xml_document>())
    {
    }

    TXXmlWriter::TXXmlWriter(const XmlWriteOptions& options)
        : doc_(std::make_unique<pugi::xml_document>()), options_(options)
    {
    }

    TXXmlWriter::~TXXmlWriter() = default;

    TXXmlWriter::TXXmlWriter(TXXmlWriter&& other) noexcept
        : doc_(std::move(other.doc_)), options_(other.options_), isValid_(other.isValid_)
    {
        other.isValid_ = false;
    }

    TXXmlWriter& TXXmlWriter::operator=(TXXmlWriter&& other) noexcept
    {
        if (this != &other)
        {
            doc_ = std::move(other.doc_);
            options_ = other.options_;
            isValid_ = other.isValid_;
            other.isValid_ = false;
        }
        return *this;
    }

    TXResult<void> TXXmlWriter::setRootNode(const XmlNodeBuilder& rootNode)
    {
        doc_->reset();
        buildNode(*doc_, rootNode);
        isValid_ = true;
        return Ok();
    }

    TXResult<void> TXXmlWriter::createDocument(const std::string& rootNodeName)
    {
        doc_->reset();
        auto root = doc_->append_child(rootNodeName.c_str());
        if (!root) {
            return Err<void>(TXErrorCode::XML_CREATE_ERROR, "Failed to create root node: " + rootNodeName);
        }
        isValid_ = true;
        return Ok();
    }

    TXResult<void> TXXmlWriter::addRootChild(const XmlNodeBuilder& node)
    {
        if (!isValid_)
        {
            return Err<void>(TXErrorCode::XML_INVALID_STATE, "Document not initialized");
        }

        auto root = doc_->document_element();
        if (!root)
        {
            return Err<void>(TXErrorCode::XML_NO_ROOT, "No root element found");
        }

        buildNode(root, node);
        return Ok();
    }

    TXResult<std::string> TXXmlWriter::generateXmlString() const
    {
        if (!isValid_) {
            return Err<std::string>(TXErrorCode::XML_INVALID_STATE, "Document is not valid");
        }
        
        std::string result = generateString();
        if (result.empty()) {
            return Err<std::string>(TXErrorCode::XML_GENERATE_ERROR, "Failed to generate XML content");
        }
        
        return Ok(std::move(result));
    }

    TXResult<void> TXXmlWriter::writeToZip(TXZipArchiveWriter& zipWriter, const std::string& xmlPath) const
    {
        auto xmlContentResult = generateXmlString();
        if (xmlContentResult.isError()) {
            return Err<void>(xmlContentResult.error().getCode(), "Failed to generate XML: " + xmlContentResult.error().getMessage());
        }

        std::vector<uint8_t> xmlData(xmlContentResult.value().begin(), xmlContentResult.value().end());
        auto writeResult = zipWriter.write(xmlPath, xmlData);

        if (writeResult.isError())
        {
            return Err<void>(TXErrorCode::ZIP_WRITE_ERROR, "Failed to write to ZIP: " + writeResult.error().getMessage());
        }

        return Ok();
    }

    TXResult<void> TXXmlWriter::writeStringToZip(TXZipArchiveWriter& zipWriter,
                                       const std::string& xmlPath,
                                       const std::string& xmlContent)
    {
        std::vector<uint8_t> xmlData(xmlContent.begin(), xmlContent.end());
        auto writeResult = zipWriter.write(xmlPath, xmlData);

        if (writeResult.isError())
        {
            return Err<void>(TXErrorCode::ZIP_WRITE_ERROR, "Failed to write to ZIP: " + writeResult.error().getMessage());
        }
        
        return Ok();
    }

    bool TXXmlWriter::isValid() const
    {
        return isValid_;
    }

    void TXXmlWriter::reset()
    {
        doc_->reset();
        isValid_ = false;
    }

    TXResult<TXXmlWriter::DocumentStats> TXXmlWriter::getStats() const
    {
        if (!isValid_)
        {
            return Err<DocumentStats>(TXErrorCode::XML_INVALID_STATE, "Document is not valid");
        }

        return Ok(calculateStats(doc_->document_element()));
    }

    void TXXmlWriter::buildNode(pugi::xml_node& parent, const XmlNodeBuilder& builder)
    {
        pugi::xml_node node = parent.append_child(builder.getName().c_str());

        // 设置属性
        for (const auto& [name, value] : builder.getAttributes())
        {
            node.append_attribute(name.c_str()).set_value(value.c_str());
        }

        // 设置文本内容
        if (!builder.getText().empty())
        {
            node.text().set(builder.getText().c_str());
        }

        // 递归添加子节点
        for (const auto& child : builder.getChildren())
        {
            buildNode(node, child);
        }
    }

    std::string TXXmlWriter::generateString() const
    {
        if (!isValid_)
        {
            return "";
        }

        std::ostringstream oss;

        // 根据选项决定如何生成XML
        if (options_.include_declaration)
        {
            // 手动添加XML声明以确保包含自定义编码
            oss << "<?xml version=\"1.0\" encoding=\"" << options_.encoding << "\"?>";
            if (options_.format_output)
            {
                oss << "\n";
            }

            // 保存文档内容，但不包含声明（因为我们已经手动添加了）
            if (options_.format_output)
            {
                doc_->save(oss, options_.indent.c_str(), pugi::format_default | pugi::format_no_declaration,
                          pugi::encoding_utf8);
            }
            else
            {
                doc_->save(oss, "", pugi::format_raw | pugi::format_no_declaration, pugi::encoding_utf8);
            }
        }
        else
        {
            // 不包含声明
            if (options_.format_output)
            {
                doc_->save(oss, options_.indent.c_str(), pugi::format_default | pugi::format_no_declaration,
                          pugi::encoding_utf8);
            }
            else
            {
                doc_->save(oss, "", pugi::format_raw | pugi::format_no_declaration, pugi::encoding_utf8);
            }
        }

        return oss.str();
    }

    TXXmlWriter::DocumentStats TXXmlWriter::calculateStats(const pugi::xml_node& node) const
    {
        DocumentStats stats;

        // 统计当前节点
        if (node.type() == pugi::node_element)
        {
            stats.nodeCount++;
            stats.attributeCount += std::distance(node.attributes_begin(), node.attributes_end());
            stats.textLength += std::strlen(node.text().as_string());
        }

        // 递归统计子节点
        for (const auto& child : node.children())
        {
            auto childStats = calculateStats(child);
            stats.nodeCount += childStats.nodeCount;
            stats.attributeCount += childStats.attributeCount;
            stats.textLength += childStats.textLength;
        }

        return stats;
    }
} // namespace TinaXlsx
