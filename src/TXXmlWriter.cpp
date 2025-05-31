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

    // TXXmlWriter::Impl 实现
    class TXXmlWriter::Impl
    {
    public:
        pugi::xml_document doc_;
        XmlWriteOptions options_;
        std::string lastError_;
        bool isValid_ = false;

        // 递归构建 XML 节点
        void buildNode(pugi::xml_node& parent, const XmlNodeBuilder& builder)
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

        // 生成格式化的 XML 字符串
        std::string generateString() const
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
                    doc_.save(oss, options_.indent.c_str(), pugi::format_default | pugi::format_no_declaration,
                              pugi::encoding_utf8);
                }
                else
                {
                    doc_.save(oss, "", pugi::format_raw | pugi::format_no_declaration, pugi::encoding_utf8);
                }
            }
            else
            {
                // 不包含声明
                if (options_.format_output)
                {
                    doc_.save(oss, options_.indent.c_str(), pugi::format_default | pugi::format_no_declaration,
                              pugi::encoding_utf8);
                }
                else
                {
                    doc_.save(oss, "", pugi::format_raw | pugi::format_no_declaration, pugi::encoding_utf8);
                }
            }

            return oss.str();
        }

        // 统计文档信息
        DocumentStats calculateStats(const pugi::xml_node& node) const
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
    };

    TXXmlWriter::TXXmlWriter() : pImpl_(std::make_unique<Impl>())
    {
    }

    TXXmlWriter::TXXmlWriter(const XmlWriteOptions& options)
        : pImpl_(std::make_unique<Impl>())
    {
        pImpl_->options_ = options;
    }

    TXXmlWriter::~TXXmlWriter() = default;

    TXXmlWriter::TXXmlWriter(TXXmlWriter&& other) noexcept
        : pImpl_(std::move(other.pImpl_))
    {
    }

    TXXmlWriter& TXXmlWriter::operator=(TXXmlWriter&& other) noexcept
    {
        if (this != &other)
        {
            pImpl_ = std::move(other.pImpl_);
        }
        return *this;
    }

    void TXXmlWriter::setRootNode(const XmlNodeBuilder& rootNode)
    {
        pImpl_->doc_.reset();
        pImpl_->buildNode(pImpl_->doc_, rootNode);
        pImpl_->isValid_ = true;
        pImpl_->lastError_.clear();
    }

    void TXXmlWriter::createDocument(const std::string& rootNodeName)
    {
        pImpl_->doc_.reset();
        auto root = pImpl_->doc_.append_child(rootNodeName.c_str());
        pImpl_->isValid_ = true;
        pImpl_->lastError_.clear();
    }

    void TXXmlWriter::addRootChild(const XmlNodeBuilder& node)
    {
        if (!pImpl_->isValid_)
        {
            pImpl_->lastError_ = "Document not initialized";
            return;
        }

        auto root = pImpl_->doc_.document_element();
        if (!root)
        {
            pImpl_->lastError_ = "No root element found";
            return;
        }

        pImpl_->buildNode(root, node);
    }

    std::string TXXmlWriter::generateXmlString() const
    {
        return pImpl_->generateString();
    }

    bool TXXmlWriter::writeToZip(TXZipArchiveWriter& zipWriter, const std::string& xmlPath) const
    {
        if (!pImpl_->isValid_)
        {
            pImpl_->lastError_ = "Document not valid";
            return false;
        }

        std::string xmlContent = generateXmlString();
        if (xmlContent.empty())
        {
            pImpl_->lastError_ = "Failed to generate XML content";
            return false;
        }

        std::vector<uint8_t> xmlData(xmlContent.begin(), xmlContent.end());
        auto success = zipWriter.write(xmlPath, xmlData);

        if (success.isError())
        {
            pImpl_->lastError_ = "Failed to write to ZIP: " + success.error().getMessage();
            return false;
        }

        return true;
    }

    bool TXXmlWriter::writeStringToZip(TXZipArchiveWriter& zipWriter,
                                       const std::string& xmlPath,
                                       const std::string& xmlContent)
    {
        std::vector<uint8_t> xmlData(xmlContent.begin(), xmlContent.end());
        auto success = zipWriter.write(xmlPath, xmlData);

        if (success.isError())
        {
            return false;
        }
        return true;
    }

    bool TXXmlWriter::isValid() const
    {
        return pImpl_->isValid_;
    }

    const std::string& TXXmlWriter::getLastError() const
    {
        return pImpl_->lastError_;
    }

    void TXXmlWriter::reset()
    {
        pImpl_->doc_.reset();
        pImpl_->isValid_ = false;
        pImpl_->lastError_.clear();
    }

    TXXmlWriter::DocumentStats TXXmlWriter::getStats() const
    {
        if (!pImpl_->isValid_)
        {
            return {};
        }

        return pImpl_->calculateStats(pImpl_->doc_.document_element());
    }
} // namespace TinaXlsx
