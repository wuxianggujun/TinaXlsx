//
// @file TXXmlWriter.hpp
// @brief XML 写入器 - 专门处理 XLSX 文件的 XML 写入操作
//

#pragma once

#include <string>
#include <vector>
#include <memory>
#include <unordered_map>

namespace TinaXlsx {

// 前向声明
class TXZipArchiveWriter;

/**
 * @brief XML 写入选项
 */
struct XmlWriteOptions {
    bool format_output = true;          ///< 是否格式化输出
    std::string indent = "  ";          ///< 缩进字符串
    bool include_declaration = true;    ///< 是否包含 XML 声明
    std::string encoding = "UTF-8";     ///< 编码格式
};

/**
 * @brief XML 节点构建器
 */
class XmlNodeBuilder {
public:
    XmlNodeBuilder(const std::string& name);
    
    /**
     * @brief 设置节点文本内容
     */
    XmlNodeBuilder& setText(const std::string& text);
    
    /**
     * @brief 添加属性
     */
    XmlNodeBuilder& addAttribute(const std::string& name, const std::string& value);
    
    /**
     * @brief 添加子节点
     */
    XmlNodeBuilder& addChild(const XmlNodeBuilder& child);
    
    /**
     * @brief 获取节点名称
     */
    const std::string& getName() const;
    
    /**
     * @brief 获取节点文本
     */
    const std::string& getText() const;
    
    /**
     * @brief 获取属性映射
     */
    const std::unordered_map<std::string, std::string>& getAttributes() const;
    
    /**
     * @brief 获取子节点列表
     */
    const std::vector<XmlNodeBuilder>& getChildren() const;

private:
    std::string name_;
    std::string text_;
    std::unordered_map<std::string, std::string> attributes_;
    std::vector<XmlNodeBuilder> children_;
};

/**
 * @brief XML 写入器
 * 
 * 专门用于生成 XLSX 文件中的 XML 内容，封装了 ZIP 写入操作
 */
class TXXmlWriter {
public:
    TXXmlWriter();
    explicit TXXmlWriter(const XmlWriteOptions& options);
    ~TXXmlWriter();

    // 禁用拷贝，支持移动
    TXXmlWriter(const TXXmlWriter&) = delete;
    TXXmlWriter& operator=(const TXXmlWriter&) = delete;
    TXXmlWriter(TXXmlWriter&& other) noexcept;
    TXXmlWriter& operator=(TXXmlWriter&& other) noexcept;

    /**
     * @brief 设置根节点
     * @param rootNode 根节点构建器
     */
    void setRootNode(const XmlNodeBuilder& rootNode);

    /**
     * @brief 创建空白文档
     * @param rootNodeName 根节点名称
     */
    void createDocument(const std::string& rootNodeName);

    /**
     * @brief 添加根级节点
     * @param node 要添加的节点
     */
    void addRootChild(const XmlNodeBuilder& node);

    /**
     * @brief 生成 XML 字符串
     * @return 格式化的 XML 字符串
     */
    std::string generateXmlString() const;

    /**
     * @brief 写入到 ZIP 文件
     * @param zipWriter ZIP 写入器引用
     * @param xmlPath XML 文件在 ZIP 中的路径
     * @return 是否成功写入
     */
    bool writeToZip(TXZipArchiveWriter& zipWriter, const std::string& xmlPath) const;

    /**
     * @brief 直接从字符串写入到 ZIP
     * @param zipWriter ZIP 写入器引用
     * @param xmlPath XML 文件在 ZIP 中的路径
     * @param xmlContent XML 内容字符串
     * @return 是否成功写入
     */
    static bool writeStringToZip(TXZipArchiveWriter& zipWriter, 
                                 const std::string& xmlPath, 
                                 const std::string& xmlContent);

    /**
     * @brief 检查写入器是否有效
     * @return true 如果写入器状态正常
     */
    bool isValid() const;

    /**
     * @brief 获取错误信息
     * @return 最后一次操作的错误信息
     */
    const std::string& getLastError() const;

    /**
     * @brief 重置写入器状态
     */
    void reset();

    /**
     * @brief 获取当前文档的统计信息
     */
    struct DocumentStats {
        size_t nodeCount = 0;
        size_t attributeCount = 0;
        size_t textLength = 0;
    };
    
    DocumentStats getStats() const;

private:
    class Impl;
    std::unique_ptr<Impl> pImpl_;
};

} // namespace TinaXlsx