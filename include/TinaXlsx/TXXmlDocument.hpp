#pragma once

#include <string>
#include <vector>
#include <memory>
#include <unordered_map>
#include <variant>

namespace TinaXlsx {

/**
 * @brief XML属性值类型
 */
using XmlAttributeValue = std::variant<std::string, int, double, bool>;

/**
 * @brief XML节点类 - 轻量级XML节点表示
 */
class TXXmlNode {
public:
    TXXmlNode();
    explicit TXXmlNode(const std::string& name);
    TXXmlNode(const std::string& name, const std::string& text);
    ~TXXmlNode();

    // 支持拷贝和移动
    TXXmlNode(const TXXmlNode& other);
    TXXmlNode& operator=(const TXXmlNode& other);
    TXXmlNode(TXXmlNode&& other) noexcept;
    TXXmlNode& operator=(TXXmlNode&& other) noexcept;

    /**
     * @brief 获取节点名称
     */
    const std::string& getName() const;

    /**
     * @brief 设置节点名称
     */
    void setName(const std::string& name);

    /**
     * @brief 获取节点文本内容
     */
    const std::string& getText() const;

    /**
     * @brief 设置节点文本内容
     */
    void setText(const std::string& text);

    /**
     * @brief 添加属性
     */
    void setAttribute(const std::string& name, const XmlAttributeValue& value);

    /**
     * @brief 获取属性值
     */
    std::string getAttribute(const std::string& name, const std::string& defaultValue = "") const;

    /**
     * @brief 检查是否有属性
     */
    bool hasAttribute(const std::string& name) const;

    /**
     * @brief 移除属性
     */
    void removeAttribute(const std::string& name);

    /**
     * @brief 获取所有属性
     */
    const std::unordered_map<std::string, std::string>& getAttributes() const;

    /**
     * @brief 添加子节点
     */
    TXXmlNode* addChild(const std::string& name);
    TXXmlNode* addChild(const std::string& name, const std::string& text);
    TXXmlNode* addChild(const TXXmlNode& child);

    /**
     * @brief 获取子节点
     */
    TXXmlNode* getChild(const std::string& name);
    const TXXmlNode* getChild(const std::string& name) const;

    /**
     * @brief 获取所有子节点
     */
    std::vector<TXXmlNode*> getChildren();
    std::vector<const TXXmlNode*> getChildren() const;

    /**
     * @brief 获取指定名称的所有子节点
     */
    std::vector<TXXmlNode*> getChildren(const std::string& name);
    std::vector<const TXXmlNode*> getChildren(const std::string& name) const;

    /**
     * @brief 移除子节点
     */
    bool removeChild(const std::string& name);
    bool removeChild(const TXXmlNode* child);

    /**
     * @brief 清空所有子节点
     */
    void clearChildren();

    /**
     * @brief 检查是否有子节点
     */
    bool hasChildren() const;

    /**
     * @brief 获取子节点数量
     */
    std::size_t getChildCount() const;

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

/**
 * @brief XML文档读取器 - 基于Desktop Commander MCP服务
 */
class TXXmlDocumentReader {
public:
    TXXmlDocumentReader();
    ~TXXmlDocumentReader();

    // 禁用拷贝，支持移动
    TXXmlDocumentReader(const TXXmlDocumentReader&) = delete;
    TXXmlDocumentReader& operator=(const TXXmlDocumentReader&) = delete;
    TXXmlDocumentReader(TXXmlDocumentReader&& other) noexcept;
    TXXmlDocumentReader& operator=(TXXmlDocumentReader&& other) noexcept;

    /**
     * @brief 从字符串解析XML
     * @param xmlContent XML内容
     * @return 成功返回true
     */
    bool parseFromString(const std::string& xmlContent);

    /**
     * @brief 从文件解析XML
     * @param filePath XML文件路径
     * @return 成功返回true
     */
    bool parseFromFile(const std::string& filePath);

    /**
     * @brief 获取根节点
     */
    TXXmlNode* getRootNode();
    const TXXmlNode* getRootNode() const;

    /**
     * @brief 查找节点（使用简单路径，如 "root/child/subchild"）
     * @param path 节点路径
     * @return 找到的节点，nullptr表示未找到
     */
    TXXmlNode* findNode(const std::string& path);
    const TXXmlNode* findNode(const std::string& path) const;

    /**
     * @brief 查找所有匹配的节点
     * @param path 节点路径
     * @return 找到的节点列表
     */
    std::vector<TXXmlNode*> findNodes(const std::string& path);
    std::vector<const TXXmlNode*> findNodes(const std::string& path) const;

    /**
     * @brief 获取最后的错误信息
     */
    const std::string& getLastError() const;

    /**
     * @brief 检查文档是否有效
     */
    bool isValid() const;

    /**
     * @brief 重置文档
     */
    void reset();

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

/**
 * @brief XML文档写入器 - 基于Desktop Commander MCP服务
 */
class TXXmlDocumentWriter {
public:
    TXXmlDocumentWriter();
    ~TXXmlDocumentWriter();

    // 禁用拷贝，支持移动
    TXXmlDocumentWriter(const TXXmlDocumentWriter&) = delete;
    TXXmlDocumentWriter& operator=(const TXXmlDocumentWriter&) = delete;
    TXXmlDocumentWriter(TXXmlDocumentWriter&& other) noexcept;
    TXXmlDocumentWriter& operator=(TXXmlDocumentWriter&& other) noexcept;

    /**
     * @brief 创建新文档
     * @param rootName 根节点名称
     * @param encoding 文档编码
     * @return 成功返回true
     */
    bool createDocument(const std::string& rootName, const std::string& encoding = "UTF-8");

    /**
     * @brief 设置根节点
     */
    void setRootNode(const TXXmlNode& rootNode);

    /**
     * @brief 获取根节点
     */
    TXXmlNode* getRootNode();
    const TXXmlNode* getRootNode() const;

    /**
     * @brief 生成XML字符串
     * @param formatted 是否格式化输出
     * @return XML字符串
     */
    std::string generateXml(bool formatted = true) const;

    /**
     * @brief 保存到文件
     * @param filePath 文件路径
     * @param formatted 是否格式化输出
     * @return 成功返回true
     */
    bool saveToFile(const std::string& filePath, bool formatted = true) const;

    /**
     * @brief 获取最后的错误信息
     */
    const std::string& getLastError() const;

    /**
     * @brief 重置文档
     */
    void reset();

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

/**
 * @brief XML文档工厂类
 */
class TXXmlDocumentFactory {
public:
    /**
     * @brief 创建XML读取器
     */
    static std::unique_ptr<TXXmlDocumentReader> createReader();

    /**
     * @brief 创建XML写入器
     */
    static std::unique_ptr<TXXmlDocumentWriter> createWriter();

    /**
     * @brief 创建XML节点
     */
    static std::unique_ptr<TXXmlNode> createNode(const std::string& name);
    static std::unique_ptr<TXXmlNode> createNode(const std::string& name, const std::string& text);
};

} // namespace TinaXlsx 