#pragma once

#include <string>
#include <vector>
#include <memory>
#include <functional>

namespace TinaXlsx {

/**
 * @brief XML节点信息结构
 */
struct TXXmlNode {
    std::string name;
    std::string text;
    std::vector<std::pair<std::string, std::string>> attributes;
    std::vector<TXXmlNode> children;
    
    /**
     * @brief 获取属性值
     * @param attrName 属性名
     * @param defaultValue 默认值
     * @return 属性值
     */
    std::string getAttribute(const std::string& attrName, const std::string& defaultValue = "") const;
    
    /**
     * @brief 检查是否有指定属性
     * @param attrName 属性名
     * @return 有属性返回true
     */
    bool hasAttribute(const std::string& attrName) const;
    
    /**
     * @brief 查找第一个匹配名称的子节点
     * @param childName 子节点名称
     * @return 子节点指针，未找到返回nullptr
     */
    const TXXmlNode* findChild(const std::string& childName) const;
    
    /**
     * @brief 查找所有匹配名称的子节点
     * @param childName 子节点名称
     * @return 子节点列表
     */
    std::vector<const TXXmlNode*> findChildren(const std::string& childName) const;
};

/**
 * @brief XML读取器 - 专门的DOM读取包装（基于pugixml）
 * 
 * 提供结构化的XML读取功能，简化XML解析操作
 */
class TXXmlReader {
public:
    TXXmlReader();
    ~TXXmlReader();

    /**
     * @brief 从字符串加载XML
     * @param xmlString XML字符串
     * @return 成功返回true
     */
    bool loadFromString(const std::string& xmlString);

    /**
     * @brief 从文件加载XML
     * @param filename 文件路径
     * @return 成功返回true
     */
    bool loadFromFile(const std::string& filename);

    /**
     * @brief 获取根节点
     * @return 根节点指针，失败返回nullptr
     */
    const TXXmlNode* getRootNode() const;

    /**
     * @brief 使用XPath查找节点
     * @param xpath XPath表达式
     * @return 匹配的节点列表
     */
    std::vector<const TXXmlNode*> findNodesByXPath(const std::string& xpath) const;

    /**
     * @brief 查找第一个匹配XPath的节点
     * @param xpath XPath表达式
     * @return 节点指针，未找到返回nullptr
     */
    const TXXmlNode* findFirstNodeByXPath(const std::string& xpath) const;

    /**
     * @brief 遍历所有节点
     * @param visitor 访问函数，参数为节点指针
     */
    void traverse(const std::function<void(const TXXmlNode*)>& visitor) const;

    /**
     * @brief 获取最后的错误信息
     * @return 错误信息
     */
    const std::string& getLastError() const;

    /**
     * @brief 重置读取器
     */
    void reset();

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

} // namespace TinaXlsx 