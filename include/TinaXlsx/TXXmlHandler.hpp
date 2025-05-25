#pragma once

#include <string>
#include <vector>
#include <memory>
#include <unordered_map>
#include <functional>

// Forward declaration for pugixml
namespace pugi {
    class xml_document;
    class xml_node;
    class xml_attribute;
}

namespace TinaXlsx {

/**
 * @brief 高性能XML处理器
 * 
 * 基于pugixml库实现的XML文档解析和生成功能，
 * 专门优化用于处理XLSX文件中的XML组件
 */
class TXXmlHandler {
public:
    /**
     * @brief XML节点信息
     */
    struct XmlNodeInfo {
        std::string name;           ///< 节点名称
        std::string value;          ///< 节点值
        std::unordered_map<std::string, std::string> attributes;  ///< 属性映射
        std::vector<XmlNodeInfo> children;  ///< 子节点
    };

    /**
     * @brief XML解析选项
     */
    struct ParseOptions {
        bool preserve_whitespace = false;    ///< 是否保留空白字符
        bool merge_pcdata = true;           ///< 是否合并PCDATA节点
        bool validate_encoding = true;      ///< 是否验证编码
        bool trim_pcdata = true;           ///< 是否去除文本首尾空白
    };

public:
    TXXmlHandler();
    explicit TXXmlHandler(const ParseOptions& options);
    ~TXXmlHandler();

    // 禁用拷贝构造和赋值
    TXXmlHandler(const TXXmlHandler&) = delete;
    TXXmlHandler& operator=(const TXXmlHandler&) = delete;

    // 支持移动构造和赋值
    TXXmlHandler(TXXmlHandler&& other) noexcept;
    TXXmlHandler& operator=(TXXmlHandler&& other) noexcept;

    /**
     * @brief 从字符串解析XML文档
     * @param xml_content XML内容字符串
     * @return 成功返回true，失败返回false
     */
    bool parseFromString(const std::string& xml_content);

    /**
     * @brief 从文件解析XML文档
     * @param filename XML文件路径
     * @return 成功返回true，失败返回false
     */
    bool parseFromFile(const std::string& filename);

    /**
     * @brief 保存XML文档到字符串
     * @param formatted 是否格式化输出
     * @return XML字符串，失败返回空字符串
     */
    std::string saveToString(bool formatted = true) const;

    /**
     * @brief 保存XML文档到文件
     * @param filename 输出文件路径
     * @param formatted 是否格式化输出
     * @return 成功返回true，失败返回false
     */
    bool saveToFile(const std::string& filename, bool formatted = true) const;

    /**
     * @brief 检查XML文档是否有效
     * @return 有效返回true，否则返回false
     */
    bool isValid() const;

    /**
     * @brief 获取根节点名称
     * @return 根节点名称，如果无效返回空字符串
     */
    std::string getRootName() const;

    /**
     * @brief 查找节点（使用XPath）
     * @param xpath XPath表达式
     * @return 找到的节点信息列表
     */
    std::vector<XmlNodeInfo> findNodes(const std::string& xpath) const;

    /**
     * @brief 查找单个节点（使用XPath）
     * @param xpath XPath表达式
     * @return 找到的节点信息，如果未找到返回空节点
     */
    XmlNodeInfo findNode(const std::string& xpath) const;

    /**
     * @brief 获取节点的文本内容
     * @param xpath XPath表达式
     * @return 节点文本内容，如果未找到返回空字符串
     */
    std::string getNodeText(const std::string& xpath) const;

    /**
     * @brief 获取节点的属性值
     * @param xpath XPath表达式
     * @param attr_name 属性名称
     * @return 属性值，如果未找到返回空字符串
     */
    std::string getNodeAttribute(const std::string& xpath, const std::string& attr_name) const;

    /**
     * @brief 设置节点的文本内容
     * @param xpath XPath表达式
     * @param text 要设置的文本
     * @return 成功返回true，失败返回false
     */
    bool setNodeText(const std::string& xpath, const std::string& text);

    /**
     * @brief 设置节点的属性值
     * @param xpath XPath表达式
     * @param attr_name 属性名称
     * @param attr_value 属性值
     * @return 成功返回true，失败返回false
     */
    bool setNodeAttribute(const std::string& xpath, const std::string& attr_name, const std::string& attr_value);

    /**
     * @brief 添加子节点
     * @param parent_xpath 父节点XPath
     * @param node_name 新节点名称
     * @param node_text 节点文本（可选）
     * @return 成功返回true，失败返回false
     */
    bool addChildNode(const std::string& parent_xpath, const std::string& node_name, const std::string& node_text = "");

    /**
     * @brief 删除节点
     * @param xpath XPath表达式
     * @return 成功删除的节点数量
     */
    std::size_t removeNodes(const std::string& xpath);

    /**
     * @brief 批量查询节点（高性能版本）
     * @param xpaths XPath表达式列表
     * @return XPath到节点信息的映射
     */
    std::unordered_map<std::string, std::vector<XmlNodeInfo>> batchFindNodes(const std::vector<std::string>& xpaths) const;

    /**
     * @brief 批量设置节点文本（高性能版本）
     * @param xpath_to_text XPath到文本的映射
     * @return 成功设置的节点数量
     */
    std::size_t batchSetNodeText(const std::unordered_map<std::string, std::string>& xpath_to_text);

    /**
     * @brief 获取最后的错误信息
     * @return 错误信息字符串
     */
    const std::string& getLastError() const;

    /**
     * @brief 重置文档
     */
    void reset();

    /**
     * @brief 创建新的XML文档
     * @param root_name 根节点名称
     * @param encoding 文档编码（默认UTF-8）
     * @return 成功返回true，失败返回false
     */
    bool createDocument(const std::string& root_name, const std::string& encoding = "UTF-8");

    /**
     * @brief 获取文档统计信息
     */
    struct DocumentStats {
        std::size_t total_nodes = 0;
        std::size_t total_attributes = 0;
        std::size_t max_depth = 0;
        std::size_t document_size = 0;  // 字符数
    };

    /**
     * @brief 获取文档统计信息
     * @return 文档统计信息
     */
    DocumentStats getDocumentStats() const;

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

} // namespace TinaXlsx 