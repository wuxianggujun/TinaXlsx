//
// @file TXXmlReader.hpp
// @brief XML 读取器 - 专门处理 XLSX 文件的 XML 读取操作
//

#pragma once

#include <string>
#include <vector>
#include <memory>
#include <unordered_map>
#include "TXResult.hpp"

// 前向声明
namespace pugi {
class xml_document;
class xml_node;
}

namespace TinaXlsx {

// 前向声明
class TXZipArchiveReader;

/**
 * @brief XML 节点信息
 */
struct XmlNodeInfo {
    std::string name;           ///< 节点名称
    std::string value;          ///< 节点值
    std::unordered_map<std::string, std::string> attributes;  ///< 属性映射
    std::vector<XmlNodeInfo> children;  ///< 子节点
};

/**
 * @brief XML 解析选项
 */
struct XmlParseOptions {
    bool preserve_whitespace = false;    ///< 是否保留空白字符
    bool merge_pcdata = true;           ///< 是否合并PCDATA节点
    bool validate_encoding = true;      ///< 是否验证编码
    bool trim_pcdata = true;           ///< 是否去除文本首尾空白
};

/**
 * @brief XML 读取器
 * 
 * 专门用于读取 XLSX 文件中的 XML 内容，封装了 ZIP 读取操作
 */
class TXXmlReader {
public:
    TXXmlReader();
    explicit TXXmlReader(const XmlParseOptions& options);
    ~TXXmlReader();

    // 禁用拷贝，支持移动
    TXXmlReader(const TXXmlReader&) = delete;
    TXXmlReader& operator=(const TXXmlReader&) = delete;
    TXXmlReader(TXXmlReader&& other) noexcept;
    TXXmlReader& operator=(TXXmlReader&& other) noexcept;

    /**
     * @brief 从 ZIP 中读取 XML 文件
     * @param zipReader ZIP 读取器引用
     * @param xmlPath XML 文件在 ZIP 中的路径
     * @return TXResult<void> 操作结果
     */
    TXResult<void> readFromZip(TXZipArchiveReader& zipReader, const std::string& xmlPath);

    /**
     * @brief 从字符串解析 XML
     * @param xmlContent XML 内容字符串
     * @return TXResult<void> 解析结果
     */
    TXResult<void> parseFromString(const std::string& xmlContent);

    /**
     * @brief 查找节点
     * @param xpath XPath 表达式
     * @return TXResult<std::vector<XmlNodeInfo>> 匹配的节点列表
     */
    TXResult<std::vector<XmlNodeInfo>> findNodes(const std::string& xpath) const;

    /**
     * @brief 获取根节点
     * @return TXResult<XmlNodeInfo> 根节点信息
     */
    TXResult<XmlNodeInfo> getRootNode() const;

    /**
     * @brief 获取节点文本内容
     * @param xpath XPath 表达式
     * @return TXResult<std::string> 节点文本
     */
    TXResult<std::string> getNodeText(const std::string& xpath) const;

    /**
     * @brief 获取节点属性值
     * @param xpath XPath 表达式
     * @param attributeName 属性名
     * @return TXResult<std::string> 属性值
     */
    TXResult<std::string> getNodeAttribute(const std::string& xpath, const std::string& attributeName) const;

    /**
     * @brief 获取所有匹配节点的文本内容
     * @param xpath XPath 表达式
     * @return TXResult<std::vector<std::string>> 所有匹配节点的文本列表
     */
    TXResult<std::vector<std::string>> getAllNodeTexts(const std::string& xpath) const;

    /**
     * @brief 检查 XML 是否有效
     * @return true 如果 XML 已加载且有效
     */
    bool isValid() const;

    /**
     * @brief 重置读取器状态
     */
    void reset();

private:
    std::unique_ptr<pugi::xml_document> doc_;
    XmlParseOptions options_;
    bool isValid_ = false;

    // 将 pugi::xml_node 转换为 XmlNodeInfo
    XmlNodeInfo convertNode(const pugi::xml_node& node) const;
    
    // 获取解析标志
    unsigned int getParseFlags() const;
};

} // namespace TinaXlsx