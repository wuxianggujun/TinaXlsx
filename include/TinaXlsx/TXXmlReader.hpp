//
// @file TXXmlReader.hpp
// @brief XML 读取器 - 专门处理 XLSX 文件的 XML 读取操作
//

#pragma once

#include <string>
#include <vector>
#include <memory>
#include <unordered_map>

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
     * @return 是否成功读取
     */
    bool readFromZip(TXZipArchiveReader& zipReader, const std::string& xmlPath);

    /**
     * @brief 从字符串解析 XML
     * @param xmlContent XML 内容字符串
     * @return 是否解析成功
     */
    bool parseFromString(const std::string& xmlContent);

    /**
     * @brief 查找节点
     * @param xpath XPath 表达式
     * @return 匹配的节点列表
     */
    std::vector<XmlNodeInfo> findNodes(const std::string& xpath) const;

    /**
     * @brief 获取根节点
     * @return 根节点信息
     */
    XmlNodeInfo getRootNode() const;

    /**
     * @brief 获取节点文本内容
     * @param xpath XPath 表达式
     * @return 节点文本，如果不存在返回空字符串
     */
    std::string getNodeText(const std::string& xpath) const;

    /**
     * @brief 获取节点属性值
     * @param xpath XPath 表达式
     * @param attributeName 属性名
     * @return 属性值，如果不存在返回空字符串
     */
    std::string getNodeAttribute(const std::string& xpath, const std::string& attributeName) const;

    /**
     * @brief 获取所有匹配节点的文本内容
     * @param xpath XPath 表达式
     * @return 所有匹配节点的文本列表
     */
    std::vector<std::string> getAllNodeTexts(const std::string& xpath) const;

    /**
     * @brief 检查 XML 是否有效
     * @return true 如果 XML 已加载且有效
     */
    bool isValid() const;

    /**
     * @brief 获取错误信息
     * @return 最后一次操作的错误信息
     */
    const std::string& getLastError() const;

    /**
     * @brief 重置读取器状态
     */
    void reset();

private:
    class Impl;
    std::unique_ptr<Impl> pImpl_;
};

} // namespace TinaXlsx