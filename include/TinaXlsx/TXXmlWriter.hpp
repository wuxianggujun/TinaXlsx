#pragma once

#include <string>
#include <vector>
#include <memory>

namespace TinaXlsx {

/**
 * @brief XML写入器 - 专门的SAX/DOM写入包装（基于pugixml）
 * 
 * 提供结构化的XML写入功能，避免字符串拼接错误
 */
class TXXmlWriter {
public:
    TXXmlWriter();
    ~TXXmlWriter();

    /**
     * @brief 开始创建XML文档
     * @param encoding 编码格式，默认UTF-8
     * @param standalone 是否独立文档
     */
    void startDocument(const std::string& encoding = "UTF-8", bool standalone = true);

    /**
     * @brief 开始元素
     * @param name 元素名称
     */
    void startElement(const std::string& name);

    /**
     * @brief 开始元素并添加属性
     * @param name 元素名称
     * @param attributes 属性映射
     */
    void startElement(const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes);

    /**
     * @brief 结束元素
     */
    void endElement();

    /**
     * @brief 添加属性到当前元素
     * @param name 属性名
     * @param value 属性值
     */
    void addAttribute(const std::string& name, const std::string& value);

    /**
     * @brief 添加文本内容
     * @param text 文本内容
     * @param encode 是否编码XML实体
     */
    void addText(const std::string& text, bool encode = true);

    /**
     * @brief 添加CDATA内容
     * @param data CDATA内容
     */
    void addCData(const std::string& data);

    /**
     * @brief 获取生成的XML字符串
     * @return XML字符串
     */
    std::string toString() const;

    /**
     * @brief 重置写入器
     */
    void reset();

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

} // namespace TinaXlsx 