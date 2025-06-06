#pragma once

#include <string>
#include <functional>
#include <vector>
#include <pugixml.hpp>
#include "TXTypes.hpp"
#include "TXResult.hpp"

namespace TinaXlsx {

    /**
     * @brief 标准XML解析器回调接口
     */
    class IStandardXmlCallback {
    public:
        virtual ~IStandardXmlCallback() = default;
        
        virtual void onStartElement(const std::string& name, const std::vector<std::pair<std::string, std::string>>& attributes) = 0;
        virtual void onEndElement(const std::string& name) = 0;
        virtual void onText(const std::string& text) = 0;
        virtual void onError(const std::string& error) = 0;
    };

    /**
     * @brief 解析选项
     */
    struct ParseOptions {
        bool preserveWhitespace = false;
        bool validateInput = true;
        size_t maxDepth = 1000;
    };

    /**
     * @brief 标准XML解析器 (从TXSIMDXmlParser简化而来)
     */
    class TXStandardXmlParser {
    public:
        TXStandardXmlParser() = default;
        ~TXStandardXmlParser() = default;

        // 删除复制和移动语义
        TXStandardXmlParser(const TXStandardXmlParser&) = delete;
        TXStandardXmlParser& operator=(const TXStandardXmlParser&) = delete;
        TXStandardXmlParser(TXStandardXmlParser&&) = delete;
        TXStandardXmlParser& operator=(TXStandardXmlParser&&) = delete;

        /**
         * @brief 解析XML内容
         * @param xmlContent XML内容字符串
         * @param callback 解析回调
         * @param options 解析选项
         * @return 解析结果
         */
        TXResult<void> parse(const std::string& xmlContent, 
                            IStandardXmlCallback& callback, 
                            const ParseOptions& options = {});

        /**
         * @brief 解析XML文件
         * @param filePath 文件路径
         * @param callback 解析回调
         * @param options 解析选项
         * @return 解析结果
         */
        TXResult<void> parseFile(const std::string& filePath, 
                                IStandardXmlCallback& callback, 
                                const ParseOptions& options = {});

        /**
         * @brief 获取解析统计信息
         */
        struct ParseStats {
            size_t elementsProcessed = 0;
            size_t attributesProcessed = 0;
            size_t textNodesProcessed = 0;
            double parseTimeMs = 0.0;
        };

        [[nodiscard]] const ParseStats& getStats() const { return stats_; }

    private:
        void processNode(pugi::xml_node node, IStandardXmlCallback& callback, size_t depth, const ParseOptions& options);
        std::vector<std::pair<std::string, std::string>> extractAttributes(pugi::xml_node node);

        ParseStats stats_;
    };

} // namespace TinaXlsx 