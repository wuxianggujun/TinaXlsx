/*
 * TinaXlsx - Modern C++ Excel File Processing Library
 * 
 * Copyright (c) 2025 wuxianggujun
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

#pragma once

#include "TXXmlHandler.hpp"
#include "TXPivotTable.hpp"
#include "TXXmlWriter.hpp"

namespace TinaXlsx {

    /**
     * @brief 透视表XML处理器
     * 
     * 负责生成Excel透视表相关的XML文件，包括：
     * - pivotTable.xml - 透视表定义
     * - pivotCacheDefinition.xml - 数据缓存定义
     * - pivotCacheRecords.xml - 数据缓存记录
     */
    class TXPivotTableXmlHandler : public TXXmlHandler {
    public:
        /**
         * @brief 构造函数
         * @param pivotTable 透视表对象
         * @param pivotTableId 透视表ID
         */
        TXPivotTableXmlHandler(const TXPivotTable* pivotTable, int pivotTableId);

        /**
         * @brief 析构函数
         */
        ~TXPivotTableXmlHandler() override = default;

        /**
         * @brief 加载透视表XML文件
         * @param zipReader ZIP读取器
         * @param context 工作簿上下文
         * @return 操作结果
         */
        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override;

        /**
         * @brief 保存透视表XML文件
         * @param zipWriter ZIP写入器
         * @param context 工作簿上下文
         * @return 操作结果
         */
        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override;

        /**
         * @brief 获取文件路径
         * @return 透视表文件路径
         */
        std::string partName() const override;

    private:
        const TXPivotTable* m_pivotTable;  ///< 透视表对象
        int m_pivotTableId;                ///< 透视表ID

        /**
         * @brief 生成透视表定义XML
         * @return XML节点构建器
         */
        XmlNodeBuilder generatePivotTableDefinitionXml() const;

        /**
         * @brief 生成透视表缓存定义XML
         * @return XML节点构建器
         */
        XmlNodeBuilder generatePivotCacheDefinitionXml() const;

        /**
         * @brief 生成透视表缓存记录XML
         * @return XML节点构建器
         */
        XmlNodeBuilder generatePivotCacheRecordsXml() const;

        /**
         * @brief 生成透视表字段XML
         * @return XML节点构建器
         */
        XmlNodeBuilder generatePivotFieldsXml() const;

        /**
         * @brief 生成行字段XML
         * @return XML节点构建器
         */
        XmlNodeBuilder generateRowFieldsXml() const;

        /**
         * @brief 生成列字段XML
         * @return XML节点构建器
         */
        XmlNodeBuilder generateColFieldsXml() const;

        /**
         * @brief 生成数据字段XML
         * @return XML节点构建器
         */
        XmlNodeBuilder generateDataFieldsXml() const;

        /**
         * @brief 获取聚合函数字符串
         * @param func 聚合函数枚举
         * @return 聚合函数字符串
         */
        std::string getAggregateString(PivotAggregateFunction func) const;
    };

    /**
     * @brief 透视表缓存XML处理器
     */
    class TXPivotCacheXmlHandler : public TXXmlHandler {
    public:
        /**
         * @brief 构造函数
         * @param pivotTable 透视表对象
         * @param cacheId 缓存ID
         */
        TXPivotCacheXmlHandler(const TXPivotTable* pivotTable, int cacheId);

        /**
         * @brief 析构函数
         */
        ~TXPivotCacheXmlHandler() override = default;

        /**
         * @brief 加载缓存XML文件
         * @param zipReader ZIP读取器
         * @param context 工作簿上下文
         * @return 操作结果
         */
        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override;

        /**
         * @brief 保存缓存XML文件
         * @param zipWriter ZIP写入器
         * @param context 工作簿上下文
         * @return 操作结果
         */
        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override;

        /**
         * @brief 获取文件路径
         * @return 缓存文件路径
         */
        std::string partName() const override;

    private:
        const TXPivotTable* m_pivotTable;  ///< 透视表对象
        int m_cacheId;                     ///< 缓存ID

        /**
         * @brief 计算记录数量
         * @return 记录数量字符串
         */
        std::string calculateRecordCount() const;
    };

} // namespace TinaXlsx
