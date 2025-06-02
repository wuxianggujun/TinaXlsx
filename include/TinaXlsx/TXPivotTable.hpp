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

#include "TXTypes.hpp"
#include "TXRange.hpp"
#include <string>
#include <vector>
#include <memory>
#include <unordered_map>

namespace TinaXlsx {

    // TODO: 透视表功能存在兼容性问题，需要进一步完善
    // 当前实现的XML结构虽然在技术上正确，但Excel/WPS仍无法正确识别
    // 可能的问题包括：
    // 1. 某些微妙的XML格式差异
    // 2. 缺少特定的Excel内部标识符或校验和
    // 3. 数据索引映射可能不完全准确
    // 4. 可能需要更精确的数据类型推断和元数据
    // 5. Excel可能对透视表XML有未公开的验证规则
    // 建议：暂时跳过透视表功能，优先实现其他更稳定的Excel功能
    // 如需继续开发，建议深入研究Excel的OLE复合文档格式或使用逆向工程方法

    // 前向声明
    class TXSheet;
    class TXPivotField;
    class TXPivotCache;

    /**
     * @brief 透视表聚合函数类型
     */
    enum class PivotAggregateFunction {
        Sum,        ///< 求和
        Count,      ///< 计数
        Average,    ///< 平均值
        Max,        ///< 最大值
        Min,        ///< 最小值
        Product,    ///< 乘积
        CountNums,  ///< 数值计数
        StdDev,     ///< 标准差
        StdDevP,    ///< 总体标准差
        Var,        ///< 方差
        VarP        ///< 总体方差
    };

    /**
     * @brief 透视表字段类型
     */
    enum class PivotFieldType {
        Row,        ///< 行字段
        Column,     ///< 列字段
        Data,       ///< 数据字段
        Filter      ///< 筛选字段
    };

    /**
     * @brief 透视表字段类
     * 
     * 表示透视表中的一个字段，包含字段名称、类型、聚合函数等信息
     */
    class TXPivotField {
    public:
        /**
         * @brief 构造函数
         * @param name 字段名称
         * @param type 字段类型
         */
        TXPivotField(const std::string& name, PivotFieldType type);

        /**
         * @brief 析构函数
         */
        ~TXPivotField() = default;

        // 基本属性
        const std::string& getName() const { return name_; }
        PivotFieldType getType() const { return type_; }
        
        // 聚合函数（仅对数据字段有效）
        void setAggregateFunction(PivotAggregateFunction func) { aggregateFunction_ = func; }
        PivotAggregateFunction getAggregateFunction() const { return aggregateFunction_; }
        
        // 字段显示名称
        void setDisplayName(const std::string& displayName) { displayName_ = displayName; }
        const std::string& getDisplayName() const { return displayName_.empty() ? name_ : displayName_; }
        
        // 排序设置
        void setSortAscending(bool ascending) { sortAscending_ = ascending; }
        bool isSortAscending() const { return sortAscending_; }

    private:
        std::string name_;                              ///< 字段名称
        std::string displayName_;                       ///< 显示名称
        PivotFieldType type_;                          ///< 字段类型
        PivotAggregateFunction aggregateFunction_;     ///< 聚合函数
        bool sortAscending_;                           ///< 是否升序排列
    };

    /**
     * @brief 透视表数据缓存类
     * 
     * 管理透视表的源数据和计算缓存
     */
    class TXPivotCache {
    public:
        /**
         * @brief 构造函数
         * @param sourceRange 源数据范围
         */
        explicit TXPivotCache(const TXRange& sourceRange);

        /**
         * @brief 析构函数
         */
        ~TXPivotCache() = default;

        // 源数据管理
        const TXRange& getSourceRange() const { return sourceRange_; }
        void setSourceRange(const TXRange& range) { sourceRange_ = range; }
        
        // 刷新数据
        void refresh();
        
        // 获取字段列表
        std::vector<std::string> getFieldNames() const;

        // 设置源工作表引用
        void setSourceSheet(const TXSheet* sheet);

        // 手动设置字段名称
        void setFieldNames(const std::vector<std::string>& fieldNames);

    private:
        TXRange sourceRange_;                          ///< 源数据范围
        std::vector<std::string> fieldNames_;         ///< 字段名称列表
        bool needsRefresh_;                            ///< 是否需要刷新
        const TXSheet* sourceSheet_;                   ///< 源工作表引用
    };

    /**
     * @brief 透视表类
     * 
     * 提供完整的Excel透视表功能，包括字段管理、数据聚合、样式设置等
     * 
     * @example
     * @code
     * // 创建透视表
     * auto pivotTable = sheet->createPivotTable("A1:D100", "F1");
     * 
     * // 添加字段
     * pivotTable->addRowField("产品类别");
     * pivotTable->addColumnField("销售月份");
     * pivotTable->addDataField("销售额", PivotAggregateFunction::Sum);
     * 
     * // 生成透视表
     * pivotTable->generate();
     * @endcode
     */
    class TXPivotTable {
    public:
        /**
         * @brief 构造函数
         * @param sourceRange 源数据范围
         * @param targetCell 透视表目标位置
         */
        TXPivotTable(const TXRange& sourceRange, const std::string& targetCell);

        /**
         * @brief 析构函数
         */
        ~TXPivotTable() = default;

        // 字段管理
        /**
         * @brief 添加行字段
         * @param fieldName 字段名称
         * @return 是否添加成功
         */
        bool addRowField(const std::string& fieldName);

        /**
         * @brief 添加列字段
         * @param fieldName 字段名称
         * @return 是否添加成功
         */
        bool addColumnField(const std::string& fieldName);

        /**
         * @brief 添加数据字段
         * @param fieldName 字段名称
         * @param aggregateFunc 聚合函数
         * @return 是否添加成功
         */
        bool addDataField(const std::string& fieldName, 
                         PivotAggregateFunction aggregateFunc = PivotAggregateFunction::Sum);

        /**
         * @brief 添加筛选字段
         * @param fieldName 字段名称
         * @return 是否添加成功
         */
        bool addFilterField(const std::string& fieldName);

        // 字段操作
        /**
         * @brief 移除字段
         * @param fieldName 字段名称
         * @return 是否移除成功
         */
        bool removeField(const std::string& fieldName);

        /**
         * @brief 获取字段
         * @param fieldName 字段名称
         * @return 字段指针，如果不存在返回nullptr
         */
        std::shared_ptr<TXPivotField> getField(const std::string& fieldName) const;

        /**
         * @brief 获取所有字段
         * @return 字段列表
         */
        const std::vector<std::shared_ptr<TXPivotField>>& getFields() const { return fields_; }

        /**
         * @brief 获取指定类型的字段
         * @param type 字段类型
         * @return 字段列表
         */
        std::vector<std::shared_ptr<TXPivotField>> getFieldsByType(PivotFieldType type) const;

        /**
         * @brief 获取数据缓存
         * @return 数据缓存指针
         */
        const TXPivotCache* getCache() const { return cache_.get(); }

        // 透视表操作
        /**
         * @brief 生成透视表
         * @return 是否生成成功
         */
        bool generate();

        /**
         * @brief 刷新透视表数据
         * @return 是否刷新成功
         */
        bool refresh();

        /**
         * @brief 将透视表集成到工作簿保存流程
         * @param zipWriter ZIP写入器
         * @param sheetName 工作表名称
         * @return 是否集成成功
         */
        bool integrateToWorkbook(class TXZipArchiveWriter& zipWriter, const std::string& sheetName);

        // 属性设置
        /**
         * @brief 设置透视表名称
         * @param name 名称
         */
        void setName(const std::string& name) { name_ = name; }

        /**
         * @brief 获取透视表名称
         * @return 透视表名称
         */
        const std::string& getName() const { return name_; }

        /**
         * @brief 设置目标位置
         * @param targetCell 目标单元格
         */
        void setTargetCell(const std::string& targetCell) { targetCell_ = targetCell; }

        /**
         * @brief 获取目标位置
         * @return 目标单元格
         */
        const std::string& getTargetCell() const { return targetCell_; }

    private:
        std::string name_;                                              ///< 透视表名称
        std::string targetCell_;                                        ///< 目标位置
        std::unique_ptr<TXPivotCache> cache_;                          ///< 数据缓存
        std::vector<std::shared_ptr<TXPivotField>> fields_;           ///< 字段列表
        std::unordered_map<std::string, std::shared_ptr<TXPivotField>> fieldMap_; ///< 字段映射

        // 内部方法
        bool validateField(const std::string& fieldName) const;
        void generatePivotTableXML();
        void generatePivotTableDefinitionXML();
        void generatePivotCacheDefinitionXML();
        void updateWorksheetRelationships();
        void calculateAggregates();

        // XML生成辅助方法
        std::string generatePivotCacheRecordsXml() const;
        std::string generateSheetRelationshipsXml() const;
        std::string generateWorkbookRelationshipsXml() const;
    };

} // namespace TinaXlsx
