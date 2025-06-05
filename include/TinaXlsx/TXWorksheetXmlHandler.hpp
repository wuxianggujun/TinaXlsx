//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include "TXXmlHandler.hpp"
#include "TXXmlReader.hpp"
#include "TXXmlWriter.hpp"
#include "TXSheet.hpp"
#include "TXCompactCell.hpp"
#include "TXRange.hpp"
#include "TXTypes.hpp"
#include "TXStreamXmlReader.hpp"
#include "TXSIMDXmlParser.hpp"
#include "TXBatchXMLGenerator.hpp"
#include "TXUnifiedMemoryManager.hpp"
#include <sstream>
#include <iomanip>
#include <memory>
#include <vector>

namespace TinaXlsx
{
    // 前向声明
    class TXPivotTable;

    class TXWorksheetXmlHandler : public TXXmlHandler
    {
    public:
        explicit TXWorksheetXmlHandler(u64 sheetIndex): m_sheetIndex(sheetIndex)
        {
            // 初始化批处理XML生成器
            initializeBatchGenerator();
        }

        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override
        {
            // 🚀 性能优化：使用SIMD优化的XML解析器
            if (m_sheetIndex >= context.sheets.size()) {
                return Err<void>(TXErrorCode::InvalidArgument, "Sheet index out of range");
            }

            // 读取XML数据
            auto xmlData = zipReader.read(partName());
            if (xmlData.isError()) {
                return Err<void>(xmlData.error().getCode(), "Failed to read " + partName());
            }

            const std::vector<uint8_t>& fileBytes = xmlData.value();
            std::string xmlContent(fileBytes.begin(), fileBytes.end());

            // 🚀 使用SIMD优化的解析器
            TXSIMDWorksheetParser parser(context.sheets[m_sheetIndex].get());
            size_t cellCount = parser.parse(xmlContent);

            // 输出统计信息（调试用）
            const auto& stats = parser.getStats();
            // TODO: 添加日志系统后输出统计信息
            // printf("SIMD解析: %zu行, %zu单元格, %.2fms\n",
            //        stats.totalRows, stats.totalCells, stats.parseTimeMs);

            return Ok();
        }

        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override;

        [[nodiscard]] std::string partName() const override {
            return "xl/worksheets/sheet" + std::to_string(m_sheetIndex + 1) + ".xml";
        }

        /**
         * @brief 设置透视表信息（用于传递透视表数据）
         * @param pivotTables 透视表列表
         */
        void setPivotTables(const std::vector<std::shared_ptr<class TXPivotTable>>& pivotTables);

    private:
        bool shouldUseInlineString(const std::string& str) const;

        /**
         * @brief 初始化批处理XML生成器
         */
        void initializeBatchGenerator();

        /**
         * @brief 使用批处理器保存（统一高性能版本）
         * @param zipWriter ZIP写入器
         * @param context 工作簿上下文
         * @return 保存结果
         */
        TXResult<void> saveWithBatchProcessor(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context);

        /**
         * @brief 使用流式写入器保存（将被废弃）
         * @param zipWriter ZIP写入器
         * @param context 工作簿上下文
         * @return 保存结果
         */
        TXResult<void> saveWithStreamWriter(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context);

        /**
         * @brief 构建单个单元格节点
         * @param cell 单元格对象
         * @param cellRef 单元格引用（如A1）
         * @param context 工作簿上下文
         * @return 单元格节点
         */
        XmlNodeBuilder buildCellNode(const TXCompactCell* cell, const std::string& cellRef,const TXWorkbookContext& context) const;

        /**
         * @brief 构建数据验证节点
         * @param sheet 工作表对象
         * @return 数据验证节点
         */
        XmlNodeBuilder buildDataValidationsNode(const TXSheet* sheet) const;

        /**
         * @brief 构建自动筛选节点
         * @param sheet 工作表对象
         * @return 自动筛选节点
         */
        XmlNodeBuilder buildAutoFilterNode(const TXSheet* sheet) const;

        /**
         * @brief 获取工作表的透视表列表
         * @param sheetName 工作表名称
         * @param context 工作簿上下文
         * @return 透视表列表
         */
        std::vector<std::shared_ptr<class TXPivotTable>> getPivotTablesForSheet(const std::string& sheetName, const TXWorkbookContext& context) const;

    private:
        std::vector<std::shared_ptr<class TXPivotTable>> m_pivotTables;  ///< 当前工作表的透视表列表

        u64 m_sheetIndex;

        // 🚀 批处理组件
        std::unique_ptr<TXUnifiedMemoryManager> memory_manager_;  ///< 内存管理器
        std::unique_ptr<TXBatchXMLGenerator> batch_generator_;    ///< 批处理XML生成器
    };
}
