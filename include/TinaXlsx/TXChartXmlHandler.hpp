//
// Created by wuxianggujun on 2025/1/15.
//

#pragma once

#include "TXXmlHandler.hpp"
#include "TXChart.hpp"
#include "TXXmlWriter.hpp"
#include <memory>

namespace TinaXlsx
{
    /**
     * @brief 图表XML处理器
     * 
     * 负责处理Excel图表的XML生成和解析
     */
    class TXChartXmlHandler : public TXXmlHandler
    {
    public:
        /**
         * @brief 构造函数
         * @param chart 图表对象
         * @param chartIndex 图表索引
         */
        explicit TXChartXmlHandler(const TXChart* chart, u32 chartIndex);

        /**
         * @brief 加载图表XML（暂不实现）
         */
        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override;

        /**
         * @brief 保存图表XML
         */
        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override;

        /**
         * @brief 获取图表XML文件路径
         */
        [[nodiscard]] std::string partName() const override;

    private:
        const TXChart* m_chart;
        u32 m_chartIndex;

        /**
         * @brief 生成图表XML节点
         */
        XmlNodeBuilder generateChartXml() const;

        /**
         * @brief 生成柱状图XML
         */
        XmlNodeBuilder generateColumnChartXml() const;

        /**
         * @brief 生成折线图XML
         */
        XmlNodeBuilder generateLineChartXml() const;

        /**
         * @brief 生成饼图XML
         */
        XmlNodeBuilder generatePieChartXml() const;

        /**
         * @brief 生成散点图XML
         */
        XmlNodeBuilder generateScatterChartXml() const;

        /**
         * @brief 生成图表标题XML
         */
        XmlNodeBuilder generateTitleXml() const;

        /**
         * @brief 生成图表图例XML
         */
        XmlNodeBuilder generateLegendXml() const;

        /**
         * @brief 生成图表数据系列XML
         */
        XmlNodeBuilder generateSeriesXml() const;

        /**
         * @brief 生成类别轴XML
         */
        XmlNodeBuilder generateCategoryAxisXml() const;

        /**
         * @brief 生成数值轴XML
         */
        XmlNodeBuilder generateValueAxisXml() const;

        /**
         * @brief 生成绘图区域XML
         */
        XmlNodeBuilder generatePlotAreaXml() const;

        /**
         * @brief 获取图表类型字符串
         */
        std::string getChartTypeString() const;
    };

    /**
     * @brief 图表关系XML处理器
     * 
     * 处理图表的关系文件
     */
    class TXChartRelsXmlHandler : public TXXmlHandler
    {
    public:
        explicit TXChartRelsXmlHandler(u32 chartIndex);

        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override;
        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override;
        [[nodiscard]] std::string partName() const override;

    private:
        u32 m_chartIndex;
    };

    /**
     * @brief 绘图XML处理器
     *
     * 处理工作表中的绘图对象
     */
    class TXDrawingXmlHandler : public TXXmlHandler
    {
    public:
        explicit TXDrawingXmlHandler(u32 sheetIndex);

        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override;
        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override;
        [[nodiscard]] std::string partName() const override;

    private:
        u32 m_sheetIndex;
    };

    /**
     * @brief 绘图关系XML处理器
     *
     * 处理绘图的关系文件
     */
    class TXDrawingRelsXmlHandler : public TXXmlHandler
    {
    public:
        explicit TXDrawingRelsXmlHandler(u32 sheetIndex);

        TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override;
        TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override;
        [[nodiscard]] std::string partName() const override;

    private:
        u32 m_sheetIndex;
    };

} // namespace TinaXlsx
