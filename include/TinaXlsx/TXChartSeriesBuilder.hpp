#pragma once

#include "TXTypes.hpp"
#include "TXXmlWriter.hpp"
#include "TXChart.hpp"
#include "TXChartStyle.hpp"
#include "TXChartSeries.hpp"
#include <memory>

namespace TinaXlsx {

/**
 * @brief 图表系列构建器基类
 * 
 * 负责为不同类型的图表生成相应的数据系列XML
 */
class TXChartSeriesBuilder {
public:
    virtual ~TXChartSeriesBuilder() = default;

    /**
     * @brief 构建图表系列XML节点（单系列）
     * @param chart 图表对象
     * @param seriesIndex 系列索引（从0开始）
     * @return 系列XML节点
     */
    virtual XmlNodeBuilder buildSeries(const TXChart* chart, u32 seriesIndex = 0) const = 0;

    /**
     * @brief 构建图表系列XML节点（多系列）
     * @param series 数据系列
     * @param seriesIndex 系列索引
     * @param style 图表样式
     * @return 系列XML节点
     */
    virtual XmlNodeBuilder buildSeries(const TXChartSeries& series, u32 seriesIndex, const TXChartStyleV2& style) const = 0;

protected:
    /**
     * @brief 构建基础系列信息（索引、顺序、标题）
     * @param chart 图表对象
     * @param seriesIndex 系列索引
     * @return 包含基础信息的系列节点
     */
    XmlNodeBuilder buildBasicSeriesInfo(const TXChart* chart, u32 seriesIndex) const;

    /**
     * @brief 获取工作表名称
     * @param chart 图表对象
     * @return 工作表名称
     */
    std::string getSheetName(const TXChart* chart) const;
};

/**
 * @brief 柱状图系列构建器
 */
class TXColumnSeriesBuilder : public TXChartSeriesBuilder {
public:
    XmlNodeBuilder buildSeries(const TXChart* chart, u32 seriesIndex = 0) const override;
    XmlNodeBuilder buildSeries(const TXChartSeries& series, u32 seriesIndex, const TXChartStyleV2& style) const override;

private:
    /**
     * @brief 添加柱状图特有的格式化属性
     * @param ser 系列节点
     * @param color 系列颜色
     */
    void addColumnFormatting(XmlNodeBuilder& ser, const std::string& color) const;
};

/**
 * @brief 折线图系列构建器
 */
class TXLineSeriesBuilder : public TXChartSeriesBuilder {
public:
    XmlNodeBuilder buildSeries(const TXChart* chart, u32 seriesIndex = 0) const override;
    XmlNodeBuilder buildSeries(const TXChartSeries& series, u32 seriesIndex, const TXChartStyleV2& style) const override;

private:
    /**
     * @brief 添加折线图特有的格式化属性
     * @param ser 系列节点
     * @param color 系列颜色
     * @param lineWidth 线条宽度
     */
    void addLineFormatting(XmlNodeBuilder& ser, const std::string& color, u32 lineWidth) const;
};

/**
 * @brief 饼图系列构建器
 */
class TXPieSeriesBuilder : public TXChartSeriesBuilder {
public:
    XmlNodeBuilder buildSeries(const TXChart* chart, u32 seriesIndex = 0) const override;
    XmlNodeBuilder buildSeries(const TXChartSeries& series, u32 seriesIndex, const TXChartStyleV2& style) const override;

private:
    /**
     * @brief 添加饼图特有的格式化属性
     * @param ser 系列节点
     * @param color 系列颜色
     */
    void addPieFormatting(XmlNodeBuilder& ser, const std::string& color) const;
};

/**
 * @brief 散点图系列构建器
 */
class TXScatterSeriesBuilder : public TXChartSeriesBuilder {
public:
    XmlNodeBuilder buildSeries(const TXChart* chart, u32 seriesIndex = 0) const override;
    XmlNodeBuilder buildSeries(const TXChartSeries& series, u32 seriesIndex, const TXChartStyleV2& style) const override;

private:
    /**
     * @brief 添加散点图特有的格式化属性
     * @param ser 系列节点
     * @param color 系列颜色
     */
    void addScatterFormatting(XmlNodeBuilder& ser, const std::string& color) const;
};

/**
 * @brief 系列构建器工厂
 */
class TXSeriesBuilderFactory {
public:
    /**
     * @brief 根据图表类型创建相应的系列构建器
     * @param chartType 图表类型
     * @return 系列构建器智能指针
     */
    static std::unique_ptr<TXChartSeriesBuilder> createBuilder(ChartType chartType);
};

} // namespace TinaXlsx
