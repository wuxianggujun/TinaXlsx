//
// Created by wuxianggujun on 2025/5/28.
//

#pragma once

#include "TXTypes.hpp"
#include "TXColor.hpp"
#include "TXRange.hpp"
#include "TXStyle.hpp"

namespace TinaXlsx
{
    class TXSheet;

    enum class ChartType :u8
    {
        Column,
        Bar,
        Line,
        Pie,
        Scatter,
        Area,
        Radar,
        Combo
    };

    enum class LegendPosition :u8
    {
        Top,
        Bottom,
        Left,
        Right,
        TopRight,
        TopLeft,
        BottomRight,
        BottomLeft
    };

    class TXChartStyle
    {
    public:
        TXChartStyle();
        ~TXChartStyle();

        TXChartStyle(const TXChartStyle& other);
        TXChartStyle& operator=(const TXChartStyle& other);

        TXChartStyle(TXChartStyle&& other) noexcept;
        TXChartStyle& operator=(TXChartStyle&& other) noexcept;

        TXChartStyle& setColorScheme(const std::vector<TXColor>& colors);

        TXChartStyle& setBackgroundColor(const TXColor& color);

        TXChartStyle& setBorderStyle(BorderStyle style, const TXColor& color = ColorConstants::BLACK);

        TXChartStyle& setFont(const TXFont& font);

        TXChartStyle& setLegendPosition(LegendPosition position);

        TXChartStyle& showDataLabels(bool show = true);

        TXChartStyle& showGridLines(bool show = true);

        TXChartStyle& setSmoothLines(bool smooth = true);

        TXChartStyle& set3D(bool enable3D = true);

    private:
        class Impl;
        std::unique_ptr<Impl> m_pImpl;
    };


    class TXChart
    {
    public:
        explicit TXChart(ChartType type);
        virtual ~TXChart();

        TXChart(const TXChart& other) = delete;
        TXChart& operator=(const TXChart& other) = delete;
        TXChart(TXChart&& other) noexcept;
        TXChart& operator=(TXChart&& other) noexcept;

        bool setDataRange(const TXSheet* sheet, const TXRange& range);

        bool setDataRange(const TXSheet* sheet, const std::string& rangesAddress);

        /**
         * @brief 获取数据工作表
         * @return 数据工作表指针，如果没有设置返回nullptr
         */
        const TXSheet* getDataSheet() const;

        /**
         * @brief 获取数据范围
         * @return 数据范围
         */
        const TXRange& getDataRange() const;

        void setTitle(const std::string& title);

        /**
         * @brief 获取图表标题
         * @return 图表标题
         */
        const std::string& getTitle() const;

        void setPosition(row_t row, column_t col);

        void setSize(u32 width, u32 height);

        ChartType getType() const;

        void setShowLegend(bool show);

        void setShowDataLabels(bool show);

        void setAxisTitle(const std::string& title,bool isXAxis = true);

        void setStyle(const TXChartStyle& style);

        bool exportAsImage(const std::string& filePath);

        const std::string& getName() const;

        void setName(const std::string& name);

    protected:
        class Impl;
        std::unique_ptr<Impl> m_pImpl;
    };

    class TXColumnChart : public TXChart
    {
    public:
        TXColumnChart();

        void setBarWidth(f32 width);

        void setBarGap(f32 gap);

        void setStacked(bool stacked);

        void set3D(bool enable3D);
    };


    class TXLineChart : public TXChart
    {
    public:
        TXLineChart();

        void setSmoothLines(bool smooth);

        void setLineWidth(f32 width);

        void setShowMarkers(bool showMarkers);

        void setStacked(bool stacked);
    };

    class TXPieChart : public TXChart
    {
    public:
        TXPieChart();

        void setDoughnut(bool doughnut);

        void setDoughnutHoleSize(f32 size);

        void setFirstSliceAngle(int angle);

        void setExplodeSlice(int index, bool explode = true);
    };

    class TXScatterChart : public TXChart
    {
    public:
        enum class TrendLineType
        {
            Linear,
            Exponential,
            Logarithmic,
            Polynomial,
            Power,
            MovingAverage
        };

        TXScatterChart();

        void setShowTrendLine(bool show);

        void setTrendLineType(TrendLineType type);

        void setMarkerSize(f32 size);
    };
} // TinaXlsx
