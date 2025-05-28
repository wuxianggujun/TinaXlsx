//
// Created by wuxianggujun on 2025/5/28.
//

#include "TinaXlsx/TXChart.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include <algorithm>
#include <stdexcept>

namespace TinaXlsx
{
    class TXChartStyle::Impl
    {
    public:
        std::vector<TXColor> colorScheme;
        TXColor backgroundColor;
        BorderStyle borderStyle;
        TXColor borderColor;
        TXFont font;
        LegendPosition legendPosition;
        bool showDataLabels;
        bool showGridLines;
        bool smoothLines;
        bool enable3D;

        Impl(): backgroundColor(ColorConstants::WHITE),
                borderStyle(BorderStyle::None),
                borderColor(ColorConstants::BLACK),
                legendPosition(LegendPosition::Right),
                showDataLabels(true),
                showGridLines(true),
                smoothLines(false),
                enable3D(false)
        {
            // 默认颜色方案
            colorScheme = {
                TXColor(0x4472C4), // 蓝色
                TXColor(0xED7D31), // 橙色
                TXColor(0xA5A5A5), // 灰色
                TXColor(0xFFC000), // 黄色
                TXColor(0x5B9BD5), // 浅蓝色
                TXColor(0x70AD47), // 绿色
                TXColor(0x264478), // 深蓝色
                TXColor(0x9E480E) // 棕色
            };
        }
    };


    TXChartStyle::TXChartStyle(): m_pImpl(std::make_unique<Impl>())
    {
    }

    TXChartStyle::~TXChartStyle() = default;

  
    TXChartStyle::TXChartStyle(const TXChartStyle& other)
        : m_pImpl(std::make_unique<Impl>(*other.m_pImpl)) {}

    
    TXChartStyle& TXChartStyle::operator=(const TXChartStyle& other)
    {
        if (this != &other)
            *m_pImpl = *other.m_pImpl;
        return *this;
    }


    TXChartStyle::TXChartStyle(TXChartStyle&& other) noexcept = default;

    TXChartStyle& TXChartStyle::operator=(TXChartStyle&& other) noexcept = default;

    TXChartStyle& TXChartStyle::setColorScheme(const std::vector<TXColor>& colors)
    {
        if (!colors.empty())
        {
            m_pImpl->colorScheme = colors;
        }
        return *this;
    }

    TXChartStyle& TXChartStyle::setBackgroundColor(const TXColor& color)
    {
        m_pImpl->backgroundColor = color;
        return *this;
    }

    TXChartStyle& TXChartStyle::setBorderStyle(BorderStyle style, const TXColor& color)
    {
        m_pImpl->borderStyle = style;
        m_pImpl->borderColor = color;
        return *this;
    }

    TXChartStyle& TXChartStyle::setFont(const TXFont& font)
    {
        m_pImpl->font = font;
        return *this;
    }

    TXChartStyle& TXChartStyle::setLegendPosition(LegendPosition position)
    {
        m_pImpl->legendPosition = position;
        return *this;
    }

    TXChartStyle& TXChartStyle::showDataLabels(bool show)
    {
        m_pImpl->showDataLabels = show;
        return *this;
    }

    TXChartStyle& TXChartStyle::showGridLines(bool show)
    {
        m_pImpl->showGridLines = show;
        return *this;
    }

    TXChartStyle& TXChartStyle::setSmoothLines(bool smooth)
    {
        m_pImpl->smoothLines = smooth;
        return *this;
    }

    TXChartStyle& TXChartStyle::set3D(bool enable3D)
    {
        m_pImpl->enable3D = enable3D;
        return *this;
    }

    class TXChart::Impl
    {
    public:
        ChartType type;
        std::string name;
        std::string title;
        const TXSheet* dataSheet;
        TXRange dataRange;
        row_t positionRow;
        column_t positionCol;
        u32 width;
        u32 height;
        bool showLegend;
        bool showDataLabels;
        std::string xAxisTitle;
        std::string yAxisTitle;
        TXChartStyle style;

        Impl(ChartType chartType):
            type(chartType),
            name("Chart1"),
            title(""),
            dataSheet(nullptr),
            positionRow(row_t(0)),
            positionCol(column_t(row_t::index_t(0))),
            width(400),
            height(300),
            showLegend(true),
            showDataLabels(true),
            xAxisTitle(""),
            yAxisTitle("")
        {
        }
    };

    TXChart::TXChart(ChartType chartType): m_pImpl(std::make_unique<Impl>(chartType))
    {
    }

    TXChart::~TXChart() = default;
    
    TXChart::TXChart(TXChart&& other) noexcept = default;

    TXChart& TXChart::operator=(TXChart&& other) noexcept = default;

    bool TXChart::setDataRange(const TXSheet* sheet, const TXRange& range)
    {
        if (!sheet)
        {
            return false;
        }

        m_pImpl->dataSheet = sheet;
        m_pImpl->dataRange = range;
        return true;
    }

    bool TXChart::setDataRange(const TXSheet* sheet, const std::string& rangesAddress)
    {
        if (!sheet)
        {
            return false;
        }

        try
        {
            TXRange range = TXRange::fromAddress(rangesAddress);
            return setDataRange(sheet, range);
        }
        catch (const std::exception& e)
        {
            return false;
        }
    }

    void TXChart::setTitle(const std::string& title)
    {
        m_pImpl->title = title;
    }

    void TXChart::setPosition(row_t row, column_t col)
    {
        m_pImpl->positionRow = row;
        m_pImpl->positionCol = col;
    }

    void TXChart::setSize(u32 width, u32 height)
    {
        m_pImpl->width = width;
        m_pImpl->height = height;
    }

    ChartType TXChart::getType() const
    {
        return m_pImpl->type;
    }

    void TXChart::setShowLegend(bool show)
    {
        m_pImpl->showLegend = show;
    }

    void TXChart::setShowDataLabels(bool show)
    {
        m_pImpl->showDataLabels = show;
    }

    void TXChart::setAxisTitle(const std::string& title,bool isXAxis)
    {
        if (isXAxis)
        {
            m_pImpl->xAxisTitle = title;
        }
        else
        {
            m_pImpl->yAxisTitle = title;
        }
    }

    void TXChart::setStyle(const TXChartStyle& style)
    {
        m_pImpl->style = std::move(style);
    }

    bool TXChart::exportAsImage(const std::string& filePath)
    {
        return false;
    }

    const std::string& TXChart::getName() const
    {
        return m_pImpl->name;
    }

    void TXChart::setName(const std::string& name)
    {
        m_pImpl->name = name;
    }
    

    TXColumnChart::TXColumnChart()
        : TXChart(ChartType::Column)
    {
        // 初始化特有属性
    }

    void TXColumnChart::setBarWidth(f32 width) {
        // 实现柱状图特有属性设置
    }

    void TXColumnChart::setBarGap(f32 gap) {
        // 实现柱状图特有属性设置
    }

    void TXColumnChart::setStacked(bool stacked) {
        // 实现柱状图特有属性设置
    }

    void TXColumnChart::set3D(bool enable3D) {
        // 实现柱状图特有属性设置
    }

    // ==================== TXLineChart 实现 ====================

    TXLineChart::TXLineChart()
        : TXChart(ChartType::Line)
    {
        // 初始化特有属性
    }

    void TXLineChart::setSmoothLines(bool smooth) {
        // 实现折线图特有属性设置
    }

    void TXLineChart::setLineWidth(float width) {
        // 实现折线图特有属性设置
    }

    void TXLineChart::setShowMarkers(bool showMarkers) {
        // 实现折线图特有属性设置
    }

    void TXLineChart::setStacked(bool stacked) {
        // 实现折线图特有属性设置
    }

    // ==================== TXPieChart 实现 ====================

    TXPieChart::TXPieChart()
        : TXChart(ChartType::Pie)
    {
        // 初始化特有属性
    }

    void TXPieChart::setDoughnut(bool doughnut) {
        // 实现饼图特有属性设置
    }

    void TXPieChart::setDoughnutHoleSize(float ratio) {
        // 实现饼图特有属性设置
    }

    void TXPieChart::setFirstSliceAngle(int angle) {
        // 实现饼图特有属性设置
    }

    void TXPieChart::setExplodeSlice(int index, bool explode) {
        // 实现饼图特有属性设置
    }

    // ==================== TXScatterChart 实现 ====================

    TXScatterChart::TXScatterChart()
        : TXChart(ChartType::Scatter)
    {
        // 初始化特有属性
    }

    void TXScatterChart::setShowTrendLine(bool show) {
        // 实现散点图特有属性设置
    }

    void TXScatterChart::setTrendLineType(TrendLineType type) {
        // 实现散点图特有属性设置
    }

    void TXScatterChart::setMarkerSize(float size) {
        // 实现散点图特有属性设置
    }
    
} // TinaXlsx
