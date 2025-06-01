#include "TinaXlsx/TXChartSeries.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include <stdexcept>

namespace TinaXlsx {

// ==================== TXChartSeries 实现 ====================

TXChartSeries::TXChartSeries(const std::string& name, const TXRange& dataRange, const std::string& color)
    : name_(name)
    , color_(color)
    , dataRange_(dataRange)
    , hasSeparateRanges_(false)
    , dataSheet_(nullptr)
    , showDataLabels_(false)
    , dataLabelFormat_("{value}")
{
}

TXChartSeries::TXChartSeries(const std::string& name, const TXRange& categoryRange, const TXRange& valueRange, const std::string& color)
    : name_(name)
    , color_(color)
    , categoryRange_(categoryRange)
    , valueRange_(valueRange)
    , hasSeparateRanges_(true)
    , dataSheet_(nullptr)
    , showDataLabels_(false)
    , dataLabelFormat_("{value}")
{
    // 构造一个包含类别和数值的总范围
    auto catStart = categoryRange.getStart();
    auto catEnd = categoryRange.getEnd();
    auto valStart = valueRange.getStart();
    auto valEnd = valueRange.getEnd();
    
    // 计算总范围的起始和结束坐标
    auto totalStart = TXCoordinate(
        std::min(catStart.getRow(), valStart.getRow()),
        std::min(catStart.getCol(), valStart.getCol())
    );
    auto totalEnd = TXCoordinate(
        std::max(catEnd.getRow(), valEnd.getRow()),
        std::max(catEnd.getCol(), valEnd.getCol())
    );
    
    dataRange_ = TXRange(totalStart, totalEnd);
}

// ==================== TXMultiSeriesChart 实现 ====================

TXMultiSeriesChart::TXMultiSeriesChart(const std::string& title)
    : title_(title)
{
}

u32 TXMultiSeriesChart::addSeries(const TXChartSeries& series) {
    series_.push_back(series);
    return static_cast<u32>(series_.size() - 1);
}

u32 TXMultiSeriesChart::addSeries(const std::string& name, const TXRange& dataRange, const std::string& color) {
    TXChartSeries series(name, dataRange, color);
    return addSeries(series);
}

u32 TXMultiSeriesChart::addSeries(const std::string& name, const TXRange& categoryRange, const TXRange& valueRange, const std::string& color) {
    TXChartSeries series(name, categoryRange, valueRange, color);
    return addSeries(series);
}

const TXChartSeries& TXMultiSeriesChart::getSeries(u32 index) const {
    if (index >= series_.size()) {
        throw std::out_of_range("Series index out of range");
    }
    return series_[index];
}

TXChartSeries& TXMultiSeriesChart::getSeries(u32 index) {
    if (index >= series_.size()) {
        throw std::out_of_range("Series index out of range");
    }
    return series_[index];
}

void TXMultiSeriesChart::removeSeries(u32 index) {
    if (index >= series_.size()) {
        throw std::out_of_range("Series index out of range");
    }
    series_.erase(series_.begin() + index);
}

void TXMultiSeriesChart::setDataSheet(const TXSheet* sheet) {
    for (auto& series : series_) {
        series.setDataSheet(sheet);
    }
}

} // namespace TinaXlsx
