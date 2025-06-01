#include "TinaXlsx/TXChartSeriesBuilder.hpp"
#include "TinaXlsx/TXRangeFormatter.hpp"
#include "TinaXlsx/TXSheet.hpp"

namespace TinaXlsx {

// ==================== TXChartSeriesBuilder 基类实现 ====================

XmlNodeBuilder TXChartSeriesBuilder::buildBasicSeriesInfo(const TXChart* chart, u32 seriesIndex) const {
    XmlNodeBuilder ser("c:ser");

    // 系列索引
    XmlNodeBuilder idx("c:idx");
    idx.addAttribute("val", std::to_string(seriesIndex));
    ser.addChild(idx);

    // 系列顺序
    XmlNodeBuilder order("c:order");
    order.addAttribute("val", std::to_string(seriesIndex));
    ser.addChild(order);

    // 系列文本（使用图表标题）
    XmlNodeBuilder tx("c:tx");
    XmlNodeBuilder v("c:v");
    v.setText(chart->getTitle().empty() ? ("系列" + std::to_string(seriesIndex + 1)) : chart->getTitle());
    tx.addChild(v);
    ser.addChild(tx);

    return ser;
}

std::string TXChartSeriesBuilder::getSheetName(const TXChart* chart) const {
    return chart->getDataSheet() ? chart->getDataSheet()->getName() : "Sheet1";
}

// ==================== TXColumnSeriesBuilder 实现 ====================

XmlNodeBuilder TXColumnSeriesBuilder::buildSeries(const TXChart* chart, u32 seriesIndex) const {
    XmlNodeBuilder ser = buildBasicSeriesInfo(chart, seriesIndex);

    const TXRange& dataRange = chart->getDataRange();
    std::string sheetName = getSheetName(chart);

    // 类别轴数据（标签）
    XmlNodeBuilder cat("c:cat");
    XmlNodeBuilder strRef("c:strRef");
    XmlNodeBuilder catF("c:f");
    catF.setText(TXRangeFormatter::formatCategoryRange(dataRange, sheetName));
    strRef.addChild(catF);
    cat.addChild(strRef);
    ser.addChild(cat);

    // 数据值
    XmlNodeBuilder val("c:val");
    XmlNodeBuilder numRef("c:numRef");
    XmlNodeBuilder f("c:f");
    f.setText(TXRangeFormatter::formatValueRange(dataRange, sheetName));
    numRef.addChild(f);
    val.addChild(numRef);
    ser.addChild(val);

    // 添加柱状图格式化（使用Office主题蓝色）
    addColumnFormatting(ser, "4F81BD");

    return ser;
}

XmlNodeBuilder TXColumnSeriesBuilder::buildSeries(const TXChartSeries& series, u32 seriesIndex, const TXChartStyleV2& style) const {
    // TODO: 实现多系列构建
    // 暂时返回空节点，后续实现
    return XmlNodeBuilder("c:ser");
}

void TXColumnSeriesBuilder::addColumnFormatting(XmlNodeBuilder& ser, const std::string& color) const {
    // 柱状图形状属性
    XmlNodeBuilder spPr("c:spPr");
    XmlNodeBuilder solidFill("a:solidFill");
    XmlNodeBuilder srgbClr("a:srgbClr");
    srgbClr.addAttribute("val", color);
    solidFill.addChild(srgbClr);
    spPr.addChild(solidFill);
    ser.addChild(spPr);
}

// ==================== TXLineSeriesBuilder 实现 ====================

XmlNodeBuilder TXLineSeriesBuilder::buildSeries(const TXChart* chart, u32 seriesIndex) const {
    XmlNodeBuilder ser = buildBasicSeriesInfo(chart, seriesIndex);

    const TXRange& dataRange = chart->getDataRange();
    std::string sheetName = getSheetName(chart);

    // 类别轴数据（标签）
    XmlNodeBuilder cat("c:cat");
    XmlNodeBuilder strRef("c:strRef");
    XmlNodeBuilder catF("c:f");
    catF.setText(TXRangeFormatter::formatCategoryRange(dataRange, sheetName));
    strRef.addChild(catF);
    cat.addChild(strRef);
    ser.addChild(cat);

    // 数据值
    XmlNodeBuilder val("c:val");
    XmlNodeBuilder numRef("c:numRef");
    XmlNodeBuilder f("c:f");
    f.setText(TXRangeFormatter::formatValueRange(dataRange, sheetName));
    numRef.addChild(f);
    val.addChild(numRef);
    ser.addChild(val);

    // 添加折线图格式化（使用彩色主题红色）
    addLineFormatting(ser, "FF6B6B", 25400);

    return ser;
}

XmlNodeBuilder TXLineSeriesBuilder::buildSeries(const TXChartSeries& series, u32 seriesIndex, const TXChartStyleV2& style) const {
    // TODO: 实现多系列构建
    // 暂时返回空节点，后续实现
    return XmlNodeBuilder("c:ser");
}

void TXLineSeriesBuilder::addLineFormatting(XmlNodeBuilder& ser, const std::string& color, u32 lineWidth) const {
    // 折线图线条属性
    XmlNodeBuilder spPr("c:spPr");
    XmlNodeBuilder ln("a:ln");
    ln.addAttribute("w", std::to_string(lineWidth)); // 线条宽度
    XmlNodeBuilder solidFill("a:solidFill");
    XmlNodeBuilder srgbClr("a:srgbClr");
    srgbClr.addAttribute("val", color);
    solidFill.addChild(srgbClr);
    ln.addChild(solidFill);
    spPr.addChild(ln);
    ser.addChild(spPr);
}

// ==================== TXPieSeriesBuilder 实现 ====================

XmlNodeBuilder TXPieSeriesBuilder::buildSeries(const TXChart* chart, u32 seriesIndex) const {
    XmlNodeBuilder ser = buildBasicSeriesInfo(chart, seriesIndex);

    const TXRange& dataRange = chart->getDataRange();
    std::string sheetName = getSheetName(chart);

    // 饼图类别轴数据（标签）
    XmlNodeBuilder cat("c:cat");
    XmlNodeBuilder strRef("c:strRef");
    XmlNodeBuilder catF("c:f");
    catF.setText(TXRangeFormatter::formatPieLabelRange(dataRange, sheetName));
    strRef.addChild(catF);
    cat.addChild(strRef);
    ser.addChild(cat);

    // 饼图数据值
    XmlNodeBuilder val("c:val");
    XmlNodeBuilder numRef("c:numRef");
    XmlNodeBuilder f("c:f");
    f.setText(TXRangeFormatter::formatPieValueRange(dataRange, sheetName));
    numRef.addChild(f);
    val.addChild(numRef);
    ser.addChild(val);

    // 添加饼图格式化（使用单色主题深灰色）
    addPieFormatting(ser, "2C3E50");

    return ser;
}

XmlNodeBuilder TXPieSeriesBuilder::buildSeries(const TXChartSeries& series, u32 seriesIndex, const TXChartStyleV2& style) const {
    // TODO: 实现多系列构建
    // 暂时返回空节点，后续实现
    return XmlNodeBuilder("c:ser");
}

void TXPieSeriesBuilder::addPieFormatting(XmlNodeBuilder& ser, const std::string& color) const {
    // 为饼图添加颜色格式化
    XmlNodeBuilder spPr("c:spPr");
    XmlNodeBuilder solidFill("a:solidFill");
    XmlNodeBuilder srgbClr("a:srgbClr");
    srgbClr.addAttribute("val", color);
    solidFill.addChild(srgbClr);
    spPr.addChild(solidFill);
    ser.addChild(spPr);
}

// ==================== TXScatterSeriesBuilder 实现 ====================

XmlNodeBuilder TXScatterSeriesBuilder::buildSeries(const TXChart* chart, u32 seriesIndex) const {
    XmlNodeBuilder ser = buildBasicSeriesInfo(chart, seriesIndex);

    const TXRange& dataRange = chart->getDataRange();
    std::string sheetName = getSheetName(chart);

    // 散点图X值
    XmlNodeBuilder xVal("c:xVal");
    XmlNodeBuilder numRefX("c:numRef");
    XmlNodeBuilder fX("c:f");
    fX.setText(TXRangeFormatter::formatScatterXRange(dataRange, sheetName));
    numRefX.addChild(fX);
    xVal.addChild(numRefX);
    ser.addChild(xVal);

    // 散点图Y值
    XmlNodeBuilder yVal("c:yVal");
    XmlNodeBuilder numRefY("c:numRef");
    XmlNodeBuilder fY("c:f");
    fY.setText(TXRangeFormatter::formatScatterYRange(dataRange, sheetName));
    numRefY.addChild(fY);
    yVal.addChild(numRefY);
    ser.addChild(yVal);

    // 添加散点图格式化
    addScatterFormatting(ser, "4F81BD");

    return ser;
}

XmlNodeBuilder TXScatterSeriesBuilder::buildSeries(const TXChartSeries& series, u32 seriesIndex, const TXChartStyleV2& style) const {
    // TODO: 实现多系列构建
    // 暂时返回空节点，后续实现
    return XmlNodeBuilder("c:ser");
}

void TXScatterSeriesBuilder::addScatterFormatting(XmlNodeBuilder& ser, const std::string& color) const {
    // 散点图标记属性
    XmlNodeBuilder marker("c:marker");
    XmlNodeBuilder symbol("c:symbol");
    symbol.addAttribute("val", "circle");
    marker.addChild(symbol);
    XmlNodeBuilder size("c:size");
    size.addAttribute("val", "5");
    marker.addChild(size);

    // 添加标记颜色
    XmlNodeBuilder spPr("c:spPr");
    XmlNodeBuilder solidFill("a:solidFill");
    XmlNodeBuilder srgbClr("a:srgbClr");
    srgbClr.addAttribute("val", color);
    solidFill.addChild(srgbClr);
    spPr.addChild(solidFill);
    marker.addChild(spPr);

    ser.addChild(marker);
}

// ==================== TXSeriesBuilderFactory 实现 ====================

std::unique_ptr<TXChartSeriesBuilder> TXSeriesBuilderFactory::createBuilder(ChartType chartType) {
    switch (chartType) {
        case ChartType::Column:
            return std::make_unique<TXColumnSeriesBuilder>();
        case ChartType::Line:
            return std::make_unique<TXLineSeriesBuilder>();
        case ChartType::Pie:
            return std::make_unique<TXPieSeriesBuilder>();
        case ChartType::Scatter:
            return std::make_unique<TXScatterSeriesBuilder>();
        default:
            return std::make_unique<TXColumnSeriesBuilder>(); // 默认为柱状图
    }
}

} // namespace TinaXlsx
