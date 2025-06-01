#include "TinaXlsx/TXChartStyle.hpp"

namespace TinaXlsx {

// ==================== TXChartStyleV2 实现 ====================

TXChartStyleV2::TXChartStyleV2()
    : theme_(ChartTheme::Office)
    , primaryColor_("4F81BD")
    , lineWidth_(25400)  // 1pt = 12700 EMU, 2pt = 25400 EMU
    , borderColor_("000000")
{
    applyOfficeTheme();
}

TXChartStyleV2::TXChartStyleV2(ChartTheme theme)
    : theme_(theme)
    , primaryColor_("4F81BD")
    , lineWidth_(25400)
    , borderColor_("000000")
{
    switch (theme) {
        case ChartTheme::Office:
            applyOfficeTheme();
            break;
        case ChartTheme::Colorful:
            applyColorfulTheme();
            break;
        case ChartTheme::Monochromatic:
            applyMonochromaticTheme();
            break;
        case ChartTheme::Custom:
            // 自定义主题保持默认值
            break;
    }
}

void TXChartStyleV2::setPrimaryColor(const std::string& color) {
    primaryColor_ = color;
}

void TXChartStyleV2::setSeriesColors(const std::vector<std::string>& colors) {
    seriesColors_ = colors;
}

std::string TXChartStyleV2::getSeriesColor(u32 index) const {
    if (seriesColors_.empty()) {
        // 如果没有设置系列颜色，返回主色调
        return primaryColor_;
    }

    // 循环使用颜色列表
    return seriesColors_[index % seriesColors_.size()];
}

void TXChartStyleV2::setTitleFont(const TXFont& font) {
    titleFont_ = font;
}

void TXChartStyleV2::setAxisLabelFont(const TXFont& font) {
    axisLabelFont_ = font;
}

void TXChartStyleV2::setDataLabelFont(const TXFont& font) {
    dataLabelFont_ = font;
}

void TXChartStyleV2::setLineWidth(u32 width) {
    lineWidth_ = width;
}

void TXChartStyleV2::setBorderColor(const std::string& color) {
    borderColor_ = color;
}

void TXChartStyleV2::applyOfficeTheme() {
    theme_ = ChartTheme::Office;
    primaryColor_ = "4F81BD";

    // Office主题的系列颜色
    seriesColors_ = {
        "4F81BD",  // 蓝色
        "F79646",  // 橙色
        "9CBB58",  // 绿色
        "8064A2",  // 紫色
        "4BACC6",  // 青色
        "F24726",  // 红色
        "9BBB59",  // 橄榄绿
        "8C564B",  // 棕色
        "17BECF",  // 浅青色
        "BCBD22"   // 黄绿色
    };

    // 字体配置
    titleFont_ = TXFont("Arial", 14).setBold(true);
    axisLabelFont_ = TXFont("Arial", 10);
    dataLabelFont_ = TXFont("Arial", 9);

    lineWidth_ = 25400;  // 2pt
    borderColor_ = "000000";
}

void TXChartStyleV2::applyColorfulTheme() {
    theme_ = ChartTheme::Colorful;
    primaryColor_ = "FF6B6B";

    // 彩色主题的系列颜色
    seriesColors_ = {
        "FF6B6B",  // 红色
        "4ECDC4",  // 青色
        "45B7D1",  // 蓝色
        "96CEB4",  // 绿色
        "FFEAA7",  // 黄色
        "DDA0DD",  // 紫色
        "98D8C8",  // 薄荷绿
        "F7DC6F",  // 金色
        "BB8FCE",  // 淡紫色
        "85C1E9"   // 天蓝色
    };

    // 字体配置
    titleFont_ = TXFont("Arial", 16).setBold(true);
    axisLabelFont_ = TXFont("Arial", 11);
    dataLabelFont_ = TXFont("Arial", 10);

    lineWidth_ = 31750;  // 2.5pt
    borderColor_ = "2C3E50";
}

void TXChartStyleV2::applyMonochromaticTheme() {
    theme_ = ChartTheme::Monochromatic;
    primaryColor_ = "2C3E50";

    // 单色主题的系列颜色（不同深浅的灰色）
    seriesColors_ = {
        "2C3E50",  // 深灰
        "34495E",  // 灰色
        "5D6D7E",  // 中灰
        "85929E",  // 浅灰
        "AEB6BF",  // 很浅灰
        "D5DBDB",  // 极浅灰
        "1B2631",  // 极深灰
        "273746",  // 深蓝灰
        "566573",  // 蓝灰
        "839192"   // 浅蓝灰
    };

    // 字体配置
    titleFont_ = TXFont("Arial", 14).setBold(true);
    axisLabelFont_ = TXFont("Arial", 10);
    dataLabelFont_ = TXFont("Arial", 9);

    lineWidth_ = 19050;  // 1.5pt
    borderColor_ = "2C3E50";
}

// ==================== TXChartConfig 实现 ====================

TXChartConfig::TXChartConfig()
    : showLegend_(true)
    , showDataLabels_(false)
    , showGridlines_(true)
    , xAxisTitle_("")
    , yAxisTitle_("")
    , style_(ChartTheme::Office)
{
}

} // namespace TinaXlsx
