#pragma once

#include "TXTypes.hpp"
#include "TXFont.hpp"
#include <string>
#include <vector>

namespace TinaXlsx {

/**
 * @brief 图表主题枚举
 */
enum class ChartTheme {
    Office,         ///< Office默认主题
    Colorful,       ///< 彩色主题
    Monochromatic,  ///< 单色主题
    Custom          ///< 自定义主题
};

/**
 * @brief 图表样式配置（第二阶段增强版）
 */
class TXChartStyleV2 {
public:
    /**
     * @brief 默认构造函数
     */
    TXChartStyleV2();

    /**
     * @brief 从主题创建样式
     * @param theme 主题类型
     */
    explicit TXChartStyleV2(ChartTheme theme);

    // ==================== 颜色配置 ====================

    /**
     * @brief 设置主色调
     * @param color 颜色（十六进制，如 "4F81BD"）
     */
    void setPrimaryColor(const std::string& color);

    /**
     * @brief 获取主色调
     */
    const std::string& getPrimaryColor() const { return primaryColor_; }

    /**
     * @brief 设置系列颜色列表
     * @param colors 颜色列表
     */
    void setSeriesColors(const std::vector<std::string>& colors);

    /**
     * @brief 获取系列颜色列表
     */
    const std::vector<std::string>& getSeriesColors() const { return seriesColors_; }

    /**
     * @brief 获取指定索引的系列颜色
     * @param index 系列索引
     * @return 颜色字符串
     */
    std::string getSeriesColor(u32 index) const;

    // ==================== 字体配置 ====================

    /**
     * @brief 设置标题字体
     * @param font 字体配置
     */
    void setTitleFont(const TXFont& font);

    /**
     * @brief 获取标题字体
     */
    const TXFont& getTitleFont() const { return titleFont_; }

    /**
     * @brief 设置轴标签字体
     * @param font 字体配置
     */
    void setAxisLabelFont(const TXFont& font);

    /**
     * @brief 获取轴标签字体
     */
    const TXFont& getAxisLabelFont() const { return axisLabelFont_; }

    /**
     * @brief 设置数据标签字体
     * @param font 字体配置
     */
    void setDataLabelFont(const TXFont& font);

    /**
     * @brief 获取数据标签字体
     */
    const TXFont& getDataLabelFont() const { return dataLabelFont_; }

    // ==================== 线条和边框配置 ====================

    /**
     * @brief 设置线条宽度
     * @param width 宽度（EMU单位）
     */
    void setLineWidth(u32 width);

    /**
     * @brief 获取线条宽度
     */
    u32 getLineWidth() const { return lineWidth_; }

    /**
     * @brief 设置边框颜色
     * @param color 颜色（十六进制）
     */
    void setBorderColor(const std::string& color);

    /**
     * @brief 获取边框颜色
     */
    const std::string& getBorderColor() const { return borderColor_; }

    // ==================== 预定义主题 ====================

    /**
     * @brief 应用Office主题
     */
    void applyOfficeTheme();

    /**
     * @brief 应用彩色主题
     */
    void applyColorfulTheme();

    /**
     * @brief 应用单色主题
     */
    void applyMonochromaticTheme();

private:
    ChartTheme theme_;                          ///< 当前主题
    std::string primaryColor_;                  ///< 主色调
    std::vector<std::string> seriesColors_;     ///< 系列颜色列表
    
    TXFont titleFont_;                          ///< 标题字体
    TXFont axisLabelFont_;                      ///< 轴标签字体
    TXFont dataLabelFont_;                      ///< 数据标签字体
    
    u32 lineWidth_;                             ///< 线条宽度
    std::string borderColor_;                   ///< 边框颜色
};

/**
 * @brief 图表配置选项
 */
class TXChartConfig {
public:
    /**
     * @brief 默认构造函数
     */
    TXChartConfig();

    // ==================== 显示选项 ====================

    /**
     * @brief 设置是否显示图例
     * @param show 是否显示
     */
    void setShowLegend(bool show) { showLegend_ = show; }

    /**
     * @brief 获取是否显示图例
     */
    bool getShowLegend() const { return showLegend_; }

    /**
     * @brief 设置是否显示数据标签
     * @param show 是否显示
     */
    void setShowDataLabels(bool show) { showDataLabels_ = show; }

    /**
     * @brief 获取是否显示数据标签
     */
    bool getShowDataLabels() const { return showDataLabels_; }

    /**
     * @brief 设置是否显示网格线
     * @param show 是否显示
     */
    void setShowGridlines(bool show) { showGridlines_ = show; }

    /**
     * @brief 获取是否显示网格线
     */
    bool getShowGridlines() const { return showGridlines_; }

    // ==================== 标题配置 ====================

    /**
     * @brief 设置X轴标题
     * @param title 标题文本
     */
    void setXAxisTitle(const std::string& title) { xAxisTitle_ = title; }

    /**
     * @brief 获取X轴标题
     */
    const std::string& getXAxisTitle() const { return xAxisTitle_; }

    /**
     * @brief 设置Y轴标题
     * @param title 标题文本
     */
    void setYAxisTitle(const std::string& title) { yAxisTitle_ = title; }

    /**
     * @brief 获取Y轴标题
     */
    const std::string& getYAxisTitle() const { return yAxisTitle_; }

    // ==================== 样式配置 ====================

    /**
     * @brief 设置图表样式
     * @param style 样式配置
     */
    void setStyle(const TXChartStyleV2& style) { style_ = style; }

    /**
     * @brief 获取图表样式
     */
    const TXChartStyleV2& getStyle() const { return style_; }

    /**
     * @brief 获取图表样式（可修改）
     */
    TXChartStyleV2& getStyle() { return style_; }

private:
    bool showLegend_;           ///< 是否显示图例
    bool showDataLabels_;       ///< 是否显示数据标签
    bool showGridlines_;        ///< 是否显示网格线
    
    std::string xAxisTitle_;    ///< X轴标题
    std::string yAxisTitle_;    ///< Y轴标题

    TXChartStyleV2 style_;      ///< 图表样式
};

} // namespace TinaXlsx
