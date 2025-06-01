#pragma once

#include "TXTypes.hpp"
#include "TXRange.hpp"
#include <string>
#include <vector>

namespace TinaXlsx {

// 前向声明
class TXSheet;

/**
 * @brief 图表数据系列
 */
class TXChartSeries {
public:
    /**
     * @brief 构造函数
     * @param name 系列名称
     * @param dataRange 数据范围
     * @param color 系列颜色（可选）
     */
    TXChartSeries(const std::string& name, const TXRange& dataRange, const std::string& color = "");

    /**
     * @brief 构造函数（指定类别和数值范围）
     * @param name 系列名称
     * @param categoryRange 类别范围（X轴标签）
     * @param valueRange 数值范围（Y轴数据）
     * @param color 系列颜色（可选）
     */
    TXChartSeries(const std::string& name, const TXRange& categoryRange, const TXRange& valueRange, const std::string& color = "");

    // ==================== 基本属性 ====================

    /**
     * @brief 获取系列名称
     */
    const std::string& getName() const { return name_; }

    /**
     * @brief 设置系列名称
     * @param name 名称
     */
    void setName(const std::string& name) { name_ = name; }

    /**
     * @brief 获取系列颜色
     */
    const std::string& getColor() const { return color_; }

    /**
     * @brief 设置系列颜色
     * @param color 颜色（十六进制）
     */
    void setColor(const std::string& color) { color_ = color; }

    // ==================== 数据范围 ====================

    /**
     * @brief 获取数据范围
     */
    const TXRange& getDataRange() const { return dataRange_; }

    /**
     * @brief 设置数据范围
     * @param range 数据范围
     */
    void setDataRange(const TXRange& range) { dataRange_ = range; }

    /**
     * @brief 获取类别范围
     */
    const TXRange& getCategoryRange() const { return categoryRange_; }

    /**
     * @brief 设置类别范围
     * @param range 类别范围
     */
    void setCategoryRange(const TXRange& range) { categoryRange_ = range; }

    /**
     * @brief 获取数值范围
     */
    const TXRange& getValueRange() const { return valueRange_; }

    /**
     * @brief 设置数值范围
     * @param range 数值范围
     */
    void setValueRange(const TXRange& range) { valueRange_ = range; }

    /**
     * @brief 检查是否有独立的类别和数值范围
     */
    bool hasSeparateRanges() const { return hasSeparateRanges_; }

    // ==================== 工作表关联 ====================

    /**
     * @brief 获取数据工作表
     */
    const TXSheet* getDataSheet() const { return dataSheet_; }

    /**
     * @brief 设置数据工作表
     * @param sheet 工作表指针
     */
    void setDataSheet(const TXSheet* sheet) { dataSheet_ = sheet; }

    // ==================== 显示选项 ====================

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
     * @brief 设置数据标签格式
     * @param format 格式字符串（如 "{value:F2}万元"）
     */
    void setDataLabelFormat(const std::string& format) { dataLabelFormat_ = format; }

    /**
     * @brief 获取数据标签格式
     */
    const std::string& getDataLabelFormat() const { return dataLabelFormat_; }

private:
    std::string name_;              ///< 系列名称
    std::string color_;             ///< 系列颜色
    
    TXRange dataRange_;             ///< 数据范围（包含类别和数值）
    TXRange categoryRange_;         ///< 类别范围（X轴标签）
    TXRange valueRange_;            ///< 数值范围（Y轴数据）
    bool hasSeparateRanges_;        ///< 是否有独立的类别和数值范围
    
    const TXSheet* dataSheet_;      ///< 数据工作表
    
    bool showDataLabels_;           ///< 是否显示数据标签
    std::string dataLabelFormat_;   ///< 数据标签格式
};

/**
 * @brief 多系列图表管理器
 */
class TXMultiSeriesChart {
public:
    /**
     * @brief 构造函数
     * @param title 图表标题
     */
    explicit TXMultiSeriesChart(const std::string& title = "");

    // ==================== 系列管理 ====================

    /**
     * @brief 添加数据系列
     * @param series 数据系列
     * @return 系列索引
     */
    u32 addSeries(const TXChartSeries& series);

    /**
     * @brief 添加数据系列（便捷方法）
     * @param name 系列名称
     * @param dataRange 数据范围
     * @param color 系列颜色（可选）
     * @return 系列索引
     */
    u32 addSeries(const std::string& name, const TXRange& dataRange, const std::string& color = "");

    /**
     * @brief 添加数据系列（指定类别和数值范围）
     * @param name 系列名称
     * @param categoryRange 类别范围
     * @param valueRange 数值范围
     * @param color 系列颜色（可选）
     * @return 系列索引
     */
    u32 addSeries(const std::string& name, const TXRange& categoryRange, const TXRange& valueRange, const std::string& color = "");

    /**
     * @brief 获取系列数量
     */
    u32 getSeriesCount() const { return static_cast<u32>(series_.size()); }

    /**
     * @brief 获取指定索引的系列
     * @param index 系列索引
     * @return 系列引用
     */
    const TXChartSeries& getSeries(u32 index) const;

    /**
     * @brief 获取指定索引的系列（可修改）
     * @param index 系列索引
     * @return 系列引用
     */
    TXChartSeries& getSeries(u32 index);

    /**
     * @brief 获取所有系列
     */
    const std::vector<TXChartSeries>& getAllSeries() const { return series_; }

    /**
     * @brief 移除指定索引的系列
     * @param index 系列索引
     */
    void removeSeries(u32 index);

    /**
     * @brief 清空所有系列
     */
    void clearSeries() { series_.clear(); }

    // ==================== 图表属性 ====================

    /**
     * @brief 获取图表标题
     */
    const std::string& getTitle() const { return title_; }

    /**
     * @brief 设置图表标题
     * @param title 标题
     */
    void setTitle(const std::string& title) { title_ = title; }

    /**
     * @brief 设置数据工作表（应用到所有系列）
     * @param sheet 工作表指针
     */
    void setDataSheet(const TXSheet* sheet);

private:
    std::string title_;                     ///< 图表标题
    std::vector<TXChartSeries> series_;     ///< 数据系列列表
};

} // namespace TinaXlsx
