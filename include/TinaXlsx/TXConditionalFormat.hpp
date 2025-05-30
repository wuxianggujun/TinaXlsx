#pragma once

#include "TXTypes.hpp"
#include "TXColor.hpp"
#include "TXStyle.hpp"
#include <string>
#include <vector>
#include <memory>
#include <functional>

#include "TXRange.hpp"

namespace TinaXlsx {

/**
 * @brief 条件格式类型枚举
 */
enum class ConditionalFormatType : uint8_t {
    CellValue = 0,          ///< 单元格值条件
    Expression = 1,         ///< 表达式条件
    ColorScale = 2,         ///< 色阶
    DataBar = 3,            ///< 数据条
    IconSet = 4,            ///< 图标集
    TopPercentile = 5,      ///< 前百分比
    BottomPercentile = 6,   ///< 后百分比
    AboveAverage = 7,       ///< 高于平均值
    BelowAverage = 8,       ///< 低于平均值
    UniqueValues = 9,       ///< 唯一值
    DuplicateValues = 10    ///< 重复值
};

/**
 * @brief 条件格式操作符枚举
 */
enum class ConditionalOperator : uint8_t {
    Equal = 0,              ///< 等于
    NotEqual = 1,           ///< 不等于
    Greater = 2,            ///< 大于
    GreaterEqual = 3,       ///< 大于等于
    Less = 4,               ///< 小于
    LessEqual = 5,          ///< 小于等于
    Between = 6,            ///< 介于
    NotBetween = 7,         ///< 不介于
    Contains = 8,           ///< 包含
    NotContains = 9,        ///< 不包含
    BeginsWith = 10,        ///< 开始于
    EndsWith = 11           ///< 结束于
};

/**
 * @brief 图标集类型枚举
 */
enum class IconSetType : uint8_t {
    ThreeArrows = 0,        ///< 三个箭头
    ThreeArrowsGray = 1,    ///< 三个灰色箭头
    ThreeFlags = 2,         ///< 三个标志
    ThreeTrafficLights = 3, ///< 三个交通灯
    ThreeTrafficLightsWithRim = 4, ///< 三个有边框的交通灯
    FourArrows = 5,         ///< 四个箭头
    FourArrowsGray = 6,     ///< 四个灰色箭头
    FourRedToBlack = 7,     ///< 四个红到黑
    FourRating = 8,         ///< 四个评级
    FourTrafficLights = 9,  ///< 四个交通灯
    FiveArrows = 10,        ///< 五个箭头
    FiveArrowsGray = 11,    ///< 五个灰色箭头
    FiveQuarters = 12,      ///< 五个四分之一
    FiveRating = 13,        ///< 五个评级
    FiveBoxes = 14          ///< 五个方框
};

/**
 * @brief 色阶颜色点结构
 */
struct ColorScalePoint {
    double value;           ///< 值
    TXColor color;          ///< 颜色
    bool isPercentile;      ///< 是否为百分位数
    bool isFormula;         ///< 是否为公式
    std::string formula;    ///< 公式字符串
    
    ColorScalePoint() : value(0.0), color(ColorConstants::WHITE), isPercentile(false), isFormula(false) {}
    ColorScalePoint(double val, const TXColor& col) : value(val), color(col), isPercentile(false), isFormula(false) {}
};

/**
 * @brief 数据条设置结构
 */
struct DataBarSettings {
    TXColor fillColor;      ///< 填充颜色
    TXColor borderColor;    ///< 边框颜色
    bool showValue;         ///< 显示数值
    bool gradient;          ///< 渐变效果
    double minValue;        ///< 最小值
    double maxValue;        ///< 最大值
    bool autoMin;           ///< 自动最小值
    bool autoMax;           ///< 自动最大值
    
    DataBarSettings() 
        : fillColor(ColorConstants::BLUE)
        , borderColor(ColorConstants::DARK_BLUE)
        , showValue(true)
        , gradient(true)
        , minValue(0.0)
        , maxValue(100.0)
        , autoMin(true)
        , autoMax(true) {}
};

/**
 * @brief 图标集设置结构
 */
struct IconSetSettings {
    IconSetType iconType;   ///< 图标类型
    bool reverseOrder;      ///< 反向排序
    bool showValue;         ///< 显示数值
    std::vector<double> thresholds; ///< 阈值
    
    IconSetSettings() : iconType(IconSetType::ThreeArrows), reverseOrder(false), showValue(true) {
        // 默认三等分阈值
        thresholds = {33.33, 66.67};
    }
};

/**
 * @brief 条件格式规则基类
 */
class TXConditionalFormatRule {
public:
    TXConditionalFormatRule(ConditionalFormatType type);
    virtual ~TXConditionalFormatRule();
    
    // 禁用拷贝，使用智能指针管理
    TXConditionalFormatRule(const TXConditionalFormatRule&) = delete;
    TXConditionalFormatRule& operator=(const TXConditionalFormatRule&) = delete;
    
    // 支持移动
    TXConditionalFormatRule(TXConditionalFormatRule&& other) noexcept;
    TXConditionalFormatRule& operator=(TXConditionalFormatRule&& other) noexcept;
    
    /**
     * @brief 获取条件格式类型
     * @return 条件格式类型
     */
    ConditionalFormatType getType() const;
    
    /**
     * @brief 设置优先级
     * @param priority 优先级（数字越小优先级越高）
     */
    void setPriority(int priority);
    
    /**
     * @brief 获取优先级
     * @return 优先级
     */
    int getPriority() const;
    
    /**
     * @brief 设置停止if true
     * @param stopIfTrue 如果为真则停止执行后续规则
     */
    void setStopIfTrue(bool stopIfTrue);
    
    /**
     * @brief 获取停止if true
     * @return 停止if true状态
     */
    bool getStopIfTrue() const;
    
    /**
     * @brief 评估条件是否满足
     * @param value 单元格值
     * @param context 上下文信息（包含所有单元格数据）
     * @return 条件满足返回true
     */
    virtual bool evaluate(const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const = 0;
    
    /**
     * @brief 应用格式
     * @param style 要修改的样式
     * @param value 单元格值
     * @param context 上下文信息
     */
    virtual void applyFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const = 0;

protected:
    ConditionalFormatType type_;
    int priority_;
    bool stopIfTrue_;
};

/**
 * @brief 单元格值条件格式规则
 */
class TXCellValueRule : public TXConditionalFormatRule {
public:
    TXCellValueRule();
    
    /**
     * @brief 设置条件
     * @param op 操作符
     * @param value1 第一个值
     * @param value2 第二个值（用于Between和NotBetween）
     * @return 自身引用
     */
    TXCellValueRule& setCondition(ConditionalOperator op, const cell_value_t& value1, const cell_value_t& value2 = cell_value_t{});
    
    /**
     * @brief 设置格式样式
     * @param style 样式
     * @return 自身引用
     */
    TXCellValueRule& setFormat(const TXCellStyle& style);
    
    bool evaluate(const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const override;
    void applyFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const override;

private:
    ConditionalOperator operator_;
    cell_value_t value1_;
    cell_value_t value2_;
    TXCellStyle format_;
};

/**
 * @brief 色阶条件格式规则
 */
class TXColorScaleRule : public TXConditionalFormatRule {
public:
    TXColorScaleRule();
    
    /**
     * @brief 设置为二色阶
     * @param minPoint 最小值点
     * @param maxPoint 最大值点
     * @return 自身引用
     */
    TXColorScaleRule& setTwoColorScale(const ColorScalePoint& minPoint, const ColorScalePoint& maxPoint);
    
    /**
     * @brief 设置为三色阶
     * @param minPoint 最小值点
     * @param midPoint 中间值点
     * @param maxPoint 最大值点
     * @return 自身引用
     */
    TXColorScaleRule& setThreeColorScale(const ColorScalePoint& minPoint, const ColorScalePoint& midPoint, const ColorScalePoint& maxPoint);
    
    bool evaluate(const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const override;
    void applyFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const override;

private:
    std::vector<ColorScalePoint> colorPoints_;
    TXColor interpolateColor(double position) const;
    double calculatePosition(double value, const std::vector<std::vector<cell_value_t>>& context) const;
};

/**
 * @brief 数据条条件格式规则
 */
class TXDataBarRule : public TXConditionalFormatRule {
public:
    TXDataBarRule();
    
    /**
     * @brief 设置数据条设置
     * @param settings 数据条设置
     * @return 自身引用
     */
    TXDataBarRule& setSettings(const DataBarSettings& settings);
    
    bool evaluate(const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const override;
    void applyFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const override;

private:
    DataBarSettings settings_;
};

/**
 * @brief 图标集条件格式规则
 */
class TXIconSetRule : public TXConditionalFormatRule {
public:
    TXIconSetRule();
    
    /**
     * @brief 设置图标集设置
     * @param settings 图标集设置
     * @return 自身引用
     */
    TXIconSetRule& setSettings(const IconSetSettings& settings);
    
    bool evaluate(const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const override;
    void applyFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const override;

private:
    IconSetSettings settings_;
    int determineIconIndex(double value, const std::vector<std::vector<cell_value_t>>& context) const;
};

/**
 * @brief 条件格式管理器
 */
class TXConditionalFormatManager {
public:
    TXConditionalFormatManager();
    ~TXConditionalFormatManager();
    
    /**
     * @brief 添加条件格式规则
     * @param rule 条件格式规则
     */
    void addRule(std::unique_ptr<TXConditionalFormatRule> rule);
    
    /**
     * @brief 移除条件格式规则
     * @param index 规则索引
     */
    void removeRule(size_t index);
    
    /**
     * @brief 清除所有规则
     */
    void clearRules();
    
    /**
     * @brief 获取规则数量
     * @return 规则数量
     */
    size_t getRuleCount() const;
    
    /**
     * @brief 应用条件格式
     * @param style 要修改的样式
     * @param value 单元格值
     * @param context 上下文信息
     */
    void applyConditionalFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const;
    
    // ==================== 快捷创建方法 ====================
    
    /**
     * @brief 创建单元格值规则
     * @param op 操作符
     * @param value1 第一个值
     * @param format 格式样式
     * @param value2 第二个值（可选）
     * @return 规则指针
     */
    static std::unique_ptr<TXCellValueRule> createCellValueRule(ConditionalOperator op, const cell_value_t& value1, const TXCellStyle& format, const cell_value_t& value2 = cell_value_t{});
    
    /**
     * @brief 创建二色阶规则
     * @param minColor 最小值颜色
     * @param maxColor 最大值颜色
     * @return 规则指针
     */
    static std::unique_ptr<TXColorScaleRule> createTwoColorScale(const TXColor& minColor, const TXColor& maxColor);
    
    /**
     * @brief 创建三色阶规则
     * @param minColor 最小值颜色
     * @param midColor 中间值颜色
     * @param maxColor 最大值颜色
     * @return 规则指针
     */
    static std::unique_ptr<TXColorScaleRule> createThreeColorScale(const TXColor& minColor, const TXColor& midColor, const TXColor& maxColor);
    
    /**
     * @brief 创建数据条规则
     * @param fillColor 填充颜色
     * @param showValue 是否显示数值
     * @return 规则指针
     */
    static std::unique_ptr<TXDataBarRule> createDataBarRule(const TXColor& fillColor = ColorConstants::BLUE, bool showValue = true);
    
    /**
     * @brief 创建图标集规则
     * @param iconType 图标类型
     * @param showValue 是否显示数值
     * @return 规则指针
     */
    static std::unique_ptr<TXIconSetRule> createIconSetRule(IconSetType iconType = IconSetType::ThreeArrows, bool showValue = true);

private:
    std::vector<std::unique_ptr<TXConditionalFormatRule>> rules_;
    TXRange range_;
};

} // namespace TinaXlsx
