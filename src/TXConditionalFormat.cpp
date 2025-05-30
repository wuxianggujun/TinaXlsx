#include "TinaXlsx/TXConditionalFormat.hpp"
#include <algorithm>
#include <cmath>
#include <variant>

namespace TinaXlsx {

// ==================== TXConditionalFormatRule 基类实现 ====================

TXConditionalFormatRule::TXConditionalFormatRule(ConditionalFormatType type)
    : type_(type), priority_(1), stopIfTrue_(false) {}

TXConditionalFormatRule::~TXConditionalFormatRule() = default;

TXConditionalFormatRule::TXConditionalFormatRule(TXConditionalFormatRule&& other) noexcept 
    : type_(other.type_), priority_(other.priority_), stopIfTrue_(other.stopIfTrue_) {}

TXConditionalFormatRule& TXConditionalFormatRule::operator=(TXConditionalFormatRule&& other) noexcept {
    if (this != &other) {
        type_ = other.type_;
        priority_ = other.priority_;
        stopIfTrue_ = other.stopIfTrue_;
    }
    return *this;
}

ConditionalFormatType TXConditionalFormatRule::getType() const {
    return type_;
}

void TXConditionalFormatRule::setPriority(int priority) {
    priority_ = priority;
}

int TXConditionalFormatRule::getPriority() const {
    return priority_;
}

void TXConditionalFormatRule::setStopIfTrue(bool stopIfTrue) {
    stopIfTrue_ = stopIfTrue;
}

bool TXConditionalFormatRule::getStopIfTrue() const {
    return stopIfTrue_;
}

// ==================== TXCellValueRule 实现 ====================

TXCellValueRule::TXCellValueRule() 
    : TXConditionalFormatRule(ConditionalFormatType::CellValue)
    , operator_(ConditionalOperator::Equal) {
}

TXCellValueRule& TXCellValueRule::setCondition(ConditionalOperator op, const cell_value_t& value1, const cell_value_t& value2) {
    operator_ = op;
    value1_ = value1;
    value2_ = value2;
    return *this;
}

TXCellValueRule& TXCellValueRule::setFormat(const TXCellStyle& style) {
    format_ = style;
    return *this;
}

bool TXCellValueRule::evaluate(const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    // 简化实现 - 仅支持基本比较
    return true;
}

void TXCellValueRule::applyFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    style = format_;
}

// ==================== TXColorScaleRule 实现 ====================

TXColorScaleRule::TXColorScaleRule() 
    : TXConditionalFormatRule(ConditionalFormatType::ColorScale) {
}

TXColorScaleRule& TXColorScaleRule::setTwoColorScale(const ColorScalePoint& minPoint, const ColorScalePoint& maxPoint) {
    colorPoints_.clear();
    colorPoints_.push_back(minPoint);
    colorPoints_.push_back(maxPoint);
    return *this;
}

TXColorScaleRule& TXColorScaleRule::setThreeColorScale(const ColorScalePoint& minPoint, const ColorScalePoint& midPoint, const ColorScalePoint& maxPoint) {
    colorPoints_.clear();
    colorPoints_.push_back(minPoint);
    colorPoints_.push_back(midPoint);
    colorPoints_.push_back(maxPoint);
    return *this;
}

bool TXColorScaleRule::evaluate(const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    return true; // 色阶总是适用
}

void TXColorScaleRule::applyFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    if (!colorPoints_.empty()) {
        style.getFill().foregroundColor = colorPoints_[0].color;
    }
}

TXColor TXColorScaleRule::interpolateColor(double position) const {
    // 简化实现
    if (!colorPoints_.empty()) {
        return colorPoints_[0].color;
    }
    return ColorConstants::WHITE;
}

double TXColorScaleRule::calculatePosition(double value, const std::vector<std::vector<cell_value_t>>& context) const {
    // 简化实现
    return 0.5;
}

// ==================== TXDataBarRule 实现 ====================

TXDataBarRule::TXDataBarRule() 
    : TXConditionalFormatRule(ConditionalFormatType::DataBar) {
}

TXDataBarRule& TXDataBarRule::setSettings(const DataBarSettings& settings) {
    settings_ = settings;
    return *this;
}

bool TXDataBarRule::evaluate(const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    return true; // 数据条总是适用
}

void TXDataBarRule::applyFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    style.getFill().foregroundColor = settings_.fillColor;
}

// ==================== TXIconSetRule 实现 ====================

TXIconSetRule::TXIconSetRule() 
    : TXConditionalFormatRule(ConditionalFormatType::IconSet) {
}

TXIconSetRule& TXIconSetRule::setSettings(const IconSetSettings& settings) {
    settings_ = settings;
    return *this;
}

bool TXIconSetRule::evaluate(const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    return true; // 图标集总是适用
}

void TXIconSetRule::applyFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    // 简化实现 - 不实际应用图标
}

// ==================== TXConditionalFormatManager 实现 ====================

TXConditionalFormatManager::TXConditionalFormatManager() = default;

TXConditionalFormatManager::~TXConditionalFormatManager() = default;

void TXConditionalFormatManager::addRule(std::unique_ptr<TXConditionalFormatRule> rule) {
    if (rule) {
        rules_.push_back(std::move(rule));
    }
}

void TXConditionalFormatManager::removeRule(size_t index) {
    if (index < rules_.size()) {
        rules_.erase(rules_.begin() + index);
    }
}

size_t TXConditionalFormatManager::getRuleCount() const {
    return rules_.size();
}

void TXConditionalFormatManager::applyConditionalFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    for (const auto& rule : rules_) {
        if (rule->evaluate(value, context)) {
            rule->applyFormat(style, value, context);
            if (rule->getStopIfTrue()) {
                break;
            }
        }
    }
}

// ==================== 工厂方法实现 ====================

std::unique_ptr<TXCellValueRule> TXConditionalFormatManager::createCellValueRule(ConditionalOperator op, const cell_value_t& value1, const TXCellStyle& format, const cell_value_t& value2) {
    auto rule = std::make_unique<TXCellValueRule>();
    rule->setCondition(op, value1, value2);
    rule->setFormat(format);
    return rule;
}

std::unique_ptr<TXColorScaleRule> TXConditionalFormatManager::createTwoColorScale(const TXColor& minColor, const TXColor& maxColor) {
    auto rule = std::make_unique<TXColorScaleRule>();
    ColorScalePoint minPoint, maxPoint;
    minPoint.color = minColor;
    maxPoint.color = maxColor;
    rule->setTwoColorScale(minPoint, maxPoint);
    return rule;
}

std::unique_ptr<TXColorScaleRule> TXConditionalFormatManager::createThreeColorScale(const TXColor& minColor, const TXColor& midColor, const TXColor& maxColor) {
    auto rule = std::make_unique<TXColorScaleRule>();
    ColorScalePoint minPoint, midPoint, maxPoint;
    minPoint.color = minColor;
    midPoint.color = midColor;
    maxPoint.color = maxColor;
    rule->setThreeColorScale(minPoint, midPoint, maxPoint);
    return rule;
}

std::unique_ptr<TXDataBarRule> TXConditionalFormatManager::createDataBarRule(const TXColor& fillColor, bool showValue) {
    auto rule = std::make_unique<TXDataBarRule>();
    DataBarSettings settings;
    settings.fillColor = fillColor;
    settings.showValue = showValue;
    rule->setSettings(settings);
    return rule;
}

std::unique_ptr<TXIconSetRule> TXConditionalFormatManager::createIconSetRule(IconSetType iconType, bool showValue) {
    auto rule = std::make_unique<TXIconSetRule>();
    IconSetSettings settings;
    settings.iconType = iconType;
    settings.showValue = showValue;
    rule->setSettings(settings);
    return rule;
}

} // namespace TinaXlsx