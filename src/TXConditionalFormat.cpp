#include "TinaXlsx/TXConditionalFormat.hpp"
#include <algorithm>
#include <cmath>
#include <variant>

namespace TinaXlsx {

// ==================== TXConditionalFormatRule 基类实现 ====================

class TXConditionalFormatRule::Impl {
public:
    ConditionalFormatType type_;
    int priority_;
    bool stopIfTrue_;
    
    Impl(ConditionalFormatType type) : type_(type), priority_(0), stopIfTrue_(false) {}
};

TXConditionalFormatRule::TXConditionalFormatRule(ConditionalFormatType type) 
    : pImpl(std::make_unique<Impl>(type)) {}

TXConditionalFormatRule::~TXConditionalFormatRule() = default;

TXConditionalFormatRule::TXConditionalFormatRule(TXConditionalFormatRule&& other) noexcept 
    : pImpl(std::move(other.pImpl)) {}

TXConditionalFormatRule& TXConditionalFormatRule::operator=(TXConditionalFormatRule&& other) noexcept {
    if (this != &other) {
        pImpl = std::move(other.pImpl);
    }
    return *this;
}

ConditionalFormatType TXConditionalFormatRule::getType() const {
    return pImpl->type_;
}

void TXConditionalFormatRule::setPriority(int priority) {
    pImpl->priority_ = priority;
}

int TXConditionalFormatRule::getPriority() const {
    return pImpl->priority_;
}

void TXConditionalFormatRule::setStopIfTrue(bool stopIfTrue) {
    pImpl->stopIfTrue_ = stopIfTrue;
}

bool TXConditionalFormatRule::getStopIfTrue() const {
    return pImpl->stopIfTrue_;
}

// ==================== TXCellValueRule 实现 ====================

TXCellValueRule::TXCellValueRule() 
    : TXConditionalFormatRule(ConditionalFormatType::CellValue)
    , operator_(ConditionalOperator::Equal) {}

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
    auto compareValues = [](const cell_value_t& a, const cell_value_t& b) -> int {
        // 如果两个值类型相同，直接比较
        if (a.index() == b.index()) {
            if (std::holds_alternative<double>(a)) {
                double va = std::get<double>(a);
                double vb = std::get<double>(b);
                if (va < vb) return -1;
                if (va > vb) return 1;
                return 0;
            } else if (std::holds_alternative<std::string>(a)) {
                return std::get<std::string>(a).compare(std::get<std::string>(b));
            } else if (std::holds_alternative<int64_t>(a)) {
                int64_t va = std::get<int64_t>(a);
                int64_t vb = std::get<int64_t>(b);
                if (va < vb) return -1;
                if (va > vb) return 1;
                return 0;
            } else if (std::holds_alternative<bool>(a)) {
                bool va = std::get<bool>(a);
                bool vb = std::get<bool>(b);
                if (va == vb) return 0;
                return va ? 1 : -1;
            }
        }
        
        // 类型不同，转换为字符串比较
        auto toString = [](const cell_value_t& v) -> std::string {
            if (std::holds_alternative<std::string>(v)) {
                return std::get<std::string>(v);
            } else if (std::holds_alternative<double>(v)) {
                return std::to_string(std::get<double>(v));
            } else if (std::holds_alternative<int64_t>(v)) {
                return std::to_string(std::get<int64_t>(v));
            } else if (std::holds_alternative<bool>(v)) {
                return std::get<bool>(v) ? "TRUE" : "FALSE";
            }
            return "";
        };
        
        return toString(a).compare(toString(b));
    };
    
    switch (operator_) {
        case ConditionalOperator::Equal:
            return compareValues(value, value1_) == 0;
        case ConditionalOperator::NotEqual:
            return compareValues(value, value1_) != 0;
        case ConditionalOperator::Greater:
            return compareValues(value, value1_) > 0;
        case ConditionalOperator::GreaterEqual:
            return compareValues(value, value1_) >= 0;
        case ConditionalOperator::Less:
            return compareValues(value, value1_) < 0;
        case ConditionalOperator::LessEqual:
            return compareValues(value, value1_) <= 0;
        case ConditionalOperator::Between:
            return compareValues(value, value1_) >= 0 && compareValues(value, value2_) <= 0;
        case ConditionalOperator::NotBetween:
            return !(compareValues(value, value1_) >= 0 && compareValues(value, value2_) <= 0);
        case ConditionalOperator::Contains:
            if (std::holds_alternative<std::string>(value) && std::holds_alternative<std::string>(value1_)) {
                return std::get<std::string>(value).find(std::get<std::string>(value1_)) != std::string::npos;
            }
            return false;
        case ConditionalOperator::NotContains:
            if (std::holds_alternative<std::string>(value) && std::holds_alternative<std::string>(value1_)) {
                return std::get<std::string>(value).find(std::get<std::string>(value1_)) == std::string::npos;
            }
            return true;
        case ConditionalOperator::BeginsWith:
            if (std::holds_alternative<std::string>(value) && std::holds_alternative<std::string>(value1_)) {
                const std::string& str = std::get<std::string>(value);
                const std::string& prefix = std::get<std::string>(value1_);
                return str.size() >= prefix.size() && str.substr(0, prefix.size()) == prefix;
            }
            return false;
        case ConditionalOperator::EndsWith:
            if (std::holds_alternative<std::string>(value) && std::holds_alternative<std::string>(value1_)) {
                const std::string& str = std::get<std::string>(value);
                const std::string& suffix = std::get<std::string>(value1_);
                return str.size() >= suffix.size() && str.substr(str.size() - suffix.size()) == suffix;
            }
            return false;
    }
    return false;
}

void TXCellValueRule::applyFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    style = format_;
}

// ==================== TXColorScaleRule 实现 ====================

TXColorScaleRule::TXColorScaleRule() 
    : TXConditionalFormatRule(ConditionalFormatType::ColorScale) {}

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
    return true; // 色阶对所有数值都适用
}

void TXColorScaleRule::applyFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    if (colorPoints_.size() < 2) return;
    
    // 计算当前值在数据范围中的位置
    double position = calculatePosition(
        std::holds_alternative<double>(value) ? std::get<double>(value) : 
        (std::holds_alternative<int64_t>(value) ? static_cast<double>(std::get<int64_t>(value)) : 0.0),
        context
    );
    
    // 插值计算颜色
    TXColor interpolatedColor = interpolateColor(position);
    style.setBackgroundColor(interpolatedColor);
}

TXColor TXColorScaleRule::interpolateColor(double position) const {
    if (colorPoints_.empty()) return ColorConstants::WHITE;
    if (colorPoints_.size() == 1) return colorPoints_[0].color;
    
    // 确保position在0-1范围内
    position = std::max(0.0, std::min(1.0, position));
    
    if (colorPoints_.size() == 2) {
        // 两色插值
        const auto& color1 = colorPoints_[0].color;
        const auto& color2 = colorPoints_[1].color;
        
        color_value_t rgb1 = color1.getValue();
        color_value_t rgb2 = color2.getValue();
        
        uint8_t r1 = (rgb1 >> 16) & 0xFF;
        uint8_t g1 = (rgb1 >> 8) & 0xFF;
        uint8_t b1 = rgb1 & 0xFF;
        
        uint8_t r2 = (rgb2 >> 16) & 0xFF;
        uint8_t g2 = (rgb2 >> 8) & 0xFF;
        uint8_t b2 = rgb2 & 0xFF;
        
        uint8_t r = static_cast<uint8_t>(r1 + (r2 - r1) * position);
        uint8_t g = static_cast<uint8_t>(g1 + (g2 - g1) * position);
        uint8_t b = static_cast<uint8_t>(b1 + (b2 - b1) * position);
        
        return TXColor((static_cast<color_value_t>(r) << 16) | (static_cast<color_value_t>(g) << 8) | b);
    } else {
        // 三色插值
        if (position <= 0.5) {
            // 在前半段，插值第一和第二种颜色
            position *= 2.0; // 映射到0-1
            const auto& color1 = colorPoints_[0].color;
            const auto& color2 = colorPoints_[1].color;
            
            color_value_t rgb1 = color1.getValue();
            color_value_t rgb2 = color2.getValue();
            
            uint8_t r1 = (rgb1 >> 16) & 0xFF;
            uint8_t g1 = (rgb1 >> 8) & 0xFF;
            uint8_t b1 = rgb1 & 0xFF;
            
            uint8_t r2 = (rgb2 >> 16) & 0xFF;
            uint8_t g2 = (rgb2 >> 8) & 0xFF;
            uint8_t b2 = rgb2 & 0xFF;
            
            uint8_t r = static_cast<uint8_t>(r1 + (r2 - r1) * position);
            uint8_t g = static_cast<uint8_t>(g1 + (g2 - g1) * position);
            uint8_t b = static_cast<uint8_t>(b1 + (b2 - b1) * position);
            
            return TXColor((static_cast<color_value_t>(r) << 16) | (static_cast<color_value_t>(g) << 8) | b);
        } else {
            // 在后半段，插值第二和第三种颜色
            position = (position - 0.5) * 2.0; // 映射到0-1
            const auto& color2 = colorPoints_[1].color;
            const auto& color3 = colorPoints_[2].color;
            
            color_value_t rgb2 = color2.getValue();
            color_value_t rgb3 = color3.getValue();
            
            uint8_t r2 = (rgb2 >> 16) & 0xFF;
            uint8_t g2 = (rgb2 >> 8) & 0xFF;
            uint8_t b2 = rgb2 & 0xFF;
            
            uint8_t r3 = (rgb3 >> 16) & 0xFF;
            uint8_t g3 = (rgb3 >> 8) & 0xFF;
            uint8_t b3 = rgb3 & 0xFF;
            
            uint8_t r = static_cast<uint8_t>(r2 + (r3 - r2) * position);
            uint8_t g = static_cast<uint8_t>(g2 + (g3 - g2) * position);
            uint8_t b = static_cast<uint8_t>(b2 + (b3 - b2) * position);
            
            return TXColor((static_cast<color_value_t>(r) << 16) | (static_cast<color_value_t>(g) << 8) | b);
        }
    }
}

double TXColorScaleRule::calculatePosition(double value, const std::vector<std::vector<cell_value_t>>& context) const {
    // 找出所有数值的最小值和最大值
    double minVal = std::numeric_limits<double>::max();
    double maxVal = std::numeric_limits<double>::min();
    
    for (const auto& row : context) {
        for (const auto& cell : row) {
            double cellValue = 0.0;
            if (std::holds_alternative<double>(cell)) {
                cellValue = std::get<double>(cell);
            } else if (std::holds_alternative<int64_t>(cell)) {
                cellValue = static_cast<double>(std::get<int64_t>(cell));
            } else {
                continue; // 跳过非数值类型
            }
            
            minVal = std::min(minVal, cellValue);
            maxVal = std::max(maxVal, cellValue);
        }
    }
    
    if (maxVal == minVal) return 0.5; // 所有值相等
    
    return (value - minVal) / (maxVal - minVal);
}

// ==================== TXDataBarRule 实现 ====================

TXDataBarRule::TXDataBarRule() 
    : TXConditionalFormatRule(ConditionalFormatType::DataBar) {}

TXDataBarRule& TXDataBarRule::setSettings(const DataBarSettings& settings) {
    settings_ = settings;
    return *this;
}

bool TXDataBarRule::evaluate(const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    return std::holds_alternative<double>(value) || std::holds_alternative<int64_t>(value);
}

void TXDataBarRule::applyFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    // 设置数据条的背景颜色（简化实现）
    style.setBackgroundColor(settings_.fillColor);
    style.getBorder().setAllBorders(BorderStyle::Thin, settings_.borderColor);
}

// ==================== TXIconSetRule 实现 ====================

TXIconSetRule::TXIconSetRule() 
    : TXConditionalFormatRule(ConditionalFormatType::IconSet) {}

TXIconSetRule& TXIconSetRule::setSettings(const IconSetSettings& settings) {
    settings_ = settings;
    return *this;
}

bool TXIconSetRule::evaluate(const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    return std::holds_alternative<double>(value) || std::holds_alternative<int64_t>(value);
}

void TXIconSetRule::applyFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    // 图标集的实现比较复杂，这里提供一个简化版本
    // 根据图标索引设置不同的字体颜色来表示图标
    int iconIndex = determineIconIndex(
        std::holds_alternative<double>(value) ? std::get<double>(value) : 
        (std::holds_alternative<int64_t>(value) ? static_cast<double>(std::get<int64_t>(value)) : 0.0),
        context
    );
    
    // 根据图标索引设置颜色
    switch (iconIndex) {
        case 0: style.setFontColor(ColorConstants::RED); break;    // 低值
        case 1: style.setFontColor(ColorConstants::YELLOW); break; // 中值
        case 2: style.setFontColor(ColorConstants::GREEN); break;  // 高值
        default: style.setFontColor(ColorConstants::BLACK); break;
    }
}

int TXIconSetRule::determineIconIndex(double value, const std::vector<std::vector<cell_value_t>>& context) const {
    // 计算百分位数
    std::vector<double> values;
    for (const auto& row : context) {
        for (const auto& cell : row) {
            if (std::holds_alternative<double>(cell)) {
                values.push_back(std::get<double>(cell));
            } else if (std::holds_alternative<int64_t>(cell)) {
                values.push_back(static_cast<double>(std::get<int64_t>(cell)));
            }
        }
    }
    
    if (values.empty()) return 0;
    
    std::sort(values.begin(), values.end());
    
    // 计算当前值的百分位数
    auto it = std::lower_bound(values.begin(), values.end(), value);
    double percentile = (static_cast<double>(it - values.begin()) / values.size()) * 100.0;
    
    // 根据阈值确定图标索引
    for (size_t i = 0; i < settings_.thresholds.size(); ++i) {
        if (percentile <= settings_.thresholds[i]) {
            return settings_.reverseOrder ? static_cast<int>(settings_.thresholds.size() - i) : static_cast<int>(i);
        }
    }
    
    return settings_.reverseOrder ? 0 : static_cast<int>(settings_.thresholds.size());
}

// ==================== TXConditionalFormatManager 实现 ====================

class TXConditionalFormatManager::Impl {
public:
    std::vector<std::unique_ptr<TXConditionalFormatRule>> rules_;
};

TXConditionalFormatManager::TXConditionalFormatManager() 
    : pImpl(std::make_unique<Impl>()) {}

TXConditionalFormatManager::~TXConditionalFormatManager() = default;

void TXConditionalFormatManager::addRule(std::unique_ptr<TXConditionalFormatRule> rule) {
    if (rule) {
        pImpl->rules_.push_back(std::move(rule));
        
        // 按优先级排序
        std::sort(pImpl->rules_.begin(), pImpl->rules_.end(),
            [](const std::unique_ptr<TXConditionalFormatRule>& a, const std::unique_ptr<TXConditionalFormatRule>& b) {
                return a->getPriority() < b->getPriority();
            });
    }
}

void TXConditionalFormatManager::removeRule(size_t index) {
    if (index < pImpl->rules_.size()) {
        pImpl->rules_.erase(pImpl->rules_.begin() + index);
    }
}

void TXConditionalFormatManager::clearRules() {
    pImpl->rules_.clear();
}

size_t TXConditionalFormatManager::getRuleCount() const {
    return pImpl->rules_.size();
}

void TXConditionalFormatManager::applyConditionalFormat(TXCellStyle& style, const cell_value_t& value, const std::vector<std::vector<cell_value_t>>& context) const {
    for (const auto& rule : pImpl->rules_) {
        if (rule->evaluate(value, context)) {
            rule->applyFormat(style, value, context);
            if (rule->getStopIfTrue()) {
                break;
            }
        }
    }
}

// ==================== 静态创建方法实现 ====================

std::unique_ptr<TXCellValueRule> TXConditionalFormatManager::createCellValueRule(ConditionalOperator op, const cell_value_t& value1, const TXCellStyle& format, const cell_value_t& value2) {
    auto rule = std::make_unique<TXCellValueRule>();
    rule->setCondition(op, value1, value2);
    rule->setFormat(format);
    return rule;
}

std::unique_ptr<TXColorScaleRule> TXConditionalFormatManager::createTwoColorScale(const TXColor& minColor, const TXColor& maxColor) {
    auto rule = std::make_unique<TXColorScaleRule>();
    rule->setTwoColorScale(ColorScalePoint(0.0, minColor), ColorScalePoint(1.0, maxColor));
    return rule;
}

std::unique_ptr<TXColorScaleRule> TXConditionalFormatManager::createThreeColorScale(const TXColor& minColor, const TXColor& midColor, const TXColor& maxColor) {
    auto rule = std::make_unique<TXColorScaleRule>();
    rule->setThreeColorScale(ColorScalePoint(0.0, minColor), ColorScalePoint(0.5, midColor), ColorScalePoint(1.0, maxColor));
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