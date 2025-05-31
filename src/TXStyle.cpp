#include "TinaXlsx/TXStyle.hpp"
#include <tuple>

namespace TinaXlsx {

// ==================== TXAlignment 实现 ====================

bool TXAlignment::operator==(const TXAlignment& other) const {
    return horizontal == other.horizontal &&
           vertical == other.vertical &&
           wrapText == other.wrapText &&
           shrinkToFit == other.shrinkToFit &&
           textRotation == other.textRotation &&
           indent == other.indent;
}

// ==================== TXBorder 实现 ====================

TXBorder& TXBorder::setAllBorders(BorderStyle style, const TXColor& color) {
    leftStyle = rightStyle = topStyle = bottomStyle = style;
    leftColor = rightColor = topColor = bottomColor = color;
    return *this;
}

TXBorder& TXBorder::setLeftBorder(BorderStyle style, const TXColor& color) {
    leftStyle = style;
    leftColor = color;
    return *this;
}

TXBorder& TXBorder::setRightBorder(BorderStyle style, const TXColor& color) {
    rightStyle = style;
    rightColor = color;
    return *this;
}

TXBorder& TXBorder::setTopBorder(BorderStyle style, const TXColor& color) {
    topStyle = style;
    topColor = color;
    return *this;
}

TXBorder& TXBorder::setBottomBorder(BorderStyle style, const TXColor& color) {
    bottomStyle = style;
    bottomColor = color;
    return *this;
}

TXBorder& TXBorder::setDiagonalBorder(BorderStyle style, const TXColor& color, bool up, bool down) {
    diagonalStyle = style;
    diagonalColor = color;
    diagonalUp = up;
    diagonalDown = down;
    return *this;
}

bool TXBorder::operator==(const TXBorder& other) const {
    return leftStyle == other.leftStyle &&
           rightStyle == other.rightStyle &&
           topStyle == other.topStyle &&
           bottomStyle == other.bottomStyle &&
           diagonalStyle == other.diagonalStyle &&
           leftColor == other.leftColor &&
           rightColor == other.rightColor &&
           topColor == other.topColor &&
           bottomColor == other.bottomColor &&
           diagonalColor == other.diagonalColor &&
           diagonalUp == other.diagonalUp &&
           diagonalDown == other.diagonalDown;
}

// ==================== TXFill 实现 ====================

TXFill& TXFill::setSolidFill(const TXColor& color) {
    pattern = FillPattern::Solid;
    foregroundColor = color;
    backgroundColor = ColorConstants::WHITE;
    return *this;
}

bool TXFill::operator==(const TXFill& other) const {
    return pattern == other.pattern &&
           foregroundColor == other.foregroundColor &&
           backgroundColor == other.backgroundColor;
}

// ==================== TXCellStyle 实现 ====================

TXCellStyle::TXCellStyle() = default;

TXCellStyle::TXCellStyle(const TXCellStyle& other) 
    : font_(other.font_)
    , alignment_(other.alignment_)
    , border_(other.border_)
    , fill_(other.fill_) {}

TXCellStyle& TXCellStyle::operator=(const TXCellStyle& other) {
    if (this != &other) {
        font_ = other.font_;
        alignment_ = other.alignment_;
        border_ = other.border_;
        fill_ = other.fill_;
    }
    return *this;
}

TXCellStyle::TXCellStyle(TXCellStyle&& other) noexcept 
    : font_(std::move(other.font_))
    , alignment_(std::move(other.alignment_))
    , border_(std::move(other.border_))
    , fill_(std::move(other.fill_)) {}

TXCellStyle& TXCellStyle::operator=(TXCellStyle&& other) noexcept {
    if (this != &other) {
        font_ = std::move(other.font_);
        alignment_ = std::move(other.alignment_);
        border_ = std::move(other.border_);
        fill_ = std::move(other.fill_);
    }
    return *this;
}

TXCellStyle& TXCellStyle::setFont(const TXFont& font) {
    font_ = font;
    return *this;
}

TXCellStyle& TXCellStyle::setAlignment(const TXAlignment& alignment) {
    alignment_ = alignment;
    return *this;
}

TXCellStyle& TXCellStyle::setBorder(const TXBorder& border) {
    border_ = border;
    return *this;
}

TXCellStyle& TXCellStyle::setFill(const TXFill& fill) {
    fill_ = fill;
    return *this;
}

// 字体便捷方法
TXCellStyle& TXCellStyle::setFontName(const std::string& name) {
    // 使用新的TXFont类的公共方法
    font_.setName(name);
    return *this;
}

TXCellStyle& TXCellStyle::setFontSize(font_size_t size) {
    font_.setSize(size);
    return *this;
}

TXCellStyle& TXCellStyle::setFontColor(const TXColor& color) {
    font_.setColor(color);
    return *this;
}

TXCellStyle& TXCellStyle::setFontColor(color_value_t color) {
    font_.setColor(color);
    return *this;
}

TXCellStyle& TXCellStyle::setFontBold(bool bold) {
    font_.setBold(bold);
    return *this;
}

TXCellStyle& TXCellStyle::setFontItalic(bool italic) {
    font_.setItalic(italic);
    return *this;
}

TXCellStyle& TXCellStyle::setFontStyle(FontStyle style) {
    font_.setStyle(style);
    return *this;
}

// 对齐便捷方法
TXCellStyle& TXCellStyle::setHorizontalAlignment(HorizontalAlignment alignment) {
    alignment_.horizontal = alignment;
    return *this;
}

TXCellStyle& TXCellStyle::setVerticalAlignment(VerticalAlignment alignment) {
    alignment_.vertical = alignment;
    return *this;
}

TXCellStyle& TXCellStyle::setWrapText(bool wrap) {
    alignment_.wrapText = wrap;
    return *this;
}

TXCellStyle& TXCellStyle::setTextRotation(uint32_t rotation) {
    alignment_.textRotation = rotation;
    return *this;
}

// 填充便捷方法
TXCellStyle& TXCellStyle::setBackgroundColor(color_value_t color) {
    fill_.setSolidFill(TXColor(color));
    return *this;
}

TXCellStyle& TXCellStyle::setBackgroundColor(const TXColor& color) {
    fill_.setSolidFill(color);
    return *this;
}

TXCellStyle& TXCellStyle::setFillPattern(FillPattern pattern) {
    fill_.pattern = pattern;
    return *this;
}

TXCellStyle& TXCellStyle::setSolidFill(const TXColor& color) {
    fill_.setSolidFill(color);
    return *this;
}

// 边框便捷方法
TXCellStyle& TXCellStyle::setAllBorders(BorderStyle style, const TXColor& color) {
    border_.setAllBorders(style, color);
    return *this;
}

TXCellStyle& TXCellStyle::setLeftBorder(BorderStyle style, const TXColor& color) {
    border_.setLeftBorder(style, color);
    return *this;
}

TXCellStyle& TXCellStyle::setRightBorder(BorderStyle style, const TXColor& color) {
    border_.setRightBorder(style, color);
    return *this;
}

TXCellStyle& TXCellStyle::setTopBorder(BorderStyle style, const TXColor& color) {
    border_.setTopBorder(style, color);
    return *this;
}

TXCellStyle& TXCellStyle::setBottomBorder(BorderStyle style, const TXColor& color) {
    border_.setBottomBorder(style, color);
    return *this;
}

// 比较操作符
bool TXCellStyle::operator==(const TXCellStyle& other) const {
    return font_ == other.font_ &&
           alignment_ == other.alignment_ &&
           border_ == other.border_ &&
           fill_ == other.fill_;
}

// 工具方法
void TXCellStyle::reset() {
    font_ = TXFont();
    alignment_ = TXAlignment();
    border_ = TXBorder();
    fill_ = TXFill();
}

bool TXCellStyle::isDefault() const {
    return font_.isDefault() &&
           alignment_.horizontal == HorizontalAlignment::Left &&
           alignment_.vertical == VerticalAlignment::Bottom &&
           !alignment_.wrapText &&
           !alignment_.shrinkToFit &&
           alignment_.textRotation == 0 &&
           alignment_.indent == 0 &&
           border_.leftStyle == BorderStyle::None &&
           fill_.pattern == FillPattern::None;
}

std::string TXCellStyle::getUniqueKey() const {
    return font_.getUniqueKey() + "|" +
           std::to_string(static_cast<int>(alignment_.horizontal)) + "|" +
           std::to_string(static_cast<int>(alignment_.vertical)) + "|" +
           std::to_string(alignment_.wrapText) + "|" +
           std::to_string(static_cast<int>(border_.leftStyle)) + "|" +
           std::to_string(static_cast<int>(fill_.pattern));
}

} // namespace TinaXlsx 
