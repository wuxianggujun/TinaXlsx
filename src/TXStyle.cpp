#include "TinaXlsx/TXStyle.hpp"
#include <tuple>

namespace TinaXlsx {

// ==================== TXFont 实现 ====================

bool TXFont::operator==(const TXFont& other) const {
    return name == other.name &&
           size == other.size &&
           color == other.color &&
           style == other.style &&
           bold == other.bold &&
           italic == other.italic &&
           underline == other.underline &&
           strikethrough == other.strikethrough;
}

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

TXCellStyle::~TXCellStyle() = default;

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

TXFont& TXCellStyle::getFont() {
    return font_;
}

const TXFont& TXCellStyle::getFont() const {
    return font_;
}

TXAlignment& TXCellStyle::getAlignment() {
    return alignment_;
}

const TXAlignment& TXCellStyle::getAlignment() const {
    return alignment_;
}

TXBorder& TXCellStyle::getBorder() {
    return border_;
}

const TXBorder& TXCellStyle::getBorder() const {
    return border_;
}

TXFill& TXCellStyle::getFill() {
    return fill_;
}

const TXFill& TXCellStyle::getFill() const {
    return fill_;
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

TXCellStyle& TXCellStyle::setFont(const std::string& name, font_size_t size) {
    font_.setName(name).setSize(size);
    return *this;
}

TXCellStyle& TXCellStyle::setFontColor(const TXColor& color) {
    font_.setColor(color);
    return *this;
}

TXCellStyle& TXCellStyle::setFontColor(color_value_t color) {
    font_.setColor(TXColor(color));
    return *this;
}

TXCellStyle& TXCellStyle::setFontStyle(FontStyle style) {
    font_.setStyle(style);
    return *this;
}

TXCellStyle& TXCellStyle::setHorizontalAlignment(HorizontalAlignment alignment) {
    alignment_.setHorizontal(alignment);
    return *this;
}

TXCellStyle& TXCellStyle::setVerticalAlignment(VerticalAlignment alignment) {
    alignment_.setVertical(alignment);
    return *this;
}

TXCellStyle& TXCellStyle::setBackgroundColor(const TXColor& color) {
    fill_.setSolidFill(color);
    return *this;
}

TXCellStyle& TXCellStyle::setBackgroundColor(color_value_t color) {
    fill_.setSolidFill(TXColor(color));
    return *this;
}

TXCellStyle& TXCellStyle::setAllBorders(BorderStyle style, const TXColor& color) {
    border_.setAllBorders(style, color);
    return *this;
}

void TXCellStyle::reset() {
    font_ = TXFont();
    alignment_ = TXAlignment();
    border_ = TXBorder();
    fill_ = TXFill();
}

bool TXCellStyle::operator==(const TXCellStyle& other) const {
    return font_ == other.font_ &&
           alignment_ == other.alignment_ &&
           border_ == other.border_ &&
           fill_ == other.fill_;
}
} // namespace TinaXlsx 
