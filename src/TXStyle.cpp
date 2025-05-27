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

class TXCellStyle::Impl {
public:
    TXFont font_;
    TXAlignment alignment_;
    TXBorder border_;
    TXFill fill_;
    
    Impl() {}
    
    Impl(const Impl& other) 
        : font_(other.font_)
        , alignment_(other.alignment_)
        , border_(other.border_)
        , fill_(other.fill_)
    {}
    
    bool operator==(const Impl& other) const {
        return font_ == other.font_ &&
               alignment_ == other.alignment_ &&
               border_ == other.border_ &&
               fill_ == other.fill_;
    }
};

TXCellStyle::TXCellStyle() : pImpl(std::make_unique<Impl>()) {}

TXCellStyle::~TXCellStyle() = default;

TXCellStyle::TXCellStyle(const TXCellStyle& other) 
    : pImpl(std::make_unique<Impl>(*other.pImpl)) {}

TXCellStyle& TXCellStyle::operator=(const TXCellStyle& other) {
    if (this != &other) {
        pImpl = std::make_unique<Impl>(*other.pImpl);
    }
    return *this;
}

TXCellStyle::TXCellStyle(TXCellStyle&& other) noexcept : pImpl(std::move(other.pImpl)) {}

TXCellStyle& TXCellStyle::operator=(TXCellStyle&& other) noexcept {
    if (this != &other) {
        pImpl = std::move(other.pImpl);
    }
    return *this;
}

TXFont& TXCellStyle::getFont() {
    return pImpl->font_;
}

const TXFont& TXCellStyle::getFont() const {
    return pImpl->font_;
}

TXAlignment& TXCellStyle::getAlignment() {
    return pImpl->alignment_;
}

const TXAlignment& TXCellStyle::getAlignment() const {
    return pImpl->alignment_;
}

TXBorder& TXCellStyle::getBorder() {
    return pImpl->border_;
}

const TXBorder& TXCellStyle::getBorder() const {
    return pImpl->border_;
}

TXFill& TXCellStyle::getFill() {
    return pImpl->fill_;
}

const TXFill& TXCellStyle::getFill() const {
    return pImpl->fill_;
}

TXCellStyle& TXCellStyle::setFont(const TXFont& font) {
    pImpl->font_ = font;
    return *this;
}

TXCellStyle& TXCellStyle::setAlignment(const TXAlignment& alignment) {
    pImpl->alignment_ = alignment;
    return *this;
}

TXCellStyle& TXCellStyle::setBorder(const TXBorder& border) {
    pImpl->border_ = border;
    return *this;
}

TXCellStyle& TXCellStyle::setFill(const TXFill& fill) {
    pImpl->fill_ = fill;
    return *this;
}

TXCellStyle& TXCellStyle::setFont(const std::string& name, font_size_t size) {
    pImpl->font_.setName(name).setSize(size);
    return *this;
}

TXCellStyle& TXCellStyle::setFontColor(const TXColor& color) {
    pImpl->font_.setColor(color);
    return *this;
}

TXCellStyle& TXCellStyle::setFontColor(color_value_t color) {
    pImpl->font_.setColor(TXColor(color));
    return *this;
}

TXCellStyle& TXCellStyle::setFontStyle(FontStyle style) {
    pImpl->font_.setStyle(style);
    return *this;
}

TXCellStyle& TXCellStyle::setHorizontalAlignment(HorizontalAlignment alignment) {
    pImpl->alignment_.setHorizontal(alignment);
    return *this;
}

TXCellStyle& TXCellStyle::setVerticalAlignment(VerticalAlignment alignment) {
    pImpl->alignment_.setVertical(alignment);
    return *this;
}

TXCellStyle& TXCellStyle::setBackgroundColor(const TXColor& color) {
    pImpl->fill_.setSolidFill(color);
    return *this;
}

TXCellStyle& TXCellStyle::setBackgroundColor(color_value_t color) {
    pImpl->fill_.setSolidFill(TXColor(color));
    return *this;
}

TXCellStyle& TXCellStyle::setAllBorders(BorderStyle style, const TXColor& color) {
    pImpl->border_.setAllBorders(style, color);
    return *this;
}

void TXCellStyle::reset() {
    pImpl = std::make_unique<Impl>();
}

bool TXCellStyle::operator==(const TXCellStyle& other) const {
    return *pImpl == *other.pImpl;
}

// ==================== 预定义样式实现 ====================

namespace Styles {

TXCellStyle createHeaderStyle() {
    TXCellStyle style;
    style.setFont("Calibri", 14)
         .setFontColor(ColorConstants::BLACK)
         .setFontStyle(FontStyle::Bold)
         .setHorizontalAlignment(HorizontalAlignment::Center)
         .setVerticalAlignment(VerticalAlignment::Middle)
         .setBackgroundColor(ColorConstants::LIGHT_GRAY)
         .setAllBorders(BorderStyle::Thin, ColorConstants::BLACK);
    return style;
}

TXCellStyle createDataStyle() {
    TXCellStyle style;
    style.setFont("Calibri", 11)
         .setFontColor(ColorConstants::BLACK)
         .setHorizontalAlignment(HorizontalAlignment::Left)
         .setVerticalAlignment(VerticalAlignment::Middle);
    return style;
}

TXCellStyle createNumberStyle() {
    TXCellStyle style;
    style.setFont("Calibri", 11)
         .setFontColor(ColorConstants::BLACK)
         .setHorizontalAlignment(HorizontalAlignment::Right)
         .setVerticalAlignment(VerticalAlignment::Middle);
    return style;
}

TXCellStyle createHighlightStyle(const TXColor& backgroundColor) {
    TXCellStyle style;
    style.setFont("Calibri", 11)
         .setFontColor(ColorConstants::BLACK)
         .setFontStyle(FontStyle::Bold)
         .setHorizontalAlignment(HorizontalAlignment::Left)
         .setVerticalAlignment(VerticalAlignment::Middle)
         .setBackgroundColor(backgroundColor);
    return style;
}

TXCellStyle createTableStyle() {
    TXCellStyle style;
    style.setFont("Calibri", 11)
         .setFontColor(ColorConstants::BLACK)
         .setHorizontalAlignment(HorizontalAlignment::Left)
         .setVerticalAlignment(VerticalAlignment::Middle)
         .setAllBorders(BorderStyle::Thin, ColorConstants::GRAY);
    return style;
}

} // namespace Styles

} // namespace TinaXlsx 