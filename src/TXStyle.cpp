#include "TinaXlsx/TXStyle.hpp"
#include <tuple>

namespace TinaXlsx {

// ==================== TXFont 实现 ====================

TXFont& TXFont::setBold(bool bold) {
    if (bold) {
        style = static_cast<FontStyle>(static_cast<uint8_t>(style) | static_cast<uint8_t>(FontStyle::Bold));
    } else {
        style = static_cast<FontStyle>(static_cast<uint8_t>(style) & ~static_cast<uint8_t>(FontStyle::Bold));
    }
    return *this;
}

TXFont& TXFont::setItalic(bool italic) {
    if (italic) {
        style = static_cast<FontStyle>(static_cast<uint8_t>(style) | static_cast<uint8_t>(FontStyle::Italic));
    } else {
        style = static_cast<FontStyle>(static_cast<uint8_t>(style) & ~static_cast<uint8_t>(FontStyle::Italic));
    }
    return *this;
}

TXFont& TXFont::setUnderline(bool underline) {
    if (underline) {
        style = static_cast<FontStyle>(static_cast<uint8_t>(style) | static_cast<uint8_t>(FontStyle::Underline));
    } else {
        style = static_cast<FontStyle>(static_cast<uint8_t>(style) & ~static_cast<uint8_t>(FontStyle::Underline));
    }
    return *this;
}

TXFont& TXFont::setStrikethrough(bool strikethrough) {
    if (strikethrough) {
        style = static_cast<FontStyle>(static_cast<uint8_t>(style) | static_cast<uint8_t>(FontStyle::Strikethrough));
    } else {
        style = static_cast<FontStyle>(static_cast<uint8_t>(style) & ~static_cast<uint8_t>(FontStyle::Strikethrough));
    }
    return *this;
}

bool TXFont::isBold() const {
    return (static_cast<uint8_t>(style) & static_cast<uint8_t>(FontStyle::Bold)) != 0;
}

bool TXFont::isItalic() const {
    return (static_cast<uint8_t>(style) & static_cast<uint8_t>(FontStyle::Italic)) != 0;
}

bool TXFont::hasUnderline() const {
    return (static_cast<uint8_t>(style) & static_cast<uint8_t>(FontStyle::Underline)) != 0;
}

bool TXFont::hasStrikethrough() const {
    return (static_cast<uint8_t>(style) & static_cast<uint8_t>(FontStyle::Strikethrough)) != 0;
}

bool TXFont::operator==(const TXFont& other) const {
    return name == other.name &&
           size == other.size &&
           color == other.color &&
           style == other.style;
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

TXBorder& TXBorder::setAllBorders(BorderStyle style, TXTypes::ColorValue color) {
    leftStyle = rightStyle = topStyle = bottomStyle = style;
    leftColor = rightColor = topColor = bottomColor = color;
    return *this;
}

TXBorder& TXBorder::setLeftBorder(BorderStyle style, TXTypes::ColorValue color) {
    leftStyle = style;
    leftColor = color;
    return *this;
}

TXBorder& TXBorder::setRightBorder(BorderStyle style, TXTypes::ColorValue color) {
    rightStyle = style;
    rightColor = color;
    return *this;
}

TXBorder& TXBorder::setTopBorder(BorderStyle style, TXTypes::ColorValue color) {
    topStyle = style;
    topColor = color;
    return *this;
}

TXBorder& TXBorder::setBottomBorder(BorderStyle style, TXTypes::ColorValue color) {
    bottomStyle = style;
    bottomColor = color;
    return *this;
}

TXBorder& TXBorder::setDiagonalBorder(BorderStyle style, TXTypes::ColorValue color, bool up, bool down) {
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

TXFill& TXFill::setSolidFill(TXTypes::ColorValue color) {
    pattern = FillPattern::Solid;
    foregroundColor = color;
    backgroundColor = Colors::WHITE;
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
    TXTypes::StyleId styleId_;
    
    Impl() : styleId_(TXTypes::INVALID_STYLE_ID) {}
    
    Impl(const Impl& other) 
        : font_(other.font_)
        , alignment_(other.alignment_)
        , border_(other.border_)
        , fill_(other.fill_)
        , styleId_(other.styleId_)
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

TXCellStyle& TXCellStyle::setFont(const std::string& name, TXTypes::FontSize size) {
    pImpl->font_.setName(name).setSize(size);
    return *this;
}

TXCellStyle& TXCellStyle::setFontColor(TXTypes::ColorValue color) {
    pImpl->font_.setColor(color);
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

TXCellStyle& TXCellStyle::setBackgroundColor(TXTypes::ColorValue color) {
    pImpl->fill_.setSolidFill(color);
    return *this;
}

TXCellStyle& TXCellStyle::setAllBorders(BorderStyle style, TXTypes::ColorValue color) {
    pImpl->border_.setAllBorders(style, color);
    return *this;
}

TXTypes::StyleId TXCellStyle::getStyleId() const {
    return pImpl->styleId_;
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
         .setFontColor(Colors::BLACK)
         .setFontStyle(FontStyle::Bold)
         .setHorizontalAlignment(HorizontalAlignment::Center)
         .setVerticalAlignment(VerticalAlignment::Middle)
         .setBackgroundColor(Colors::LIGHT_GRAY)
         .setAllBorders(BorderStyle::Thin, Colors::BLACK);
    return style;
}

TXCellStyle createDataStyle() {
    TXCellStyle style;
    style.setFont("Calibri", 11)
         .setFontColor(Colors::BLACK)
         .setHorizontalAlignment(HorizontalAlignment::Left)
         .setVerticalAlignment(VerticalAlignment::Middle);
    return style;
}

TXCellStyle createNumberStyle() {
    TXCellStyle style;
    style.setFont("Calibri", 11)
         .setFontColor(Colors::BLACK)
         .setHorizontalAlignment(HorizontalAlignment::Right)
         .setVerticalAlignment(VerticalAlignment::Middle);
    return style;
}

TXCellStyle createHighlightStyle(TXTypes::ColorValue backgroundColor) {
    TXCellStyle style;
    style.setFont("Calibri", 11)
         .setFontColor(Colors::BLACK)
         .setFontStyle(FontStyle::Bold)
         .setHorizontalAlignment(HorizontalAlignment::Left)
         .setVerticalAlignment(VerticalAlignment::Middle)
         .setBackgroundColor(backgroundColor);
    return style;
}

TXCellStyle createTableStyle() {
    TXCellStyle style;
    style.setFont("Calibri", 11)
         .setFontColor(Colors::BLACK)
         .setHorizontalAlignment(HorizontalAlignment::Left)
         .setVerticalAlignment(VerticalAlignment::Middle)
         .setAllBorders(BorderStyle::Thin, Colors::GRAY);
    return style;
}

} // namespace Styles

} // namespace TinaXlsx 