/**
 * @file Format.cpp
 * @brief Excel格式处理实现
 */

#include "TinaXlsx/Format.hpp"
#include <xlsxwriter.h>
#include <unordered_map>

namespace TinaXlsx {

struct Format::Impl {
    lxw_format* format = nullptr;
    lxw_workbook* workbook = nullptr;
    
    Impl(lxw_workbook* wb) : workbook(wb) {
        format = workbook_add_format(workbook);
    }
    
    ~Impl() {
        // format由workbook管理，不需要手动释放
    }
};

Format::Format(lxw_workbook* workbook) 
    : pImpl_(std::make_unique<Impl>(workbook)) {
}

Format::Format(Format&& other) noexcept 
    : pImpl_(std::move(other.pImpl_)) {
}

Format& Format::operator=(Format&& other) noexcept {
    if (this != &other) {
        pImpl_ = std::move(other.pImpl_);
    }
    return *this;
}

Format::~Format() = default;

Format& Format::setFontName(const std::string& fontName) {
    if (pImpl_->format) {
        format_set_font_name(pImpl_->format, fontName.c_str());
    }
    return *this;
}

Format& Format::setFontSize(double fontSize) {
    if (pImpl_->format) {
        format_set_font_size(pImpl_->format, fontSize);
    }
    return *this;
}

Format& Format::setFontColor(Color color) {
    if (pImpl_->format) {
        format_set_font_color(pImpl_->format, color);
    }
    return *this;
}

Format& Format::setBold(bool bold) {
    if (pImpl_->format) {
        format_set_bold(pImpl_->format);
    }
    return *this;
}

Format& Format::setItalic(bool italic) {
    if (pImpl_->format) {
        format_set_italic(pImpl_->format);
    }
    return *this;
}

Format& Format::setUnderline(bool underline) {
    if (pImpl_->format) {
        format_set_underline(pImpl_->format, LXW_UNDERLINE_SINGLE);
    }
    return *this;
}

Format& Format::setStrikeout(bool strikeout) {
    if (pImpl_->format) {
        format_set_font_strikeout(pImpl_->format);
    }
    return *this;
}

Format& Format::setBackgroundColor(Color color) {
    if (pImpl_->format) {
        format_set_bg_color(pImpl_->format, color);
    }
    return *this;
}

Format& Format::setForegroundColor(Color color) {
    if (pImpl_->format) {
        format_set_fg_color(pImpl_->format, color);
    }
    return *this;
}

Format& Format::setAlignment(Alignment alignment) {
    if (pImpl_->format) {
        format_set_align(pImpl_->format, static_cast<uint8_t>(alignment));
    }
    return *this;
}

Format& Format::setVerticalAlignment(VerticalAlignment alignment) {
    if (pImpl_->format) {
        format_set_align(pImpl_->format, static_cast<uint8_t>(alignment));
    }
    return *this;
}

Format& Format::setTextWrap(bool wrap) {
    if (pImpl_->format) {
        format_set_text_wrap(pImpl_->format);
    }
    return *this;
}

Format& Format::setIndent(int indent) {
    if (pImpl_->format) {
        format_set_indent(pImpl_->format, indent);
    }
    return *this;
}

Format& Format::setRotation(int angle) {
    if (pImpl_->format) {
        format_set_rotation(pImpl_->format, angle);
    }
    return *this;
}

Format& Format::setBorder(BorderStyle style) {
    if (pImpl_->format) {
        format_set_border(pImpl_->format, static_cast<uint8_t>(style));
    }
    return *this;
}

Format& Format::setLeftBorder(BorderStyle style) {
    if (pImpl_->format) {
        format_set_left(pImpl_->format, static_cast<uint8_t>(style));
    }
    return *this;
}

Format& Format::setRightBorder(BorderStyle style) {
    if (pImpl_->format) {
        format_set_right(pImpl_->format, static_cast<uint8_t>(style));
    }
    return *this;
}

Format& Format::setTopBorder(BorderStyle style) {
    if (pImpl_->format) {
        format_set_top(pImpl_->format, static_cast<uint8_t>(style));
    }
    return *this;
}

Format& Format::setBottomBorder(BorderStyle style) {
    if (pImpl_->format) {
        format_set_bottom(pImpl_->format, static_cast<uint8_t>(style));
    }
    return *this;
}

Format& Format::setBorderColor(Color color) {
    if (pImpl_->format) {
        format_set_border_color(pImpl_->format, color);
    }
    return *this;
}

Format& Format::setNumberFormat(NumberFormat format) {
    if (!pImpl_->format) return *this;
    
    static const std::unordered_map<NumberFormat, const char*> formatMap = {
        {NumberFormat::General, "General"},
        {NumberFormat::Number, "0"},
        {NumberFormat::Number2, "0.00"},
        {NumberFormat::Percentage, "0%"},
        {NumberFormat::Date, "m/d/yy"},
        {NumberFormat::Time, "h:mm:ss AM/PM"},
        {NumberFormat::DateTime, "m/d/yy h:mm"},
        {NumberFormat::Currency, "$#,##0.00"},
        {NumberFormat::Accounting, "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)"},
        {NumberFormat::Scientific, "0.00E+00"},
        {NumberFormat::Text, "@"}
    };
    
    auto it = formatMap.find(format);
    if (it != formatMap.end()) {
        format_set_num_format(pImpl_->format, it->second);
    }
    
    return *this;
}

Format& Format::setCustomNumberFormat(const std::string& formatString) {
    if (pImpl_->format) {
        format_set_num_format(pImpl_->format, formatString.c_str());
    }
    return *this;
}

Format& Format::setLocked(bool locked) {
    if (pImpl_->format) {
        if (!locked) {
            // 在libxlsxwriter中，单元格默认是锁定的，只能解锁
            format_set_unlocked(pImpl_->format);
        }
        // 如果要锁定，由于默认就是锁定状态，所以不需要特殊操作
    }
    return *this;
}

Format& Format::setHidden(bool hidden) {
    if (pImpl_->format) {
        format_set_hidden(pImpl_->format);
    }
    return *this;
}

lxw_format* Format::getInternalFormat() const {
    return pImpl_->format;
}

std::unique_ptr<Format> Format::clone(lxw_workbook* workbook) const {
    auto newFormat = std::make_unique<Format>(workbook);
    // libxlsxwriter不支持直接克隆格式，需要手动复制所有属性
    // 这里为了简化，返回新的基础格式
    return newFormat;
}

// FormatBuilder 实现

std::unique_ptr<Format> FormatBuilder::createBasic() const {
    return std::make_unique<Format>(workbook_);
}

std::unique_ptr<Format> FormatBuilder::createTitle(double fontSize, bool bold) const {
    auto format = createBasic();
    format->setFontSize(fontSize);
    if (bold) {
        format->setBold();
    }
    format->setAlignment(Alignment::Center);
    return format;
}

std::unique_ptr<Format> FormatBuilder::createHeader(Color backgroundColor, Color fontColor) const {
    auto format = createBasic();
    format->setBackgroundColor(backgroundColor)
          .setFontColor(fontColor)
          .setBold()
          .setAlignment(Alignment::Center)
          .setBorder(BorderStyle::Thin);
    return format;
}

std::unique_ptr<Format> FormatBuilder::createNumber(int decimalPlaces) const {
    auto format = createBasic();
    std::string formatStr = "0";
    if (decimalPlaces > 0) {
        formatStr += "." + std::string(decimalPlaces, '0');
    }
    format->setCustomNumberFormat(formatStr);
    return format;
}

std::unique_ptr<Format> FormatBuilder::createPercentage(int decimalPlaces) const {
    auto format = createBasic();
    std::string formatStr = "0";
    if (decimalPlaces > 0) {
        formatStr += "." + std::string(decimalPlaces, '0');
    }
    formatStr += "%";
    format->setCustomNumberFormat(formatStr);
    return format;
}

std::unique_ptr<Format> FormatBuilder::createDate() const {
    auto format = createBasic();
    format->setNumberFormat(Format::NumberFormat::Date);
    return format;
}

std::unique_ptr<Format> FormatBuilder::createCurrency(const std::string& currency) const {
    auto format = createBasic();
    std::string formatStr = currency + "#,##0.00";
    format->setCustomNumberFormat(formatStr);
    return format;
}

std::unique_ptr<Format> FormatBuilder::createBorder(BorderStyle style, Color color) const {
    auto format = createBasic();
    format->setBorder(style).setBorderColor(color);
    return format;
}

} // namespace TinaXlsx 