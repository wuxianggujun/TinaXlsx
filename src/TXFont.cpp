#include "TinaXlsx/TXFont.hpp"
#include "TinaXlsx/TXError.hpp"
#include <sstream>
#include <regex>
#include <unordered_set>

namespace TinaXlsx {

// ==================== 常量定义 ====================

constexpr font_size_t MIN_FONT_SIZE = 1;
constexpr font_size_t MAX_FONT_SIZE = 409;
constexpr uint8_t DEFAULT_CHARSET = 1;
constexpr uint8_t DEFAULT_FAMILY = 2;  // Swiss family

static const std::unordered_set<std::string> VALID_SCHEMES = {
    "major", "minor", "none", ""
};

static const std::unordered_set<uint8_t> VALID_FAMILIES = {
    1, 2, 3, 4, 5  // Roman, Swiss, Modern, Script, Decorative
};

// ==================== 构造函数实现 ====================

TXFont::TXFont()
    : name_("Calibri")
    , size_(DEFAULT_FONT_SIZE)
    , color_(ColorConstants::BLACK)
    , bold_(false)
    , italic_(false)
    , underline_style_(UnderlineStyle::None)
    , strikethrough_(false)
    , charset_(DEFAULT_CHARSET)
    , family_(DEFAULT_FAMILY)
    , scheme_("none")
{
}

TXFont::TXFont(std::string name, font_size_t size)
    : name_(std::move(name))
    , size_(size)
    , color_(ColorConstants::BLACK)
    , bold_(false)
    , italic_(false)
    , underline_style_(UnderlineStyle::None)
    , strikethrough_(false)
    , charset_(DEFAULT_CHARSET)
    , family_(DEFAULT_FAMILY)
    , scheme_("none")
{
    // 验证参数有效性
    if (auto result = TXFont::validateName(name_); result.isError()) {
        name_ = "Calibri";  // 回退到默认值
    }
    if (auto result = TXFont::validateSize(size_); result.isError()) {
        size_ = DEFAULT_FONT_SIZE;  // 回退到默认值
    }
}

TXFont::TXFont(std::string name, font_size_t size, const TXColor& color, 
               bool bold, bool italic)
    : name_(std::move(name))
    , size_(size)
    , color_(color)
    , bold_(bold)
    , italic_(italic)
    , underline_style_(UnderlineStyle::None)
    , strikethrough_(false)
    , charset_(DEFAULT_CHARSET)
    , family_(DEFAULT_FAMILY)
    , scheme_("none")
{
    // 验证参数有效性
    if (auto result = TXFont::validateName(name_); result.isError()) {
        name_ = "Calibri";
    }
    if (auto result = TXFont::validateSize(size_); result.isError()) {
        size_ = DEFAULT_FONT_SIZE;
    }
}

// ==================== 设置方法实现 ====================

TXResult<void> TXFont::setName(const std::string& name) {
    auto result = validateName(name);
    if (result.isError()) {
        return result;
    }
    name_ = name;
    return Ok();
}

TXResult<void> TXFont::setSize(font_size_t size) {
    auto result = validateSize(size);
    if (result.isError()) {
        return result;
    }
    size_ = size;
    return Ok();
}

TXResult<void> TXFont::setColor(const TXColor& color) {
    color_ = color;
    return Ok();
}

TXResult<void> TXFont::setColor(color_value_t colorValue) {
    color_ = TXColor(colorValue);
    return Ok();
}

TXFont& TXFont::setBold(bool enable) noexcept {
    bold_ = enable;
    return *this;
}

TXFont& TXFont::setItalic(bool enable) noexcept {
    italic_ = enable;
    return *this;
}

TXFont& TXFont::setUnderline(UnderlineStyle style) noexcept {
    underline_style_ = style;
    return *this;
}

TXFont& TXFont::setStrikethrough(bool enable) noexcept {
    strikethrough_ = enable;
    return *this;
}

TXFont& TXFont::setStyle(FontStyle style) noexcept {
    // 重置所有样式
    bold_ = false;
    italic_ = false;
    setUnderline(UnderlineStyle::None);
    strikethrough_ = false;
    
    // 根据样式设置对应的属性
    if ((static_cast<uint8_t>(style) & static_cast<uint8_t>(FontStyle::Bold)) != 0) {
        bold_ = true;
    }
    if ((static_cast<uint8_t>(style) & static_cast<uint8_t>(FontStyle::Italic)) != 0) {
        italic_ = true;
    }
    if ((static_cast<uint8_t>(style) & static_cast<uint8_t>(FontStyle::Underline)) != 0) {
        setUnderline(UnderlineStyle::Single);
    }
    if ((static_cast<uint8_t>(style) & static_cast<uint8_t>(FontStyle::Strikethrough)) != 0) {
        strikethrough_ = true;
    }
    
    return *this;
}

TXResult<void> TXFont::setCharset(uint8_t charset) {
    auto result = validateCharset(charset);
    if (result.isError()) {
        return result;
    }
    charset_ = charset;
    return Ok();
}

TXResult<void> TXFont::setFamily(uint8_t family) {
    auto result = validateFamily(family);
    if (result.isError()) {
        return result;
    }
    family_ = family;
    return Ok();
}

TXResult<void> TXFont::setScheme(const std::string& scheme) {
    auto result = validateScheme(scheme);
    if (result.isError()) {
        return result;
    }
    scheme_ = scheme;
    return Ok();
}

// ==================== 查询方法实现 ====================

TXResult<void> TXFont::validate() const {
    if (auto result = validateName(name_); result.isError()) {
        return result;
    }
    if (auto result = validateSize(size_); result.isError()) {
        return result;
    }
    if (auto result = validateCharset(charset_); result.isError()) {
        return result;
    }
    if (auto result = validateFamily(family_); result.isError()) {
        return result;
    }
    if (auto result = validateScheme(scheme_); result.isError()) {
        return result;
    }
    return Ok();
}

bool TXFont::isDefault() const noexcept {
    return name_ == "Calibri" && 
           size_ == DEFAULT_FONT_SIZE && 
           color_ == ColorConstants::BLACK &&
           !bold_ && !italic_ && 
           underline_style_ == UnderlineStyle::None && 
           !strikethrough_ &&
           charset_ == DEFAULT_CHARSET &&
           family_ == DEFAULT_FAMILY;
}

std::string TXFont::getDisplayName() const {
    std::ostringstream oss;
    oss << name_ << " " << size_;
    
    if (bold_ || italic_) {
        oss << " (";
        if (bold_) oss << "粗体";
        if (bold_ && italic_) oss << ", ";
        if (italic_) oss << "斜体";
        oss << ")";
    }
    
    if (hasUnderline()) {
        oss << " [下划线]";
    }
    
    if (strikethrough_) {
        oss << " [删除线]";
    }
    
    return oss.str();
}

std::string TXFont::getUniqueKey() const {
    std::ostringstream oss;
    oss << name_ << "|" << size_ << "|" << color_.getValue() << "|"
        << (bold_ ? "1" : "0") << (italic_ ? "1" : "0") 
        << static_cast<int>(underline_style_) << (strikethrough_ ? "1" : "0")
        << "|" << static_cast<int>(charset_) << "|" << static_cast<int>(family_)
        << "|" << scheme_;
    return oss.str();
}

// ==================== 序列化方法实现 ====================

std::string TXFont::toString() const {
    std::ostringstream oss;
    oss << "Font{name=" << name_ 
        << ", size=" << size_
        << ", color=0x" << std::hex << color_.getValue() << std::dec
        << ", bold=" << (bold_ ? "true" : "false")
        << ", italic=" << (italic_ ? "true" : "false")
        << ", underline=" << static_cast<int>(underline_style_)
        << ", strikethrough=" << (strikethrough_ ? "true" : "false")
        << ", charset=" << static_cast<int>(charset_)
        << ", family=" << static_cast<int>(family_)
        << ", scheme=" << scheme_ << "}";
    return oss.str();
}

TXResult<TXFont> TXFont::fromString(const std::string& description) {
    // 简单的解析实现，可以根据需要扩展
    std::regex pattern(R"(Font\{name=([^,]+), size=(\d+), color=0x([0-9a-fA-F]+), bold=(true|false), italic=(true|false), underline=(\d+), strikethrough=(true|false), charset=(\d+), family=(\d+), scheme=([^}]*)\})");
    std::smatch matches;
    
    if (!std::regex_match(description, matches, pattern)) {
        return Err<TXFont>(TXErrorCode::InvalidArgument, "Invalid font description format");
    }
    
    try {
        TXFont font;
        
        auto nameResult = font.setName(matches[1].str());
        if (nameResult.isError()) return Err<TXFont>(nameResult.error());
        
        auto sizeResult = font.setSize(static_cast<font_size_t>(std::stoi(matches[2].str())));
        if (sizeResult.isError()) return Err<TXFont>(sizeResult.error());
        
        font.setColor(static_cast<color_value_t>(std::stoul(matches[3].str(), nullptr, 16)));
        font.setBold(matches[4].str() == "true");
        font.setItalic(matches[5].str() == "true");
        font.setUnderline(static_cast<UnderlineStyle>(std::stoi(matches[6].str())));
        font.setStrikethrough(matches[7].str() == "true");
        
        auto charsetResult = font.setCharset(static_cast<uint8_t>(std::stoi(matches[8].str())));
        if (charsetResult.isError()) return Err<TXFont>(charsetResult.error());
        
        auto familyResult = font.setFamily(static_cast<uint8_t>(std::stoi(matches[9].str())));
        if (familyResult.isError()) return Err<TXFont>(familyResult.error());
        
        auto schemeResult = font.setScheme(matches[10].str());
        if (schemeResult.isError()) return Err<TXFont>(schemeResult.error());
        
        return Ok(std::move(font));
        
    } catch (const std::exception&) {
        return Err<TXFont>(TXErrorCode::InvalidArgument, "Failed to parse font description values");
    }
}

// ==================== 比较操作符实现 ====================

bool TXFont::operator==(const TXFont& other) const noexcept {
    return name_ == other.name_ &&
           size_ == other.size_ &&
           color_ == other.color_ &&
           bold_ == other.bold_ &&
           italic_ == other.italic_ &&
           underline_style_ == other.underline_style_ &&
           strikethrough_ == other.strikethrough_ &&
           charset_ == other.charset_ &&
           family_ == other.family_ &&
           scheme_ == other.scheme_;
}

bool TXFont::operator<(const TXFont& other) const noexcept {
    // 用于std::map等容器的排序
    if (name_ != other.name_) return name_ < other.name_;
    if (size_ != other.size_) return size_ < other.size_;
    if (color_.getValue() != other.color_.getValue()) return color_.getValue() < other.color_.getValue();
    if (bold_ != other.bold_) return bold_ < other.bold_;
    if (italic_ != other.italic_) return italic_ < other.italic_;
    if (underline_style_ != other.underline_style_) return underline_style_ < other.underline_style_;
    if (strikethrough_ != other.strikethrough_) return strikethrough_ < other.strikethrough_;
    if (charset_ != other.charset_) return charset_ < other.charset_;
    if (family_ != other.family_) return family_ < other.family_;
    return scheme_ < other.scheme_;
}

// ==================== 静态工厂方法实现 ====================

TXFont TXFont::createDefault() {
    return TXFont();
}

TXFont TXFont::createHeading(font_size_t size) {
    TXFont font("Calibri", size);
    font.setBold(true);
    return font;
}

TXFont TXFont::createEmphasis(const TXFont& base_font) {
    TXFont font = base_font;
    font.setItalic(true);
    return font;
}

// ==================== 私有验证方法实现 ====================

TXResult<void> TXFont::validateName(const std::string& name) {
    if (name.empty()) {
        return Err<void>(TXErrorCode::InvalidArgument, "Font name cannot be empty");
    }
    if (name.length() > 31) {  // Excel限制
        return Err<void>(TXErrorCode::InvalidArgument, "Font name too long (max 31 characters)");
    }
    // 检查是否包含非法字符
    if (name.find('\0') != std::string::npos) {
        return Err<void>(TXErrorCode::InvalidArgument, "Font name contains null character");
    }
    return Ok();
}

TXResult<void> TXFont::validateSize(font_size_t size) {
    if (size < MIN_FONT_SIZE || size > MAX_FONT_SIZE) {
        return Err<void>(TXErrorCode::InvalidArgument, 
            "Font size must be between " + std::to_string(MIN_FONT_SIZE) + 
            " and " + std::to_string(MAX_FONT_SIZE) + " points");
    }
    return Ok();
}

TXResult<void> TXFont::validateCharset(uint8_t charset) {
    // 大部分字符集ID都是有效的，只检查一些明确无效的值
    // 详细的字符集验证可以根据需要添加
    return Ok();
}

TXResult<void> TXFont::validateFamily(uint8_t family) {
    if (VALID_FAMILIES.find(family) == VALID_FAMILIES.end()) {
        return Err<void>(TXErrorCode::InvalidArgument, 
            "Invalid font family ID: " + std::to_string(family));
    }
    return Ok();
}

TXResult<void> TXFont::validateScheme(const std::string& scheme) {
    if (VALID_SCHEMES.find(scheme) == VALID_SCHEMES.end()) {
        return Err<void>(TXErrorCode::InvalidArgument, 
            "Invalid font scheme: " + scheme);
    }
    return Ok();
}

} // namespace TinaXlsx 