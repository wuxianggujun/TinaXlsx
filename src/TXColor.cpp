#include "TinaXlsx/TXColor.hpp"
#include <algorithm>
#include <sstream>
#include <iomanip>
#include <cctype>
#include <cmath>

namespace TinaXlsx {

// ==================== TXColor 构造函数实现 ====================

TXColor::TXColor(uint8_t r, uint8_t g, uint8_t b, uint8_t a) {
    value_ = (static_cast<color_value_t>(a) << 24) |
             (static_cast<color_value_t>(r) << 16) |
             (static_cast<color_value_t>(g) << 8)  |
             static_cast<color_value_t>(b);
}

TXColor::TXColor(const std::string& hex) {
    *this = fromHex(hex);
}

// ==================== TXColor 获取器实现 ====================

uint8_t TXColor::getRed() const {
    return static_cast<uint8_t>((value_ >> 16) & 0xFF);
}

uint8_t TXColor::getGreen() const {
    return static_cast<uint8_t>((value_ >> 8) & 0xFF);
}

uint8_t TXColor::getBlue() const {
    return static_cast<uint8_t>(value_ & 0xFF);
}

uint8_t TXColor::getAlpha() const {
    return static_cast<uint8_t>((value_ >> 24) & 0xFF);
}

std::tuple<uint8_t, uint8_t, uint8_t, uint8_t> TXColor::getComponents() const {
    return {getRed(), getGreen(), getBlue(), getAlpha()};
}

// ==================== TXColor 设置器实现 ====================

TXColor& TXColor::setValue(color_value_t color) {
    value_ = color;
    return *this;
}

TXColor& TXColor::setRGB(uint8_t r, uint8_t g, uint8_t b) {
    value_ = (value_ & 0xFF000000) | // 保持alpha不变
             (static_cast<color_value_t>(r) << 16) |
             (static_cast<color_value_t>(g) << 8)  |
             static_cast<color_value_t>(b);
    return *this;
}

TXColor& TXColor::setARGB(uint8_t r, uint8_t g, uint8_t b, uint8_t a) {
    value_ = (static_cast<color_value_t>(a) << 24) |
             (static_cast<color_value_t>(r) << 16) |
             (static_cast<color_value_t>(g) << 8)  |
             static_cast<color_value_t>(b);
    return *this;
}

TXColor& TXColor::setRed(uint8_t r) {
    value_ = (value_ & 0xFF00FFFF) | (static_cast<color_value_t>(r) << 16);
    return *this;
}

TXColor& TXColor::setGreen(uint8_t g) {
    value_ = (value_ & 0xFFFF00FF) | (static_cast<color_value_t>(g) << 8);
    return *this;
}

TXColor& TXColor::setBlue(uint8_t b) {
    value_ = (value_ & 0xFFFFFF00) | static_cast<color_value_t>(b);
    return *this;
}

TXColor& TXColor::setAlpha(uint8_t a) {
    value_ = (value_ & 0x00FFFFFF) | (static_cast<color_value_t>(a) << 24);
    return *this;
}

// ==================== TXColor 转换方法实现 ====================

std::string TXColor::toHex(bool with_alpha, bool with_prefix) const {
    std::ostringstream oss;
    if (with_prefix) {
        oss << "#";
    }
    
    if (with_alpha) {
        oss << std::hex << std::uppercase << std::setfill('0') << std::setw(8) << value_;
    } else {
        oss << std::hex << std::uppercase << std::setfill('0') << std::setw(6) << (value_ & 0x00FFFFFF);
    }
    
    return oss.str();
}

std::string TXColor::toRgbString() const {
    std::ostringstream oss;
    oss << "rgb(" << static_cast<int>(getRed()) << ", " 
        << static_cast<int>(getGreen()) << ", " 
        << static_cast<int>(getBlue()) << ")";
    return oss.str();
}

std::string TXColor::toRgbaString() const {
    std::ostringstream oss;
    oss << "rgba(" << static_cast<int>(getRed()) << ", " 
        << static_cast<int>(getGreen()) << ", " 
        << static_cast<int>(getBlue()) << ", " 
        << std::fixed << std::setprecision(3) << (getAlpha() / 255.0) << ")";
    return oss.str();
}

// ==================== TXColor 颜色操作实现 ====================

TXColor TXColor::adjustBrightness(double factor) const {
    factor = std::max(0.0, std::min(2.0, factor)); // 限制范围
    
    uint8_t r = clamp(getRed() * factor);
    uint8_t g = clamp(getGreen() * factor);
    uint8_t b = clamp(getBlue() * factor);
    
    return TXColor(r, g, b, getAlpha());
}

TXColor TXColor::withAlpha(uint8_t alpha) const {
    return TXColor(getRed(), getGreen(), getBlue(), alpha);
}

TXColor TXColor::blend(const TXColor& other, double ratio) const {
    ratio = std::max(0.0, std::min(1.0, ratio)); // 限制范围
    
    uint8_t r = clamp(getRed() * (1.0 - ratio) + other.getRed() * ratio);
    uint8_t g = clamp(getGreen() * (1.0 - ratio) + other.getGreen() * ratio);
    uint8_t b = clamp(getBlue() * (1.0 - ratio) + other.getBlue() * ratio);
    uint8_t a = clamp(getAlpha() * (1.0 - ratio) + other.getAlpha() * ratio);
    
    return TXColor(r, g, b, a);
}

TXColor TXColor::getComplementary() const {
    return TXColor(255 - getRed(), 255 - getGreen(), 255 - getBlue(), getAlpha());
}

bool TXColor::isDark() const {
    // 使用亮度计算公式判断是否为暗色
    double luminance = 0.299 * getRed() + 0.587 * getGreen() + 0.114 * getBlue();
    return luminance < 128.0;
}

std::string TXColor::toARGBHexString() const
{
    std::ostringstream oss;
    oss << std::hex << std::setfill('0')
       << std::setw(2) << static_cast<int>(getAlpha())   // AA
       << std::setw(2) << static_cast<int>(getRed())     // RR
       << std::setw(2) << static_cast<int>(getGreen())   // GG
       << std::setw(2) << static_cast<int>(getBlue());    // BB
    return oss.str();
}

// ==================== TXColor 静态工厂方法实现 ====================

TXColor TXColor::fromHSV(int h, int s, int v, uint8_t a) {
    h = h % 360; // 确保色相在0-359范围内
    s = std::max(0, std::min(100, s));
    v = std::max(0, std::min(100, v));
    
    double c = (v / 100.0) * (s / 100.0);
    double x = c * (1 - std::abs(std::fmod(h / 60.0, 2) - 1));
    double m = (v / 100.0) - c;
    
    double r_prime, g_prime, b_prime;
    
    if (h >= 0 && h < 60) {
        r_prime = c; g_prime = x; b_prime = 0;
    } else if (h >= 60 && h < 120) {
        r_prime = x; g_prime = c; b_prime = 0;
    } else if (h >= 120 && h < 180) {
        r_prime = 0; g_prime = c; b_prime = x;
    } else if (h >= 180 && h < 240) {
        r_prime = 0; g_prime = x; b_prime = c;
    } else if (h >= 240 && h < 300) {
        r_prime = x; g_prime = 0; b_prime = c;
    } else {
        r_prime = c; g_prime = 0; b_prime = x;
    }
    
    uint8_t r = clamp((r_prime + m) * 255);
    uint8_t g = clamp((g_prime + m) * 255);
    uint8_t b = clamp((b_prime + m) * 255);
    
    return TXColor(r, g, b, a);
}

TXColor TXColor::fromHSL(int h, int s, int l, uint8_t a) {
    h = h % 360; // 确保色相在0-359范围内
    s = std::max(0, std::min(100, s));
    l = std::max(0, std::min(100, l));
    
    double c = (1 - std::abs(2 * (l / 100.0) - 1)) * (s / 100.0);
    double x = c * (1 - std::abs(std::fmod(h / 60.0, 2) - 1));
    double m = (l / 100.0) - c / 2;
    
    double r_prime, g_prime, b_prime;
    
    if (h >= 0 && h < 60) {
        r_prime = c; g_prime = x; b_prime = 0;
    } else if (h >= 60 && h < 120) {
        r_prime = x; g_prime = c; b_prime = 0;
    } else if (h >= 120 && h < 180) {
        r_prime = 0; g_prime = c; b_prime = x;
    } else if (h >= 180 && h < 240) {
        r_prime = 0; g_prime = x; b_prime = c;
    } else if (h >= 240 && h < 300) {
        r_prime = x; g_prime = 0; b_prime = c;
    } else {
        r_prime = c; g_prime = 0; b_prime = x;
    }
    
    uint8_t r = clamp((r_prime + m) * 255);
    uint8_t g = clamp((g_prime + m) * 255);
    uint8_t b = clamp((b_prime + m) * 255);
    
    return TXColor(r, g, b, a);
}

TXColor TXColor::fromHex(const std::string& hex) {
    std::string cleanHex = hex;
    
    // 移除开头的 #
    if (!cleanHex.empty() && cleanHex[0] == '#') {
        cleanHex = cleanHex.substr(1);
    }
    
    // 验证16进制字符
    for (char c : cleanHex) {
        if (!std::isxdigit(c)) {
            return TXColor(); // 返回默认黑色
        }
    }
    
    // 支持的格式: RGB (6位), ARGB (8位)
    if (cleanHex.length() == 6) {
        // RGB 格式，添加不透明度
        cleanHex = "FF" + cleanHex;
    } else if (cleanHex.length() != 8) {
        return TXColor(); // 返回默认黑色
    }
    
    try {
        unsigned long colorLong = std::stoul(cleanHex, nullptr, 16);
        return TXColor(static_cast<color_value_t>(colorLong));
    } catch (...) {
        return TXColor(); // 返回默认黑色
    }
}

// ==================== TXColor 辅助方法实现 ====================

uint8_t TXColor::clamp(double value) {
    return static_cast<uint8_t>(std::max(0.0, std::min(255.0, value)));
}

double TXColor::normalize(uint8_t value) {
    return value / 255.0;
}

// ==================== 预定义颜色对象实现 ====================

namespace ColorConstants {
    const TXColor BLACK(0xFF000000);     // 黑色
    const TXColor WHITE(0xFFFFFFFF);     // 白色  
    const TXColor RED(0xFFFF0000);       // 红色
    const TXColor GREEN(0xFF00FF00);     // 绿色
    const TXColor BLUE(0xFF0000FF);      // 蓝色
    const TXColor DARK_BLUE(0xFF000080); // 深蓝色
    const TXColor YELLOW(0xFFFFFF00);    // 黄色
    const TXColor CYAN(0xFF00FFFF);      // 青色
    const TXColor MAGENTA(0xFFFF00FF);   // 洋红色
    const TXColor GRAY(0xFF808080);      // 灰色
    const TXColor DARK_GRAY(0xFF404040); // 深灰色
    const TXColor LIGHT_GRAY(0xFFC0C0C0);// 浅灰色
    const TXColor TRANSPARENT_COLOR(0x00000000);// 透明
}

} // namespace TinaXlsx 
