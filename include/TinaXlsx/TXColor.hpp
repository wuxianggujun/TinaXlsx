#pragma once

#include "TXTypes.hpp"
#include <string>
#include <tuple>

namespace TinaXlsx {

/**
 * @brief 颜色类
 * 
 * 封装所有颜色相关的操作，包括创建、转换、解析等
 */
class TXColor {
public:
    // ==================== 构造函数 ====================
    
    /**
     * @brief 默认构造函数 - 黑色
     */
    TXColor() : value_(0xFF000000) {}
    
    /**
     * @brief 从颜色值构造
     * @param color ARGB颜色值
     */
    explicit TXColor(color_value_t color) : value_(color) {}
    
    /**
     * @brief 从RGB分量构造
     * @param r 红色分量 (0-255)
     * @param g 绿色分量 (0-255)
     * @param b 蓝色分量 (0-255)
     * @param a 透明度分量 (0-255, 默认255不透明)
     */
    TXColor(uint8_t r, uint8_t g, uint8_t b, uint8_t a = 255);
    
    /**
     * @brief 从16进制字符串构造
     * @param hex 16进制颜色字符串 (如 "#FF0000", "FF0000", "#FFFF0000")
     */
    explicit TXColor(const std::string& hex);
    
    // ==================== 获取器 ====================
    
    /**
     * @brief 获取ARGB颜色值
     * @return ARGB颜色值
     */
    color_value_t getValue() const { return value_; }
    
    /**
     * @brief 获取红色分量
     * @return 红色分量 (0-255)
     */
    uint8_t getRed() const;
    
    /**
     * @brief 获取绿色分量
     * @return 绿色分量 (0-255)
     */
    uint8_t getGreen() const;
    
    /**
     * @brief 获取蓝色分量
     * @return 蓝色分量 (0-255)
     */
    uint8_t getBlue() const;
    
    /**
     * @brief 获取透明度分量
     * @return 透明度分量 (0-255)
     */
    uint8_t getAlpha() const;
    
    /**
     * @brief 获取所有颜色分量
     * @return {r, g, b, a}
     */
    std::tuple<uint8_t, uint8_t, uint8_t, uint8_t> getComponents() const;
    
    // ==================== 设置器 ====================
    
    /**
     * @brief 设置ARGB颜色值
     * @param color ARGB颜色值
     * @return 自身引用，支持链式调用
     */
    TXColor& setValue(color_value_t color);
    
    /**
     * @brief 设置RGB分量
     * @param r 红色分量 (0-255)
     * @param g 绿色分量 (0-255)
     * @param b 蓝色分量 (0-255)
     * @return 自身引用，支持链式调用
     */
    TXColor& setRGB(uint8_t r, uint8_t g, uint8_t b);
    
    /**
     * @brief 设置ARGB分量
     * @param r 红色分量 (0-255)
     * @param g 绿色分量 (0-255)
     * @param b 蓝色分量 (0-255)
     * @param a 透明度分量 (0-255)
     * @return 自身引用，支持链式调用
     */
    TXColor& setARGB(uint8_t r, uint8_t g, uint8_t b, uint8_t a);
    
    /**
     * @brief 设置红色分量
     * @param r 红色分量 (0-255)
     * @return 自身引用，支持链式调用
     */
    TXColor& setRed(uint8_t r);
    
    /**
     * @brief 设置绿色分量
     * @param g 绿色分量 (0-255)
     * @return 自身引用，支持链式调用
     */
    TXColor& setGreen(uint8_t g);
    
    /**
     * @brief 设置蓝色分量
     * @param b 蓝色分量 (0-255)
     * @return 自身引用，支持链式调用
     */
    TXColor& setBlue(uint8_t b);
    
    /**
     * @brief 设置透明度分量
     * @param a 透明度分量 (0-255)
     * @return 自身引用，支持链式调用
     */
    TXColor& setAlpha(uint8_t a);
    
    // ==================== 转换方法 ====================
    
    /**
     * @brief 转换为16进制字符串
     * @param with_alpha 是否包含透明度 (默认true)
     * @param with_prefix 是否包含#前缀 (默认true)
     * @return 16进制字符串
     */
    std::string toHex(bool with_alpha = true, bool with_prefix = true) const;
    
    /**
     * @brief 转换为RGB字符串
     * @return RGB字符串，如 "rgb(255, 0, 0)"
     */
    std::string toRgbString() const;
    
    /**
     * @brief 转换为RGBA字符串
     * @return RGBA字符串，如 "rgba(255, 0, 0, 1.0)"
     */
    std::string toRgbaString() const;
    
    // ==================== 颜色操作 ====================
    
    /**
     * @brief 调整亮度
     * @param factor 亮度因子 (0.0-2.0, 1.0为原始亮度)
     * @return 新的颜色对象
     */
    TXColor adjustBrightness(double factor) const;
    
    /**
     * @brief 调整透明度
     * @param alpha 新的透明度 (0-255)
     * @return 新的颜色对象
     */
    TXColor withAlpha(uint8_t alpha) const;
    
    /**
     * @brief 混合颜色
     * @param other 另一个颜色
     * @param ratio 混合比例 (0.0-1.0, 0.0为当前颜色，1.0为另一个颜色)
     * @return 混合后的颜色
     */
    TXColor blend(const TXColor& other, double ratio) const;
    
    /**
     * @brief 获取补色
     * @return 补色
     */
    TXColor getComplementary() const;
    
    /**
     * @brief 判断是否为暗色
     * @return 暗色返回true，亮色返回false
     */
    bool isDark() const;
    
    /**
     * @brief 判断是否为亮色
     * @return 亮色返回true，暗色返回false
     */
    bool isLight() const { return !isDark(); }
    
    // ==================== 运算符重载 ====================
    
    bool operator==(const TXColor& other) const { return value_ == other.value_; }
    bool operator!=(const TXColor& other) const { return value_ != other.value_; }
    
    /**
     * @brief 隐式转换为颜色值
     * @return ARGB颜色值
     */
    operator color_value_t() const { return value_; }
    
    // ==================== 静态工厂方法 ====================
    
    /**
     * @brief 从HSV创建颜色
     * @param h 色相 (0-360)
     * @param s 饱和度 (0-100)
     * @param v 明度 (0-100)
     * @param a 透明度 (0-255)
     * @return 颜色对象
     */
    static TXColor fromHSV(int h, int s, int v, uint8_t a = 255);
    
    /**
     * @brief 从HSL创建颜色
     * @param h 色相 (0-360)
     * @param s 饱和度 (0-100)
     * @param l 亮度 (0-100)
     * @param a 透明度 (0-255)
     * @return 颜色对象
     */
    static TXColor fromHSL(int h, int s, int l, uint8_t a = 255);
    
    /**
     * @brief 解析16进制字符串
     * @param hex 16进制字符串
     * @return 颜色对象，解析失败返回黑色
     */
    static TXColor fromHex(const std::string& hex);

private:
    color_value_t value_;
    
    // 辅助方法
    static uint8_t clamp(double value);
    static double normalize(uint8_t value);
};

// ==================== 预定义颜色对象 ====================
namespace ColorConstants {
    extern const TXColor BLACK;
    extern const TXColor WHITE;
    extern const TXColor RED;
    extern const TXColor GREEN;
    extern const TXColor BLUE;
    extern const TXColor YELLOW;
    extern const TXColor CYAN;
    extern const TXColor MAGENTA;
    extern const TXColor GRAY;
    extern const TXColor DARK_GRAY;
    extern const TXColor LIGHT_GRAY;
    extern const TXColor TRANSPARENT;
}

} // namespace TinaXlsx 