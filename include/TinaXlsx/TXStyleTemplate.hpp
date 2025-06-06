//
// @file TXStyleTemplate.hpp
// @brief 样式模板系统 - 预设主题和样式模板
//

#pragma once

#include "TXStyle.hpp"
#include "TXColor.hpp"
#include "TXFont.hpp"
#include <string>
#include <vector>
#include <unordered_map>

namespace TinaXlsx {

/**
 * @brief 样式模板 - 预设样式组合
 */
class TXStyleTemplate {
public:
    /**
     * @brief 预设主题枚举
     */
    enum class Theme {
        Default,        ///< 默认主题
        Professional,   ///< 专业主题
        Modern,         ///< 现代主题
        Classic,        ///< 经典主题
        Colorful,       ///< 彩色主题
        Minimal         ///< 简约主题
    };

    /**
     * @brief 样式类型
     */
    enum class StyleType {
        Header,         ///< 标题样式
        SubHeader,      ///< 副标题样式
        Normal,         ///< 普通样式
        Highlight,      ///< 高亮样式
        Warning,        ///< 警告样式
        Error,          ///< 错误样式
        Success,        ///< 成功样式
        Number,         ///< 数字样式
        Currency,       ///< 货币样式
        Percentage      ///< 百分比样式
    };

    /**
     * @brief 构造函数
     * @param theme 主题
     */
    explicit TXStyleTemplate(Theme theme = Theme::Default);

    /**
     * @brief 获取指定类型的样式
     * @param type 样式类型
     * @return 样式对象
     */
    TXCellStyle getStyle(StyleType type) const;

    /**
     * @brief 设置主题
     * @param theme 主题
     */
    void setTheme(Theme theme);

    /**
     * @brief 获取当前主题
     * @return 主题
     */
    Theme getTheme() const { return current_theme_; }

    /**
     * @brief 自定义样式
     * @param type 样式类型
     * @param style 样式对象
     */
    void setCustomStyle(StyleType type, const TXCellStyle& style);

    /**
     * @brief 重置为默认样式
     */
    void resetToDefaults();

    /**
     * @brief 获取所有可用样式类型
     * @return 样式类型列表
     */
    static std::vector<StyleType> getAvailableStyleTypes();

    /**
     * @brief 获取主题名称
     * @param theme 主题
     * @return 主题名称
     */
    static std::string getThemeName(Theme theme);

private:
    Theme current_theme_;
    std::unordered_map<StyleType, TXCellStyle> custom_styles_;

    /**
     * @brief 初始化默认样式
     */
    void initializeDefaultStyles();

    /**
     * @brief 创建指定主题的样式
     * @param theme 主题
     * @param type 样式类型
     * @return 样式对象
     */
    TXCellStyle createThemeStyle(Theme theme, StyleType type) const;

    /**
     * @brief 获取主题颜色方案
     * @param theme 主题
     * @return 颜色方案
     */
    struct ColorScheme {
        TXColor primary;
        TXColor secondary;
        TXColor accent;
        TXColor background;
        TXColor text;
        TXColor highlight;
        TXColor warning;
        TXColor error;
        TXColor success;
    };

    ColorScheme getThemeColors(Theme theme) const;
};

/**
 * @brief 样式模板管理器 - 全局样式模板管理
 */
class TXStyleTemplateManager {
public:
    /**
     * @brief 获取单例实例
     * @return 管理器实例
     */
    static TXStyleTemplateManager& getInstance();

    /**
     * @brief 注册样式模板
     * @param name 模板名称
     * @param template_obj 模板对象
     */
    void registerTemplate(const std::string& name, const TXStyleTemplate& template_obj);

    /**
     * @brief 获取样式模板
     * @param name 模板名称
     * @return 模板对象指针，如果不存在返回nullptr
     */
    const TXStyleTemplate* getTemplate(const std::string& name) const;

    /**
     * @brief 获取所有模板名称
     * @return 模板名称列表
     */
    std::vector<std::string> getTemplateNames() const;

    /**
     * @brief 删除模板
     * @param name 模板名称
     * @return 是否成功删除
     */
    bool removeTemplate(const std::string& name);

    /**
     * @brief 清空所有模板
     */
    void clear();

private:
    std::unordered_map<std::string, TXStyleTemplate> templates_;

    TXStyleTemplateManager() = default;
    ~TXStyleTemplateManager() = default;
    TXStyleTemplateManager(const TXStyleTemplateManager&) = delete;
    TXStyleTemplateManager& operator=(const TXStyleTemplateManager&) = delete;
};

/**
 * @brief 快速样式创建工具
 */
namespace StyleTemplates {

/**
 * @brief 创建标题样式
 * @param theme 主题
 * @return 标题样式
 */
TXCellStyle createHeaderStyle(TXStyleTemplate::Theme theme = TXStyleTemplate::Theme::Default);

/**
 * @brief 创建数据样式
 * @param theme 主题
 * @return 数据样式
 */
TXCellStyle createDataStyle(TXStyleTemplate::Theme theme = TXStyleTemplate::Theme::Default);

/**
 * @brief 创建高亮样式
 * @param theme 主题
 * @return 高亮样式
 */
TXCellStyle createHighlightStyle(TXStyleTemplate::Theme theme = TXStyleTemplate::Theme::Default);

/**
 * @brief 创建货币样式
 * @param theme 主题
 * @return 货币样式
 */
TXCellStyle createCurrencyStyle(TXStyleTemplate::Theme theme = TXStyleTemplate::Theme::Default);

/**
 * @brief 创建百分比样式
 * @param theme 主题
 * @return 百分比样式
 */
TXCellStyle createPercentageStyle(TXStyleTemplate::Theme theme = TXStyleTemplate::Theme::Default);

} // namespace StyleTemplates

} // namespace TinaXlsx
