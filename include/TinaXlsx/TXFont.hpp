#pragma once

#include "TinaXlsx/TXResult.hpp"
#include "TinaXlsx/TXColor.hpp"
#include "TinaXlsx/TXTypes.hpp"
#include <string>
#include <memory>

namespace TinaXlsx
{
    /**
     * @brief 字体样式枚举
     */
    enum class FontStyle : uint8_t
    {
        Normal = 0, ///< 正常
        Bold = 1, ///< 粗体
        Italic = 2, ///< 斜体
        BoldItalic = 3, ///< 粗斜体
        Underline = 4, ///< 下划线
        Strikethrough = 8 ///< 删除线
    };

    /**
     * @brief 下划线样式
     */
    enum class UnderlineStyle : uint8_t
    {
        None = 0, ///< 无下划线
        Single = 1, ///< 单下划线
        Double = 2, ///< 双下划线
        SingleAccounting = 3, ///< 会计单下划线
        DoubleAccounting = 4 ///< 会计双下划线
    };

    /**
     * @brief 字体类 - 完整的字体样式管理
     * 
     * 独立管理字体的所有属性，包括验证和错误处理
     */
    class TXFont
    {
    public:
        // ==================== 构造函数 ====================

        /**
         * @brief 默认构造函数 - 创建默认字体(Calibri 11pt)
         */
        TXFont();

        /**
         * @brief 构造函数
         * @param name 字体名称
         * @param size 字体大小
         */
        explicit TXFont(std::string name, font_size_t size = DEFAULT_FONT_SIZE);

        /**
         * @brief 完整构造函数
         * @param name 字体名称
         * @param size 字体大小
         * @param color 字体颜色
         * @param bold 是否粗体
         * @param italic 是否斜体
         */
        TXFont(std::string name, font_size_t size, const TXColor& color,
               bool bold = false, bool italic = false);

        // 默认拷贝/移动构造和赋值
        TXFont(const TXFont&) = default;
        TXFont(TXFont&&) noexcept = default;
        TXFont& operator=(const TXFont&) = default;
        TXFont& operator=(TXFont&&) noexcept = default;

        ~TXFont() = default;

        // ==================== 基础属性设置 ====================

        /**
         * @brief 设置字体名称
         * @param name 字体名称
         * @return TXResult<void> 成功或包含错误信息
         */
        TXResult<void> setName(const std::string& name);

        /**
         * @brief 设置字体大小
         * @param size 字体大小(1-409磅)
         * @return TXResult<void> 成功或包含错误信息
         */
        TXResult<void> setSize(font_size_t size);

        /**
         * @brief 设置字体颜色
         * @param color 字体颜色
         * @return TXResult<void> 成功或包含错误信息
         */
        TXResult<void> setColor(const TXColor& color);

        /**
         * @brief 设置字体颜色(十六进制值)
         * @param colorValue 颜色值
         * @return TXResult<void> 成功或包含错误信息
         */
        TXResult<void> setColor(color_value_t colorValue);

        // ==================== 样式设置 ====================

        /**
         * @brief 设置粗体
         * @param enable 是否启用粗体
         * @return 自身引用，支持链式调用
         */
        TXFont& setBold(bool enable = true) noexcept;

        /**
         * @brief 设置斜体
         * @param enable 是否启用斜体
         * @return 自身引用，支持链式调用
         */
        TXFont& setItalic(bool enable = true) noexcept;

        /**
         * @brief 设置下划线样式
         * @param style 下划线样式
         * @return 自身引用，支持链式调用
         */
        TXFont& setUnderline(UnderlineStyle style = UnderlineStyle::Single) noexcept;

        /**
         * @brief 设置删除线
         * @param enable 是否启用删除线
         * @return 自身引用，支持链式调用
         */
        TXFont& setStrikethrough(bool enable = true) noexcept;

        /**
         * @brief 设置组合字体样式
         * @param style 字体样式枚举
         * @return 自身引用，支持链式调用
         */
        TXFont& setStyle(FontStyle style) noexcept;

        // ==================== 高级设置 ====================

        /**
         * @brief 设置字体字符集
         * @param charset 字符集ID (默认1为默认字符集)
         * @return TXResult<void> 成功或包含错误信息
         */
        [[nodiscard]] TXResult<void> setCharset(uint8_t charset);

        /**
         * @brief 设置字体族
         * @param family 字体族ID (1=Roman, 2=Swiss, 3=Modern, 4=Script, 5=Decorative)
         * @return TXResult<void> 成功或包含错误信息
         */
        [[nodiscard]] TXResult<void> setFamily(uint8_t family);

        /**
         * @brief 设置字体方案
         * @param scheme 方案类型 ("major", "minor", "none")
         * @return TXResult<void> 成功或包含错误信息
         */
        [[nodiscard]] TXResult<void> setScheme(const std::string& scheme);

        // ==================== 属性访问 ====================

        [[nodiscard]] const std::string& getName() const noexcept { return name_; }
        [[nodiscard]] font_size_t getSize() const noexcept { return size_; }
        [[nodiscard]] const TXColor& getColor() const noexcept { return color_; }
        [[nodiscard]] bool isBold() const noexcept { return bold_; }
        [[nodiscard]] bool isItalic() const noexcept { return italic_; }
        [[nodiscard]] UnderlineStyle getUnderlineStyle() const noexcept { return underline_style_; }
        [[nodiscard]] bool hasUnderline() const noexcept { return underline_style_ != UnderlineStyle::None; }
        [[nodiscard]] bool hasStrikethrough() const noexcept { return strikethrough_; }
        [[nodiscard]] uint8_t getCharset() const noexcept { return charset_; }
        [[nodiscard]] uint8_t getFamily() const noexcept { return family_; }
        [[nodiscard]] const std::string& getScheme() const noexcept { return scheme_; }

        // ==================== 便捷查询 ====================

        /**
         * @brief 检查字体是否有效
         * @return TXResult<void> 成功表示有效，错误包含问题描述
         */
        [[nodiscard]] TXResult<void> validate() const;

        /**
         * @brief 检查是否为默认字体
         * @return 是默认字体返回true
         */
        [[nodiscard]] bool isDefault() const noexcept;

        /**
         * @brief 获取字体的显示名称（用于UI显示）
         * @return 显示名称
         */
        [[nodiscard]] std::string getDisplayName() const;

        /**
         * @brief 获取字体的唯一标识键（用于样式管理器）
         * @return 唯一标识字符串
         */
        [[nodiscard]] std::string getUniqueKey() const;

        // ==================== 序列化支持 ====================

        /**
         * @brief 转换为字符串描述
         * @return 字体的字符串描述
         */
        [[nodiscard]] std::string toString() const;

        /**
         * @brief 从字符串描述创建字体
         * @param description 字体描述字符串
         * @return TXResult<TXFont> 成功返回字体对象，失败返回错误
         */
        [[nodiscard]] static TXResult<TXFont> fromString(const std::string& description);

        // ==================== 比较操作符 ====================

        bool operator==(const TXFont& other) const noexcept;
        bool operator!=(const TXFont& other) const noexcept { return !(*this == other); }
        bool operator<(const TXFont& other) const noexcept;

        // ==================== 静态工厂方法 ====================

        /**
         * @brief 创建默认字体
         * @return 默认字体对象
         */
        [[nodiscard]] static TXFont createDefault();

        /**
         * @brief 创建标题字体
         * @param size 字体大小
         * @return 标题字体对象
         */
        [[nodiscard]] static TXFont createHeading(font_size_t size = 16);

        /**
         * @brief 创建强调字体
         * @param base_font 基础字体
         * @return 强调字体对象
         */
        [[nodiscard]] static TXFont createEmphasis(const TXFont& base_font);

    private:
        // ==================== 私有成员 ====================

        std::string name_; ///< 字体名称
        font_size_t size_; ///< 字体大小
        TXColor color_; ///< 字体颜色
        bool bold_; ///< 是否粗体
        bool italic_; ///< 是否斜体
        UnderlineStyle underline_style_; ///< 下划线样式
        bool strikethrough_; ///< 是否删除线
        uint8_t charset_; ///< 字符集
        uint8_t family_; ///< 字体族
        std::string scheme_; ///< 字体方案

        // ==================== 私有方法 ====================

        /**
         * @brief 验证字体名称是否有效
         * @param name 字体名称
         * @return TXResult<void> 验证结果
         */
        [[nodiscard]] static TXResult<void> validateName(const std::string& name);

        /**
         * @brief 验证字体大小是否有效
         * @param size 字体大小
         * @return TXResult<void> 验证结果
         */
        [[nodiscard]] static TXResult<void> validateSize(font_size_t size);

        /**
         * @brief 验证字符集是否有效
         * @param charset 字符集
         * @return TXResult<void> 验证结果
         */
        [[nodiscard]] static TXResult<void> validateCharset(uint8_t charset);

        /**
         * @brief 验证字体族是否有效
         * @param family 字体族
         * @return TXResult<void> 验证结果
         */
        [[nodiscard]] static TXResult<void> validateFamily(uint8_t family);

        /**
         * @brief 验证字体方案是否有效
         * @param scheme 字体方案
         * @return TXResult<void> 验证结果
         */
        [[nodiscard]] static TXResult<void> validateScheme(const std::string& scheme);
    };
} // namespace TinaXlsx 
