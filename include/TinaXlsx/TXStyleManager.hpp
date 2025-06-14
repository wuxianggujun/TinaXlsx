#pragma once

#include <map>

#include "TXStyle.hpp" // 确保 TXFont, TXFill, TXBorder, TXCellStyle, FontStyle 等都在这里
#include "TXNumberFormat.hpp" // 添加数字格式头文件
#include "TXXmlWriter.hpp" // 用于生成XML
#include "TXColor.hpp"   // 确保 TXColor 被包含
#include "TXTypes.hpp"   // 确保 u32 等类型被包含

#include <vector>
#include <string>
#include <memory>
#include <unordered_map>
#include <sstream> // 用于Key的生成

namespace TinaXlsx
{
    class TXStyleManager
    {
    public:
        TXStyleManager();
        ~TXStyleManager();

        // 禁用拷贝构造和赋值
        TXStyleManager(const TXStyleManager&) = delete;
        TXStyleManager& operator=(const TXStyleManager&) = delete;

        // 支持移动构造和赋值
        TXStyleManager(TXStyleManager&& other) noexcept;
        TXStyleManager& operator=(TXStyleManager&& other) noexcept;

        // 注册独立的样式组件并返回其ID (索引)
        u32 registerFont(const TXFont& font);
        u32 registerFill(const TXFill& fill);
        u32 registerBorder(const TXBorder& border);
        // u32 registerNumberFormat(const std::string& formatCode, u32 id); // 自定义数字格式可以有ID

        /**
         * @brief 注册数字格式 (通过 NumberFormatDefinition)
         * @param definition 数字格式定义
         * @return 数字格式ID
         */
        u32 registerNumberFormat(const TXCellStyle::NumberFormatDefinition& definition);



        /**
         * @brief 注册一个完整的单元格样式 (XF record) - 新版本
         * 
         * 从 TXCellStyle 对象中提取所有样式组件（字体、填充、边框、对齐、数字格式），
         * 自动注册这些组件并创建XF记录。
         *
         * @param style TXCellStyle 对象，包含了所有样式信息
         * @param applyFont 是否应用此XF中的字体设置
         * @param applyFill 是否应用此XF中的填充设置
         * @param applyBorder 是否应用此XF中的边框设置
         * @param applyAlignment 是否应用此XF中的对齐设置
         * @return u32 此XF在styles.xml中<cellXfs>内的索引
         */
        u32 registerCellStyleXF(const TXCellStyle& style,
                                bool applyFont = true,
                                bool applyFill = true,
                                bool applyBorder = true,
                                bool applyAlignment = true);



        // 生成 styles.xml 内容的 XmlNodeBuilder 对象
        XmlNodeBuilder createStylesXmlNode() const;

        /**
         * @brief 从XF索引反向构造样式对象 (用于 getCellEffectiveStyle)
         * @param xfIndex XF记录的索引
         * @return 重构的TXCellStyle对象
         */
        TXCellStyle getStyleObjectFromXfIndex(u32 xfIndex) const;

    private:
        // 辅助函数，用于将枚举转换为XML字符串
        std::string horizontalAlignmentToString(HorizontalAlignment alignment) const;
        std::string verticalAlignmentToString(VerticalAlignment alignment) const;
        std::string borderStyleToString(BorderStyle style) const;
        std::string fillPatternToString(FillPattern pattern) const;

        // 内部结构，代表一个XF记录 (cellXfs中的一个xf元素)
        struct CellXF
        {
            u32 font_id_ = 0;
            u32 fill_id_ = 0;
            u32 border_id_ = 0;
            u32 num_fmt_id_ = 0;
            u32 xf_id_ = 0; // 指向cellStyleXfs的索引, 通常为0

            bool apply_font_ = false;
            bool apply_fill_ = false;
            bool apply_border_ = false;
            bool apply_alignment_ = false;
            bool apply_number_format_ = false;
            bool apply_protection_ = false;

            TXAlignment alignment_; // 对齐信息直接存储在XF中
            bool locked_ = true; // 单元格锁定状态，Excel默认为true

            // 用于比较和去重
            bool operator==(const CellXF& other) const
            {
                return font_id_ == other.font_id_ &&
                    fill_id_ == other.fill_id_ &&
                    border_id_ == other.border_id_ &&
                    num_fmt_id_ == other.num_fmt_id_ &&
                    xf_id_ == other.xf_id_ &&
                    apply_font_ == other.apply_font_ &&
                    apply_fill_ == other.apply_fill_ &&
                    apply_border_ == other.apply_border_ &&
                    apply_alignment_ == other.apply_alignment_ &&
                    apply_number_format_ == other.apply_number_format_ &&
                    apply_protection_ == other.apply_protection_ &&
                    alignment_ == other.alignment_ &&
                    locked_ == other.locked_;
            }

            // 生成唯一键用于去重
            std::string generateKey() const
            {
                std::ostringstream oss;
                oss << "f:" << font_id_ << ";fi:" << fill_id_ << ";b:" << border_id_
                    << ";n:" << num_fmt_id_ << ";xfid:" << xf_id_
                    << ";apF:" << apply_font_ << ";apFi:" << apply_fill_ << ";apB:" << apply_border_
                    << ";apA:" << apply_alignment_ << ";apN:" << apply_number_format_ << ";apP:" << apply_protection_
                    << ";alH:" << static_cast<int>(alignment_.horizontal)
                    << ";alV:" << static_cast<int>(alignment_.vertical)
                    << ";alWrap:" << alignment_.wrapText
                    << ";alShrink:" << alignment_.shrinkToFit
                    << ";alRot:" << alignment_.textRotation
                    << ";alIndent:" << alignment_.indent
                    << ";locked:" << locked_;
                return oss.str();
            }
        };

        // 数字格式条目结构 (用于XML生成)
        struct NumFmtEntry {
            u32 id_;            ///< numFmtId
            std::string formatCode_;  ///< Excel格式代码
            
            NumFmtEntry(u32 id, const std::string& formatCode)
                : id_(id), formatCode_(formatCode) {}
        };



        // 存储池
        std::vector<std::shared_ptr<TXFont>> fonts_pool_;
        std::vector<std::shared_ptr<TXFill>> fills_pool_;
        std::vector<std::shared_ptr<TXBorder>> borders_pool_;

        std::vector<CellXF> cell_xfs_pool_; // 存储 <cellXfs> 的 <xf> 元素数据

        // 新的数字格式管理
        std::vector<NumFmtEntry> num_fmts_pool_new_;  ///< 需要写入XML的自定义/非内置numFmt条目
        std::map<std::string, u32> num_fmt_lookup_new_;  ///< 从formatCode到numFmtId的映射
        u32 next_custom_num_fmt_id_;  ///< 自定义numFmtId的起始值

        // 用于快速查找已注册组件和XF的哈希表
        std::unordered_map<std::string, u32> font_lookup_;
        std::unordered_map<std::string, u32> fill_lookup_;
        std::unordered_map<std::string, u32> border_lookup_;

        std::unordered_map<std::string, u32> cell_xf_lookup_;

        // 样式反向构造缓存 (优化性能)
        mutable std::unordered_map<u32, TXCellStyle> style_cache_;

        void initializeDefaultStyles();
        void initializeBuiltinNumberFormats();  ///< 初始化内置数字格式映射
        
        /**
         * @brief 从Excel格式代码解析为NumberFormatDefinition
         * @param formatCode Excel格式代码字符串
         * @return 解析后的NumberFormatDefinition对象
         */
        TXCellStyle::NumberFormatDefinition parseFormatCodeToDefinition(const std::string& formatCode) const;
        
        // 内置格式映射 (从格式代码到numFmtId)
        static const std::unordered_map<std::string, u32> S_BUILTIN_NUMBER_FORMATS;
    };
} // namespace TinaXlsx
