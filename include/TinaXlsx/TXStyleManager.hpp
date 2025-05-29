#pragma once

#include "TinaXlsx/TXStyle.hpp" // 确保 TXFont, TXFill, TXBorder, TXCellStyle, FontStyle 等都在这里
#include "TinaXlsx/TXXmlWriter.hpp" // 用于生成XML
#include "TinaXlsx/TXColor.hpp"   // 确保 TXColor 被包含
#include "TinaXlsx/TXTypes.hpp"   // 确保 u32 等类型被包含

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

        // 注册独立的样式组件并返回其ID (索引)
        u32 registerFont(const TXFont& font);
        u32 registerFill(const TXFill& fill);
        u32 registerBorder(const TXBorder& border);
        // u32 registerNumberFormat(const std::string& formatCode, u32 id); // 自定义数字格式可以有ID

        /**
         * @brief 注册一个完整的单元格样式 (XF record)
         *
         * @param style TXCellStyle 对象，包含了字体、填充、边框和对齐等信息
         * @param numFmtId 数字格式ID (0 代表 General)
         * @param applyFont 是否应用此XF中的字体设置
         * @param applyFill 是否应用此XF中的填充设置
         * @param applyBorder 是否应用此XF中的边框设置
         * @param applyAlignment 是否应用此XF中的对齐设置
         * @param applyNumberFormat 是否应用此XF中的数字格式设置
         * @return u32 此XF在styles.xml中<cellXfs>内的索引
         */
        u32 registerCellStyleXF(const TXCellStyle& style,
                                u32 numFmtId = 0, // 默认为 General
                                bool applyFont = true,
                                bool applyFill = true,
                                bool applyBorder = true,
                                bool applyAlignment = true,
                                bool applyNumberFormat = true);

        // 生成 styles.xml 内容的 XmlNodeBuilder 对象
        XmlNodeBuilder createStylesXmlNode() const;

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

            TXAlignment alignment_; // 对齐信息直接存储在XF中

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
                    alignment_ == other.alignment_; // TXAlignment 需要实现 operator==
            }

            // 生成唯一键用于去重
            std::string generateKey() const
            {
                std::ostringstream oss;
                oss << "f:" << font_id_ << ";fi:" << fill_id_ << ";b:" << border_id_
                    << ";n:" << num_fmt_id_ << ";xfid:" << xf_id_
                    << ";apF:" << apply_font_ << ";apFi:" << apply_fill_ << ";apB:" << apply_border_
                    << ";apA:" << apply_alignment_ << ";apN:" << apply_number_format_
                    << ";alH:" << static_cast<int>(alignment_.horizontal)
                    << ";alV:" << static_cast<int>(alignment_.vertical)
                    << ";alWrap:" << alignment_.wrapText
                    << ";alShrink:" << alignment_.shrinkToFit
                    << ";alRot:" << alignment_.textRotation
                    << ";alIndent:" << alignment_.indent;
                return oss.str();
            }
        };

        // 存储池
        std::vector<std::shared_ptr<TXFont>> fonts_pool_;
        std::vector<std::shared_ptr<TXFill>> fills_pool_;
        std::vector<std::shared_ptr<TXBorder>> borders_pool_;
        // std::vector<YourNumberFormatStruct> num_fmts_pool_; // 如果管理自定义数字格式

        std::vector<CellXF> cell_xfs_pool_; // 存储 <cellXfs> 的 <xf> 元素数据

        // 用于快速查找已注册组件和XF的哈希表
        std::unordered_map<std::string, u32> font_lookup_;
        std::unordered_map<std::string, u32> fill_lookup_;
        std::unordered_map<std::string, u32> border_lookup_;
        // std::unordered_map<std::string, u32> num_fmt_lookup_;
        std::unordered_map<std::string, u32> cell_xf_lookup_;

        void initializeDefaultStyles();
    };
} // namespace TinaXlsx
