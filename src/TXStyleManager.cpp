//
// Created by wuxianggujun on 2025/5/29.
//

#include "TinaXlsx/TXStyleManager.hpp"
#include "TinaXlsx/TXStyle.hpp"     //
#include "TinaXlsx/TXColor.hpp"     //
#include "TinaXlsx/TXTypes.hpp"     //
#include "TinaXlsx/TXNumberFormat.hpp" // 添加数字格式头文件

#include <sstream>
#include <algorithm> // For std::to_string on some compilers if not in <string>

namespace TinaXlsx
{
    // --- TXStyleManager Implementation ---

    TXStyleManager::TXStyleManager()
        : next_custom_num_fmt_id_(164)  // 自定义格式从164开始
    {
        initializeDefaultStyles();
        initializeBuiltinNumberFormats();
    }

    TXStyleManager::~TXStyleManager() = default;

    TXStyleManager::TXStyleManager(TXStyleManager&& other) noexcept
        : fonts_pool_(std::move(other.fonts_pool_))
        , fills_pool_(std::move(other.fills_pool_))
        , borders_pool_(std::move(other.borders_pool_))
        , cell_xfs_pool_(std::move(other.cell_xfs_pool_))
        , num_fmts_pool_new_(std::move(other.num_fmts_pool_new_))
        , num_fmt_lookup_new_(std::move(other.num_fmt_lookup_new_))
        , next_custom_num_fmt_id_(other.next_custom_num_fmt_id_)
        , font_lookup_(std::move(other.font_lookup_))
        , fill_lookup_(std::move(other.fill_lookup_))
        , border_lookup_(std::move(other.border_lookup_))
        , cell_xf_lookup_(std::move(other.cell_xf_lookup_))
        , style_cache_(std::move(other.style_cache_)) {
    }

    TXStyleManager& TXStyleManager::operator=(TXStyleManager&& other) noexcept {
        if (this != &other) {
            fonts_pool_ = std::move(other.fonts_pool_);
            fills_pool_ = std::move(other.fills_pool_);
            borders_pool_ = std::move(other.borders_pool_);
            cell_xfs_pool_ = std::move(other.cell_xfs_pool_);
            num_fmts_pool_new_ = std::move(other.num_fmts_pool_new_);
            num_fmt_lookup_new_ = std::move(other.num_fmt_lookup_new_);
            next_custom_num_fmt_id_ = other.next_custom_num_fmt_id_;
            font_lookup_ = std::move(other.font_lookup_);
            fill_lookup_ = std::move(other.fill_lookup_);
            border_lookup_ = std::move(other.border_lookup_);
            cell_xf_lookup_ = std::move(other.cell_xf_lookup_);
            style_cache_ = std::move(other.style_cache_);
        }
        return *this;
    }

    void TXStyleManager::initializeDefaultStyles()
    {
        // Font ID 0 (Default Font: Calibri, 11pt, Black, Normal)
        TXFont default_font; // Uses TXFont's default constructor
        std::string font_key_oss_str;
        {
            std::ostringstream font_key_oss;
            font_key_oss << default_font.getName() << "_" << default_font.getSize()
                << "_" << default_font.getColor().toARGBHexString() // Assuming TXColor has toARGBHexString()
                << "_b" << default_font.isBold() << "_i" << default_font.isItalic()
                << "_u" << default_font.hasUnderline() << "_s" << default_font.hasStrikethrough();
            font_key_oss_str = font_key_oss.str();
        }
        fonts_pool_.push_back(std::make_shared<TXFont>(default_font));
        font_lookup_[font_key_oss_str] = 0;

        // Fill ID 0 (No fill - patternType="none")
        TXFill no_fill(FillPattern::None); //
        std::string no_fill_key = fillPatternToString(no_fill.pattern);
        fills_pool_.push_back(std::make_shared<TXFill>(no_fill));
        fill_lookup_[no_fill_key] = 0;

        // Fill ID 1 (Gray125 fill - patternType="gray125") - Excel's second default fill
        TXFill gray125_fill(FillPattern::Gray125); //
        std::string gray125_fill_key = fillPatternToString(gray125_fill.pattern);
        // Ensure key is unique if colors matter for this pattern (they usually don't for gray125 directly)
        fills_pool_.push_back(std::make_shared<TXFill>(gray125_fill));
        fill_lookup_[gray125_fill_key] = 1;


        // Border ID 0 (No border)
        TXBorder no_border; // Uses TXBorder's default constructor
        std::ostringstream border_key_oss;
        border_key_oss << "L:" << borderStyleToString(no_border.leftStyle) << "_" << no_border.leftColor.toARGBHexString()
            << ";R:" << borderStyleToString(no_border.rightStyle) << "_" << no_border.rightColor.toARGBHexString()
            << ";T:" << borderStyleToString(no_border.topStyle) << "_" << no_border.topColor.toARGBHexString()
            << ";B:" << borderStyleToString(no_border.bottomStyle) << "_" << no_border.bottomColor.toARGBHexString()
            << ";D:" << borderStyleToString(no_border.diagonalStyle) << "_" << no_border.diagonalColor.toARGBHexString()
            << ";u" << no_border.diagonalUp << ";d" << no_border.diagonalDown;
        borders_pool_.push_back(std::make_shared<TXBorder>(no_border));
        border_lookup_[border_key_oss.str()] = 0;

        // CellXF ID 0 (Default XF referencing the default components)
        CellXF default_xf;
        default_xf.font_id_ = 0; // Default font registered above
        default_xf.fill_id_ = 0; // Default fill (none) registered above
        default_xf.border_id_ = 0; // Default border (none) registered above
        default_xf.num_fmt_id_ = 0; // 0 is "General" number format
        default_xf.xf_id_ = 0; // Points to the first <cellStyleXfs><xf> record (master formatting record)

        default_xf.apply_font_ = true; // true for default XF is typical
        default_xf.apply_fill_ = true;
        default_xf.apply_border_ = true;
        default_xf.apply_number_format_ = true;
        default_xf.apply_alignment_ = true; // Even if alignment is default, apply it from master XF

        // default_xf.alignment_ is already default-initialized by TXAlignment's constructor
        cell_xfs_pool_.push_back(default_xf);
        cell_xf_lookup_[default_xf.generateKey()] = 0;
    }


    u32 TXStyleManager::registerFont(const TXFont& font)
    {
        std::ostringstream key_ss;
        key_ss << font.getName() << "_" << font.getSize()
            << "_" << font.getColor().toARGBHexString()
            << "_b" << font.isBold() << "_i" << font.isItalic()
            << "_u" << font.hasUnderline() << "_s" << font.hasStrikethrough();
        // Consider also font.style enum if it carries more info than bold/italic/underline/strikethrough booleans
        std::string key = key_ss.str();

        auto it = font_lookup_.find(key);
        if (it != font_lookup_.end())
        {
            return it->second;
        }
        fonts_pool_.push_back(std::make_shared<TXFont>(font));
        u32 index = static_cast<u32>(fonts_pool_.size() - 1);
        font_lookup_[key] = index;
        return index;
    }

    u32 TXStyleManager::registerFill(const TXFill& fill)
    {
        std::ostringstream key_ss;
        key_ss << fillPatternToString(fill.pattern);
        if (fill.pattern != FillPattern::None)
        {
            // For "none", colors are irrelevant for key
            key_ss << "_fg:" << fill.foregroundColor.toARGBHexString();
            // For non-solid patterns, background color can also be part of the key if it's used
            if (fill.pattern != FillPattern::Solid)
            {
                key_ss << "_bg:" << fill.backgroundColor.toARGBHexString();
            }
        }
        std::string key = key_ss.str();

        auto it = fill_lookup_.find(key);
        if (it != fill_lookup_.end())
        {
            return it->second;
        }
        fills_pool_.push_back(std::make_shared<TXFill>(fill));
        u32 index = static_cast<u32>(fills_pool_.size() - 1);
        fill_lookup_[key] = index;
        return index;
    }

    u32 TXStyleManager::registerBorder(const TXBorder& border)
    {
        std::ostringstream key_ss;
        key_ss << "L:" << borderStyleToString(border.leftStyle) << "_" << border.leftColor.toARGBHexString()
            << ";R:" << borderStyleToString(border.rightStyle) << "_" << border.rightColor.toARGBHexString()
            << ";T:" << borderStyleToString(border.topStyle) << "_" << border.topColor.toARGBHexString()
            << ";B:" << borderStyleToString(border.bottomStyle) << "_" << border.bottomColor.toARGBHexString()
            << ";D:" << borderStyleToString(border.diagonalStyle) << "_" << border.diagonalColor.toARGBHexString()
            << ";u" << border.diagonalUp << ";d" << border.diagonalDown;
        std::string key = key_ss.str();

        auto it = border_lookup_.find(key);
        if (it != border_lookup_.end())
        {
            return it->second;
        }
        borders_pool_.push_back(std::make_shared<TXBorder>(border));
        u32 index = static_cast<u32>(borders_pool_.size() - 1);
        border_lookup_[key] = index;
        return index;
    }





    // --- XML String Converters ---
    std::string TXStyleManager::horizontalAlignmentToString(HorizontalAlignment alignment) const
    {
        switch (alignment)
        {
        case HorizontalAlignment::Left: return "left";
        case HorizontalAlignment::Center: return "center";
        case HorizontalAlignment::Right: return "right";
        case HorizontalAlignment::Fill: return "fill";
        case HorizontalAlignment::Justify: return "justify";
        case HorizontalAlignment::CenterAcrossSelection: return "centerContinuous";
        case HorizontalAlignment::General: return "general"; // Often omitted if it's general
        default: return "general"; // Fallback
        }
    }

    std::string TXStyleManager::verticalAlignmentToString(VerticalAlignment alignment) const
    {
        switch (alignment)
        {
        case VerticalAlignment::Top: return "top";
        case VerticalAlignment::Middle: return "center";
        case VerticalAlignment::Bottom: return "bottom";
        case VerticalAlignment::Justify: return "justify";
        case VerticalAlignment::Distributed: return "distributed";
        default: return "bottom"; // Fallback
        }
    }

    std::string TXStyleManager::borderStyleToString(BorderStyle style) const
    {
        // These need to map to valid OOXML border style strings
        switch (style)
        {
        case BorderStyle::None: return "none"; // Usually means omitting the border part
        case BorderStyle::Thin: return "thin";
        case BorderStyle::Medium: return "medium";
        case BorderStyle::Dashed: return "dashed";
        case BorderStyle::Dotted: return "dotted";
        case BorderStyle::Thick: return "thick";
        case BorderStyle::Double: return "double";
        case BorderStyle::DashDot: return "dashDot";
        case BorderStyle::DashDotDot: return "dashDotDot";
        // Add more as per OOXML spec for:
        // BorderStyle::SlantDashDot, MediumDashed, MediumDashDot, MediumDashDotDot
        default: return "none";
        }
    }

    std::string TXStyleManager::fillPatternToString(FillPattern pattern) const
    {
        // These need to map to valid OOXML patternType strings
        switch (pattern)
        {
        case FillPattern::None: return "none";
        case FillPattern::Solid: return "solid";
        case FillPattern::Gray50: return "mediumGray";
        case FillPattern::Gray75: return "darkGray";
        case FillPattern::Gray25: return "lightGray";
        case FillPattern::Gray125: return "gray125";
        case FillPattern::Gray0625: return "gray0625";
        // Add more: DarkHorizontal, DarkVertical, DarkDown, DarkUp, DarkGrid, DarkTrellis
        // LightHorizontal, LightVertical, LightDown, LightUp, LightGrid, LightTrellis
        default: return "none";
        }
    }


    XmlNodeBuilder TXStyleManager::createStylesXmlNode() const
    {
        XmlNodeBuilder styleSheet_node("styleSheet");
        styleSheet_node.addAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                       .addAttribute("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
                       .addAttribute("mc:Ignorable", "x14ac x16r2") // Common compatibility attributes
                       .addAttribute("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
                       .addAttribute("xmlns:x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main");

        // --- Number Formats (<numFmts>) ---
        XmlNodeBuilder numFmts_node("numFmts");
        
        // 使用新的数字格式池，只包含自定义格式 (ID >= 164)
        numFmts_node.addAttribute("count", std::to_string(num_fmts_pool_new_.size()));
        
        for (const auto& fmt_entry : num_fmts_pool_new_) {
            XmlNodeBuilder numFmt_node("numFmt");
            numFmt_node.addAttribute("numFmtId", std::to_string(fmt_entry.id_))
                       .addAttribute("formatCode", fmt_entry.formatCode_);
            numFmts_node.addChild(numFmt_node);
        }
        
        styleSheet_node.addChild(numFmts_node);

        // --- Fonts (<fonts>) ---
        XmlNodeBuilder fonts_node("fonts");
        fonts_node.addAttribute("count", std::to_string(fonts_pool_.size()));
        for (const auto& font_ptr : fonts_pool_)
        {
            XmlNodeBuilder font_node("font");
            font_node.addChild(XmlNodeBuilder("sz").addAttribute("val", std::to_string(font_ptr->getSize())));
            font_node.addChild(XmlNodeBuilder("color").addAttribute("rgb", font_ptr->getColor().toARGBHexString()));
            font_node.addChild(XmlNodeBuilder("name").addAttribute("val", font_ptr->getName()));
            if (font_ptr->isBold()) font_node.addChild(XmlNodeBuilder("b"));
            if (font_ptr->isItalic()) font_node.addChild(XmlNodeBuilder("i"));
            if (font_ptr->hasUnderline()) font_node.addChild(XmlNodeBuilder("u"));
            // Basic underline. Could be <u val="single"/> etc.
            if (font_ptr->hasStrikethrough()) font_node.addChild(XmlNodeBuilder("strike"));
            // Could also add <family val="..."/> and <scheme val="..."/> if needed.
            fonts_node.addChild(font_node);
        }
        styleSheet_node.addChild(fonts_node);

        // --- Fills (<fills>) ---
        XmlNodeBuilder fills_node("fills");
        fills_node.addAttribute("count", std::to_string(fills_pool_.size()));
        for (const auto& fill_ptr : fills_pool_)
        {
            XmlNodeBuilder fill_node("fill");
            XmlNodeBuilder patternFill_node("patternFill");
            patternFill_node.addAttribute("patternType", fillPatternToString(fill_ptr->pattern));
            if (fill_ptr->pattern == FillPattern::Solid)
            {
                patternFill_node.addChild(
                    XmlNodeBuilder("fgColor").addAttribute("rgb", fill_ptr->foregroundColor.toARGBHexString()));
            }
            else if (fill_ptr->pattern != FillPattern::None)
            {
                // For other patterns like gray125, etc.
                // fgColor is used for the pattern color itself for non-solid patterns too.
                patternFill_node.addChild(
                    XmlNodeBuilder("fgColor").addAttribute("rgb", fill_ptr->foregroundColor.toARGBHexString()));
                // bgColor can be specified for patterns if it's not the default (usually white/transparent)
                // if (fill_ptr->backgroundColor != ColorConstants::WHITE) { // Or some other default
                //    patternFill_node.addChild(XmlNodeBuilder("bgColor").addAttribute("rgb", fill_ptr->backgroundColor.toARGBHexString()));
                // }
            }
            fill_node.addChild(patternFill_node);
            fills_node.addChild(fill_node);
        }
        styleSheet_node.addChild(fills_node);

        // --- Borders (<borders>) ---
        XmlNodeBuilder borders_node("borders");
        borders_node.addAttribute("count", std::to_string(borders_pool_.size()));
        for (const auto& border_ptr : borders_pool_)
        {
            XmlNodeBuilder border_node("border");
            if (border_ptr->diagonalUp) border_node.addAttribute("diagonalUp", "1");
            if (border_ptr->diagonalDown) border_node.addAttribute("diagonalDown", "1");

            auto createBorderPartNode = [&](const std::string& partName, BorderStyle style, const TXColor& color)
            {
                XmlNodeBuilder part_node(partName);
                if (style != BorderStyle::None)
                {
                    part_node.addAttribute("style", borderStyleToString(style));
                    part_node.addChild(XmlNodeBuilder("color").addAttribute("rgb", color.toARGBHexString()));
                }
                // If style is None, an empty <left/>, <right/> etc. element is still written.
                return part_node;
            };
            border_node.addChild(createBorderPartNode("left", border_ptr->leftStyle, border_ptr->leftColor));
            border_node.addChild(createBorderPartNode("right", border_ptr->rightStyle, border_ptr->rightColor));
            border_node.addChild(createBorderPartNode("top", border_ptr->topStyle, border_ptr->topColor));
            border_node.addChild(createBorderPartNode("bottom", border_ptr->bottomStyle, border_ptr->bottomColor));
            border_node.
                addChild(createBorderPartNode("diagonal", border_ptr->diagonalStyle, border_ptr->diagonalColor));
            borders_node.addChild(border_node);
        }
        styleSheet_node.addChild(borders_node);

        // --- Cell Style XFs (<cellStyleXfs>) ---
        // These are master formatting records referenced by cellXfs. Typically at least one.
        XmlNodeBuilder cellStyleXfs_node("cellStyleXfs");
        cellStyleXfs_node.addAttribute("count", "1"); // Default count is 1 for the main "Normal" style XF
        XmlNodeBuilder cs_xf_node("xf"); // The master XF for "Normal" style
        cs_xf_node.addAttribute("numFmtId", "0").addAttribute("fontId", "0")
                  .addAttribute("fillId", "0").addAttribute("borderId", "0");
        cellStyleXfs_node.addChild(cs_xf_node);
        styleSheet_node.addChild(cellStyleXfs_node);

        // --- Cell XFs (<cellXfs>) ---
        // These are the actual formatting records applied to cells.
        XmlNodeBuilder cellXfs_node("cellXfs");
        cellXfs_node.addAttribute("count", std::to_string(cell_xfs_pool_.size()));
        for (const auto& xf_data : cell_xfs_pool_)
        {
            XmlNodeBuilder xf_node("xf");
            xf_node.addAttribute("numFmtId", std::to_string(xf_data.num_fmt_id_))
                   .addAttribute("fontId", std::to_string(xf_data.font_id_))
                   .addAttribute("fillId", std::to_string(xf_data.fill_id_))
                   .addAttribute("borderId", std::to_string(xf_data.border_id_))
                   .addAttribute("xfId", std::to_string(xf_data.xf_id_)); // Link to cellStyleXfs

            // Apply attributes only if true, to keep XML cleaner.
            // However, OOXML often expects these if the corresponding component ID is not 0
            // or if it's overriding the cellStyleXf.
            // A common practice is to always include applyFont etc. if the fontId is not 0.
            // For simplicity and explicitness here, let's add them if true.
            if (xf_data.apply_font_) xf_node.addAttribute("applyFont", "1");
            if (xf_data.apply_fill_) xf_node.addAttribute("applyFill", "1");
            if (xf_data.apply_border_) xf_node.addAttribute("applyBorder", "1");
            if (xf_data.apply_number_format_) xf_node.addAttribute("applyNumberFormat", "1");
            if (xf_data.apply_alignment_) xf_node.addAttribute("applyAlignment", "1");
            if (xf_data.apply_protection_) xf_node.addAttribute("applyProtection", "1");


            // Alignment (only add <alignment> element if any non-default alignment is set)
            const auto& align = xf_data.alignment_;
            bool has_specific_alignment = false;
            XmlNodeBuilder alignment_node("alignment");

            if (align.horizontal != HorizontalAlignment::General)
            {
                // Check against actual default
                alignment_node.addAttribute("horizontal", horizontalAlignmentToString(align.horizontal));
                has_specific_alignment = true;
            }
            if (align.vertical != VerticalAlignment::Bottom)
            {
                // Check against actual default
                alignment_node.addAttribute("vertical", verticalAlignmentToString(align.vertical));
                has_specific_alignment = true;
            }
            if (align.wrapText)
            {
                alignment_node.addAttribute("wrapText", "1");
                has_specific_alignment = true;
            }
            if (align.shrinkToFit)
            {
                alignment_node.addAttribute("shrinkToFit", "1");
                has_specific_alignment = true;
            }
            if (align.textRotation != 0)
            {
                alignment_node.addAttribute("textRotation", std::to_string(align.textRotation));
                has_specific_alignment = true;
            }
            if (align.indent != 0)
            {
                alignment_node.addAttribute("indent", std::to_string(align.indent));
                has_specific_alignment = true;
            }

            if (has_specific_alignment)
            {
                xf_node.addChild(alignment_node);
                // If specific alignment attributes are present, applyAlignment should be set.
                // The check for xf_data.apply_alignment_ above already handles this.
            }

            // Protection (只有当锁定状态不是默认值时才添加protection元素)
            if (xf_data.apply_protection_)
            {
                XmlNodeBuilder protection_node("protection");
                if (!xf_data.locked_)
                {
                    protection_node.addAttribute("locked", "0");
                }
                // 可以在这里添加其他保护属性，如hidden等
                xf_node.addChild(protection_node);
            }

            cellXfs_node.addChild(xf_node);
        }
        styleSheet_node.addChild(cellXfs_node);

        // --- Cell Styles (<cellStyles>) ---
        // Defines named styles like "Normal", "Heading 1", etc.
        XmlNodeBuilder cellStyles_node("cellStyles");
        cellStyles_node.addAttribute("count", "1"); // At least "Normal"
        XmlNodeBuilder cellStyle_item_node("cellStyle");
        cellStyle_item_node.addAttribute("name", "Normal")
                           .addAttribute("xfId", "0") // Points to the first XF in cellXfs (our default XF)
                           .addAttribute("builtinId", "0"); // 0 for "Normal"
        cellStyles_node.addChild(cellStyle_item_node);
        styleSheet_node.addChild(cellStyles_node);

        // --- Differential Formats (<dxfs>) --- (Used for conditional formatting)
        XmlNodeBuilder dxfs_node("dxfs");
        dxfs_node.addAttribute("count", "0"); // For conditional formatting styles
        styleSheet_node.addChild(dxfs_node);

        // --- Table Styles (<tableStyles>) ---
        XmlNodeBuilder tableStyles_node("tableStyles");
        tableStyles_node.addAttribute("count", "0") // For predefined table styles
                        .addAttribute("defaultTableStyle", "TableStyleMedium9") // Common default
                        .addAttribute("defaultPivotStyle", "PivotStyleLight16"); // Common default
        styleSheet_node.addChild(tableStyles_node);

        // Could also add <colors> (theme colors, indexed colors), <mruColors> etc. if needed

        return styleSheet_node;
    }

    // ==================== 内置数字格式映射表 ====================
    
    const std::unordered_map<std::string, u32> TXStyleManager::S_BUILTIN_NUMBER_FORMATS = {
        {"General", 0},
        {"0", 1},
        {"0.00", 2},
        {"#,##0", 3},
        {"#,##0.00", 4},
        // {"$#,##0_);($#,##0)", 5},  // 会因为货币符号不同而有变化
        // {"$#,##0_);[Red]($#,##0)", 6},
        // {"$#,##0.00_);($#,##0.00)", 7},
        // {"$#,##0.00_);[Red]($#,##0.00)", 8},
        {"0%", 9},
        {"0.00%", 10},
        {"0.00E+00", 11},
        {"# ?/?", 12},
        {"# ??/??", 13},
        {"m/d/yy", 14},
        {"d-mmm-yy", 15},
        {"d-mmm", 16},
        {"mmm-yy", 17},
        {"h:mm AM/PM", 18},
        {"h:mm:ss AM/PM", 19},
        {"h:mm", 20},
        {"h:mm:ss", 21},
        {"m/d/yy h:mm", 22},
        {"yyyy-mm-dd", 14},  // 常用日期格式，映射到14
        {"hh:mm:ss", 21},    // 常用时间格式，映射到21
        {"@", 49}            // 文本格式
    };

    void TXStyleManager::initializeBuiltinNumberFormats() {
        // 这里可以添加额外的初始化逻辑，如果需要的话
        // 目前静态映射表已经足够使用
    }

    // ==================== 新的数字格式注册方法 ====================
    
    u32 TXStyleManager::registerNumberFormat(const TXCellStyle::NumberFormatDefinition& definition) {
        // 验证格式定义
        if (!definition.isValid()) {
            // 如果格式定义无效，返回常规格式
            return 0;
        }
        
        // 如果是常规格式，直接返回0
        if (definition.isGeneral()) {
            return 0;
        }
        
        // 生成Excel格式代码
        std::string formatCode = definition.generateExcelFormatCode();
        
        // 检查是否为内置格式
        auto builtin_it = S_BUILTIN_NUMBER_FORMATS.find(formatCode);
        if (builtin_it != S_BUILTIN_NUMBER_FORMATS.end()) {
            return builtin_it->second;
        }
        
        // 检查是否已经注册过
        auto existing_it = num_fmt_lookup_new_.find(formatCode);
        if (existing_it != num_fmt_lookup_new_.end()) {
            return existing_it->second;
        }
        
        // 注册新的自定义格式
        u32 newId = next_custom_num_fmt_id_++;
        num_fmts_pool_new_.emplace_back(newId, formatCode);
        num_fmt_lookup_new_[formatCode] = newId;
        
        return newId;
    }

    // ==================== 新的样式注册方法 ====================
    
    u32 TXStyleManager::registerCellStyleXF(const TXCellStyle& style,
                                            bool applyFont, bool applyFill,
                                            bool applyBorder, bool applyAlignment) {
        CellXF xf_data;
        
        // 注册各个组件
        xf_data.font_id_ = registerFont(style.getFont());
        xf_data.fill_id_ = registerFill(style.getFill());
        xf_data.border_id_ = registerBorder(style.getBorder());
        
        // 注册数字格式
        const auto& num_fmt_def = style.getNumberFormatDefinition();
        xf_data.num_fmt_id_ = registerNumberFormat(num_fmt_def);
        xf_data.apply_number_format_ = !num_fmt_def.isGeneral();
        
        xf_data.xf_id_ = 0; // 总是指向cellStyleXfs的第一个XF
        
        // 设置应用标志
        xf_data.apply_font_ = applyFont && !style.getFont().isDefault();
        xf_data.apply_fill_ = applyFill && (style.getFill().pattern != FillPattern::None);
        xf_data.apply_border_ = applyBorder && (style.getBorder().leftStyle != BorderStyle::None ||
                                               style.getBorder().rightStyle != BorderStyle::None ||
                                               style.getBorder().topStyle != BorderStyle::None ||
                                               style.getBorder().bottomStyle != BorderStyle::None);
        xf_data.apply_alignment_ = applyAlignment && (
            style.getAlignment().horizontal != HorizontalAlignment::Left ||
            style.getAlignment().vertical != VerticalAlignment::Bottom ||
            style.getAlignment().wrapText ||
            style.getAlignment().shrinkToFit ||
            style.getAlignment().textRotation != 0 ||
            style.getAlignment().indent != 0
        );
        
        xf_data.alignment_ = style.getAlignment();

        // 设置保护信息
        xf_data.locked_ = style.isLocked();
        xf_data.apply_protection_ = (style.isLocked() != true); // 只有当锁定状态不是默认值时才应用保护

        // 检查是否已存在
        std::string key = xf_data.generateKey();
        auto it = cell_xf_lookup_.find(key);
        if (it != cell_xf_lookup_.end()) {
            return it->second;
        }
        
        // 添加新的XF
        cell_xfs_pool_.push_back(xf_data);
        u32 index = static_cast<u32>(cell_xfs_pool_.size() - 1);
        cell_xf_lookup_[key] = index;
        return index;
    }

    // ==================== 反向样式构造方法 ====================
    
    TXCellStyle TXStyleManager::getStyleObjectFromXfIndex(u32 xfIndex) const {
        // 检查索引有效性
        if (xfIndex >= cell_xfs_pool_.size()) {
            return TXCellStyle(); // 返回默认样式
        }
        
        // 检查缓存
        auto cache_it = style_cache_.find(xfIndex);
        if (cache_it != style_cache_.end()) {
            return cache_it->second;
        }
        
        const auto& xf = cell_xfs_pool_[xfIndex];
        TXCellStyle style;
        
        // 重构字体
        if (xf.font_id_ < fonts_pool_.size()) {
            style.setFont(*fonts_pool_[xf.font_id_]);
        }
        
        // 重构填充
        if (xf.fill_id_ < fills_pool_.size()) {
            style.setFill(*fills_pool_[xf.fill_id_]);
        }
        
        // 重构边框
        if (xf.border_id_ < borders_pool_.size()) {
            style.setBorder(*borders_pool_[xf.border_id_]);
        }
        
        // 重构对齐
        style.setAlignment(xf.alignment_);
        
        // 重构数字格式
        TXCellStyle::NumberFormatDefinition numFmtDef;
        if (xf.num_fmt_id_ == 0) {
            // 常规格式
            numFmtDef = TXCellStyle::NumberFormatDefinition();
        } else {
            // 尝试从内置格式中找到
            bool found = false;
            for (const auto& pair : S_BUILTIN_NUMBER_FORMATS) {
                if (pair.second == xf.num_fmt_id_) {
                    // 找到内置格式，尝试解析
                    const std::string& formatCode = pair.first;
                    numFmtDef = parseFormatCodeToDefinition(formatCode);
                    found = true;
                    break;
                }
            }
            
            if (!found) {
                // 查找自定义格式
                for (const auto& entry : num_fmts_pool_new_) {
                    if (entry.id_ == xf.num_fmt_id_) {
                        numFmtDef = parseFormatCodeToDefinition(entry.formatCode_);
                        found = true;
                        break;
                    }
                }
            }
            
            if (!found) {
                // 如果都找不到，使用常规格式
                numFmtDef = TXCellStyle::NumberFormatDefinition();
            }
        }
        
        style.setNumberFormatDefinition(numFmtDef);

        // 设置锁定状态
        style.setLocked(xf.locked_);

        // 缓存结果
        style_cache_[xfIndex] = style;
        
        return style;
    }

    TXCellStyle::NumberFormatDefinition TXStyleManager::parseFormatCodeToDefinition(const std::string& formatCode) const {
        // 简化的格式代码解析器
        if (formatCode == "General") {
            return TXCellStyle::NumberFormatDefinition(); // 常规格式
        }
        
        // 检查是否为百分比格式
        if (formatCode.find('%') != std::string::npos) {
            int decimalPlaces = 0;
            size_t dotPos = formatCode.find('.');
            if (dotPos != std::string::npos) {
                size_t percentPos = formatCode.find('%', dotPos);
                if (percentPos != std::string::npos) {
                    decimalPlaces = static_cast<int>(percentPos - dotPos - 1);
                }
            }
            return TXCellStyle::NumberFormatDefinition(TXNumberFormat::FormatType::Percentage, decimalPlaces);
        }
        
        // 检查是否为货币格式
        if (formatCode.find('$') != std::string::npos || formatCode.find('¥') != std::string::npos ||
            formatCode.find('€') != std::string::npos || formatCode.find('£') != std::string::npos) {
            int decimalPlaces = 0;
            bool useThousandSeparator = formatCode.find("#,##0") != std::string::npos;
            
            size_t dotPos = formatCode.find('.');
            if (dotPos != std::string::npos) {
                // 计算小数位数
                size_t pos = dotPos + 1;
                while (pos < formatCode.length() && formatCode[pos] == '0') {
                    decimalPlaces++;
                    pos++;
                }
            }
            
            std::string currencySymbol = "$";
            if (formatCode.find('¥') != std::string::npos) currencySymbol = "¥";
            else if (formatCode.find('€') != std::string::npos) currencySymbol = "€";
            else if (formatCode.find('£') != std::string::npos) currencySymbol = "£";
            
            return TXCellStyle::NumberFormatDefinition(TXNumberFormat::FormatType::Currency, 
                                                     decimalPlaces, useThousandSeparator, currencySymbol);
        }
        
        // 检查是否为科学计数法
        if (formatCode.find("E+") != std::string::npos || formatCode.find("e+") != std::string::npos) {
            return TXCellStyle::NumberFormatDefinition(TXNumberFormat::FormatType::Scientific, 2);
        }
        
        // 检查是否为日期格式
        if (formatCode.find("yyyy") != std::string::npos || formatCode.find("mm") != std::string::npos ||
            formatCode.find("dd") != std::string::npos || formatCode.find("yy") != std::string::npos) {
            if (formatCode.find("hh") != std::string::npos || formatCode.find("ss") != std::string::npos) {
                return TXCellStyle::NumberFormatDefinition(TXNumberFormat::FormatType::DateTime, 0);
            } else {
                return TXCellStyle::NumberFormatDefinition(TXNumberFormat::FormatType::Date, 0);
            }
        }
        
        // 检查是否为时间格式
        if (formatCode.find("hh") != std::string::npos || formatCode.find("mm") != std::string::npos ||
            formatCode.find("ss") != std::string::npos) {
            return TXCellStyle::NumberFormatDefinition(TXNumberFormat::FormatType::Time, 0);
        }
        
        // 检查是否为文本格式
        if (formatCode == "@") {
            return TXCellStyle::NumberFormatDefinition(TXNumberFormat::FormatType::Text, 0);
        }
        
        // 检查是否为数字格式
        if (formatCode.find('#') != std::string::npos || formatCode.find('0') != std::string::npos) {
            int decimalPlaces = 0;
            bool useThousandSeparator = formatCode.find("#,##0") != std::string::npos;
            
            size_t dotPos = formatCode.find('.');
            if (dotPos != std::string::npos) {
                // 计算小数位数
                size_t pos = dotPos + 1;
                while (pos < formatCode.length() && formatCode[pos] == '0') {
                    decimalPlaces++;
                    pos++;
                }
            }
            
            return TXCellStyle::NumberFormatDefinition(TXNumberFormat::FormatType::Number, 
                                                     decimalPlaces, useThousandSeparator);
        }
        
        // 默认情况下，作为自定义格式处理
        return TXCellStyle::NumberFormatDefinition(formatCode);
    }

} // namespace TinaXlsx
