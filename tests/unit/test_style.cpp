//
// TXStyle and TXTypes Unit Tests
// Comprehensive testing for style and type systems
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXTypes.hpp"
#include "TinaXlsx/TXStyle.hpp"
#include "TinaXlsx/TXCoordinate.hpp"

using namespace TinaXlsx;

// ==================== TXTypes 测试 ====================

class TXTypesTest : public ::testing::Test {
protected:
    void SetUp() override {}
    void TearDown() override {}
};

TEST_F(TXTypesTest, ColumnIndexToName) {
    // 测试基本转换
    EXPECT_EQ("A", column_t::column_string_from_index(1));
    EXPECT_EQ("B", column_t::column_string_from_index(2));
    EXPECT_EQ("Z", column_t::column_string_from_index(26));
    EXPECT_EQ("AA", column_t::column_string_from_index(27));
    EXPECT_EQ("AB", column_t::column_string_from_index(28));
    EXPECT_EQ("AZ", column_t::column_string_from_index(52));
    EXPECT_EQ("BA", column_t::column_string_from_index(53));
    
    // 测试边界情况
    EXPECT_EQ("", column_t::column_string_from_index(0));
    EXPECT_EQ("", column_t::column_string_from_index(column_t::MAX_COLUMNS + 1));
    
    // 测试大数值
    EXPECT_EQ("XFD", column_t::column_string_from_index(column_t::MAX_COLUMNS));
}

TEST_F(TXTypesTest, ColumnNameToIndex) {
    // 测试基本转换
    EXPECT_EQ(1U, column_t::column_index_from_string("A"));
    EXPECT_EQ(2U, column_t::column_index_from_string("B"));
    EXPECT_EQ(26U, column_t::column_index_from_string("Z"));
    EXPECT_EQ(27U, column_t::column_index_from_string("AA"));
    EXPECT_EQ(28U, column_t::column_index_from_string("AB"));
    EXPECT_EQ(52U, column_t::column_index_from_string("AZ"));
    EXPECT_EQ(53U, column_t::column_index_from_string("BA"));
    
    // 测试大写字母
    EXPECT_EQ(1U, column_t::column_index_from_string("A"));
    EXPECT_EQ(27U, column_t::column_index_from_string("AA"));
    
    // 测试无效输入
    EXPECT_EQ(0U, column_t::column_index_from_string(""));
    EXPECT_EQ(0U, column_t::column_index_from_string("1"));
    EXPECT_EQ(0U, column_t::column_index_from_string("A1"));
}

TEST_F(TXTypesTest, CoordinateToAddress) {
    // 测试基本转换
    EXPECT_EQ("A1", TXCoordinate(row_t(1), column_t(1)).toAddress());
    EXPECT_EQ("B5", TXCoordinate(row_t(5), column_t(2)).toAddress());
    EXPECT_EQ("Z26", TXCoordinate(row_t(26), column_t(26)).toAddress());
    EXPECT_EQ("AA100", TXCoordinate(row_t(100), column_t(27)).toAddress());
    
    // 测试无效坐标
    EXPECT_FALSE(TXCoordinate(row_t(0), column_t(1)).isValid());
    EXPECT_FALSE(TXCoordinate(row_t(1), column_t(static_cast<u32>(0))).isValid());
    EXPECT_FALSE(TXCoordinate(row_t(row_t::MAX_ROWS + 1), column_t(1)).isValid());
    EXPECT_FALSE(TXCoordinate(row_t(1), column_t(column_t::MAX_COLUMNS + 1)).isValid());
}

TEST_F(TXTypesTest, AddressToCoordinate) {
    // 测试基本转换
    TXCoordinate coord1("A1");
    EXPECT_EQ(1U, coord1.getRow().index());
    EXPECT_EQ(1U, coord1.getCol().index());
    
    TXCoordinate coord2("B5");
    EXPECT_EQ(5U, coord2.getRow().index());
    EXPECT_EQ(2U, coord2.getCol().index());
    
    TXCoordinate coord3("AA100");
    EXPECT_EQ(100U, coord3.getRow().index());
    EXPECT_EQ(27U, coord3.getCol().index());
    
    // 测试无效输入 - 期望构造失败或者isValid()返回false
    try {
        TXCoordinate invalid1("");
        EXPECT_FALSE(invalid1.isValid());
    } catch (...) {
        // 构造失败是期望的
    }
    
    try {
        TXCoordinate invalid2("1A");
        EXPECT_FALSE(invalid2.isValid());
    } catch (...) {
        // 构造失败是期望的
    }
    
    try {
        TXCoordinate invalid3("A");
        EXPECT_FALSE(invalid3.isValid());
    } catch (...) {
        // 构造失败是期望的
    }
}

TEST_F(TXTypesTest, ValidityChecks) {
    // 测试行有效性
    EXPECT_FALSE(row_t(0).is_valid());
    EXPECT_TRUE(row_t(1).is_valid());
    EXPECT_TRUE(row_t(row_t::MAX_ROWS).is_valid());
    EXPECT_FALSE(row_t(row_t::MAX_ROWS + 1).is_valid());
    
    // 测试列有效性
    EXPECT_FALSE(column_t(static_cast<u32>(0)).is_valid());
    EXPECT_TRUE(column_t(static_cast<u32>(1)).is_valid());
    EXPECT_TRUE(column_t(column_t::MAX_COLUMNS).is_valid());
    EXPECT_FALSE(column_t(column_t::MAX_COLUMNS + 1).is_valid());
    
    // 测试坐标有效性
    EXPECT_TRUE(TXCoordinate::isValidCoordinate(row_t(1), column_t(static_cast<u32>(1))));
    EXPECT_FALSE(TXCoordinate::isValidCoordinate(row_t(0), column_t(static_cast<u32>(1))));
    EXPECT_FALSE(TXCoordinate::isValidCoordinate(row_t(1), column_t(static_cast<u32>(0))));
    EXPECT_FALSE(TXCoordinate::isValidCoordinate(row_t(0), column_t(static_cast<u32>(0))));
    
    // 测试字体大小有效性 - 已移除，使用类内部验证
    
    // 测试工作表名称有效性 - 已移除，使用类内部验证
}

// ==================== TXFont 测试 ====================

class TXFontTest : public ::testing::Test {
protected:
    void SetUp() override {}
    void TearDown() override {}
};

TEST_F(TXFontTest, DefaultConstructor) {
    TXFont font;
    EXPECT_EQ("Calibri", font.name);
    EXPECT_EQ(DEFAULT_FONT_SIZE, font.size);
    EXPECT_EQ(DEFAULT_COLOR, font.color);
    EXPECT_EQ(FontStyle::Normal, font.style);
}

TEST_F(TXFontTest, ParameterizedConstructor) {
    TXFont font("Arial", 12);
    EXPECT_EQ("Arial", font.name);
    EXPECT_EQ(12U, font.size);
    EXPECT_EQ(DEFAULT_COLOR, font.color);
    EXPECT_EQ(FontStyle::Normal, font.style);
}

TEST_F(TXFontTest, StyleMethods) {
    TXFont font;
    
    // 测试粗体
    font.setBold(true);
    EXPECT_TRUE(font.isBold());
    EXPECT_FALSE(font.isItalic());
    
    // 测试斜体
    font.setItalic(true);
    EXPECT_TRUE(font.isBold());
    EXPECT_TRUE(font.isItalic());
    
    // 测试下划线
    font.setUnderline(true);
    EXPECT_TRUE(font.hasUnderline());
    
    // 测试删除线
    font.setStrikethrough(true);
    EXPECT_TRUE(font.hasStrikethrough());
    
    // 测试取消粗体
    font.setBold(false);
    EXPECT_FALSE(font.isBold());
    EXPECT_TRUE(font.isItalic()); // 其他样式保持
}

TEST_F(TXFontTest, ChainedCalls) {
    TXFont font;
    font.setName("Times New Roman")
        .setSize(14)
        .setColor(ColorConstants::BLUE)
        .setBold(true)
        .setItalic(true);
    
    EXPECT_EQ("Times New Roman", font.name);
    EXPECT_EQ(14U, font.size);
    EXPECT_EQ(ColorConstants::BLUE, font.color);
    EXPECT_TRUE(font.isBold());
    EXPECT_TRUE(font.isItalic());
}

TEST_F(TXFontTest, Equality) {
    TXFont font1("Arial", 12);
    TXFont font2("Arial", 12);
    TXFont font3("Calibri", 12);
    
    EXPECT_EQ(font1, font2);
    EXPECT_NE(font1, font3);
    
    font1.setBold(true);
    EXPECT_NE(font1, font2);
    
    font2.setBold(true);
    EXPECT_EQ(font1, font2);
}

// ==================== TXAlignment 测试 ====================

class TXAlignmentTest : public ::testing::Test {
protected:
    void SetUp() override {}
    void TearDown() override {}
};

TEST_F(TXAlignmentTest, DefaultValues) {
    TXAlignment alignment;
    EXPECT_EQ(HorizontalAlignment::Left, alignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Bottom, alignment.vertical);
    EXPECT_FALSE(alignment.wrapText);
    EXPECT_FALSE(alignment.shrinkToFit);
    EXPECT_EQ(0U, alignment.textRotation);
    EXPECT_EQ(0U, alignment.indent);
}

TEST_F(TXAlignmentTest, ChainedMethods) {
    TXAlignment alignment;
    alignment.setHorizontal(HorizontalAlignment::Center)
             .setVertical(VerticalAlignment::Middle)
             .setWrapText(true)
             .setShrinkToFit(true)
             .setTextRotation(45)
             .setIndent(2);
    
    EXPECT_EQ(HorizontalAlignment::Center, alignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Middle, alignment.vertical);
    EXPECT_TRUE(alignment.wrapText);
    EXPECT_TRUE(alignment.shrinkToFit);
    EXPECT_EQ(45U, alignment.textRotation);
    EXPECT_EQ(2U, alignment.indent);
}

TEST_F(TXAlignmentTest, Equality) {
    TXAlignment alignment1;
    TXAlignment alignment2;
    
    EXPECT_EQ(alignment1, alignment2);
    
    alignment1.setHorizontal(HorizontalAlignment::Center);
    EXPECT_NE(alignment1, alignment2);
    
    alignment2.setHorizontal(HorizontalAlignment::Center);
    EXPECT_EQ(alignment1, alignment2);
}

// ==================== TXBorder 测试 ====================

class TXBorderTest : public ::testing::Test {
protected:
    void SetUp() override {}
    void TearDown() override {}
};

TEST_F(TXBorderTest, DefaultValues) {
    TXBorder border;
    EXPECT_EQ(BorderStyle::None, border.leftStyle);
    EXPECT_EQ(BorderStyle::None, border.rightStyle);
    EXPECT_EQ(BorderStyle::None, border.topStyle);
    EXPECT_EQ(BorderStyle::None, border.bottomStyle);
    EXPECT_EQ(BorderStyle::None, border.diagonalStyle);
    EXPECT_FALSE(border.diagonalUp);
    EXPECT_FALSE(border.diagonalDown);
}

TEST_F(TXBorderTest, SetAllBorders) {
    TXBorder border;
    border.setAllBorders(BorderStyle::Thin, ColorConstants::BLACK);
    
    EXPECT_EQ(BorderStyle::Thin, border.leftStyle);
    EXPECT_EQ(BorderStyle::Thin, border.rightStyle);
    EXPECT_EQ(BorderStyle::Thin, border.topStyle);
    EXPECT_EQ(BorderStyle::Thin, border.bottomStyle);
    EXPECT_EQ(ColorConstants::BLACK, border.leftColor);
    EXPECT_EQ(ColorConstants::BLACK, border.rightColor);
    EXPECT_EQ(ColorConstants::BLACK, border.topColor);
    EXPECT_EQ(ColorConstants::BLACK, border.bottomColor);
}

TEST_F(TXBorderTest, IndividualBorders) {
    TXBorder border;
    
    border.setLeftBorder(BorderStyle::Thick, ColorConstants::RED)
          .setRightBorder(BorderStyle::Thin, ColorConstants::BLUE)
          .setTopBorder(BorderStyle::Double, ColorConstants::GREEN)
          .setBottomBorder(BorderStyle::Dotted, ColorConstants::YELLOW);
    
    EXPECT_EQ(BorderStyle::Thick, border.leftStyle);
    EXPECT_EQ(BorderStyle::Thin, border.rightStyle);
    EXPECT_EQ(BorderStyle::Double, border.topStyle);
    EXPECT_EQ(BorderStyle::Dotted, border.bottomStyle);
    
    EXPECT_EQ(ColorConstants::RED, border.leftColor);
    EXPECT_EQ(ColorConstants::BLUE, border.rightColor);
    EXPECT_EQ(ColorConstants::GREEN, border.topColor);
    EXPECT_EQ(ColorConstants::YELLOW, border.bottomColor);
}

TEST_F(TXBorderTest, DiagonalBorder) {
    TXBorder border;
    border.setDiagonalBorder(BorderStyle::Medium, ColorConstants::GRAY, true, false);
    
    EXPECT_EQ(BorderStyle::Medium, border.diagonalStyle);
    EXPECT_EQ(ColorConstants::GRAY, border.diagonalColor);
    EXPECT_TRUE(border.diagonalUp);
    EXPECT_FALSE(border.diagonalDown);
}

// ==================== TXFill 测试 ====================

class TXFillTest : public ::testing::Test {
protected:
    void SetUp() override {}
    void TearDown() override {}
};

TEST_F(TXFillTest, DefaultValues) {
    TXFill fill;
    EXPECT_EQ(FillPattern::None, fill.pattern);
    EXPECT_EQ(ColorConstants::BLACK, fill.foregroundColor);
    EXPECT_EQ(ColorConstants::WHITE, fill.backgroundColor);
}

TEST_F(TXFillTest, ParameterizedConstructor) {
    TXFill fill(FillPattern::Solid, ColorConstants::RED, ColorConstants::BLUE);
    EXPECT_EQ(FillPattern::Solid, fill.pattern);
    EXPECT_EQ(ColorConstants::RED, fill.foregroundColor);
    EXPECT_EQ(ColorConstants::BLUE, fill.backgroundColor);
}

TEST_F(TXFillTest, SolidFill) {
    TXFill fill;
    fill.setSolidFill(ColorConstants::GREEN);
    
    EXPECT_EQ(FillPattern::Solid, fill.pattern);
    EXPECT_EQ(ColorConstants::GREEN, fill.foregroundColor);
    EXPECT_EQ(ColorConstants::WHITE, fill.backgroundColor);
}

TEST_F(TXFillTest, ChainedMethods) {
    TXFill fill;
    fill.setPattern(FillPattern::Gray50)
        .setForegroundColor(ColorConstants::BLUE)
        .setBackgroundColor(ColorConstants::YELLOW);
    
    EXPECT_EQ(FillPattern::Gray50, fill.pattern);
    EXPECT_EQ(ColorConstants::BLUE, fill.foregroundColor);
    EXPECT_EQ(ColorConstants::YELLOW, fill.backgroundColor);
}

// ==================== TXCellStyle 测试 ====================

class TXCellStyleTest : public ::testing::Test {
protected:
    void SetUp() override {}
    void TearDown() override {}
};

TEST_F(TXCellStyleTest, DefaultConstructor) {
    TXCellStyle style;
    
    // 检查默认字体
    const auto& font = style.getFont();
    EXPECT_EQ("Calibri", font.name);
    EXPECT_EQ(11.0, font.size);  // 默认字体大小
    
    // 检查默认对齐
    const auto& alignment = style.getAlignment();
    EXPECT_EQ(HorizontalAlignment::Left, alignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Bottom, alignment.vertical);
    
    // 检查默认边框
    const auto& border = style.getBorder();
    EXPECT_EQ(BorderStyle::None, border.leftStyle);
    
    // 检查默认填充
    const auto& fill = style.getFill();
    EXPECT_EQ(FillPattern::None, fill.pattern);
}

TEST_F(TXCellStyleTest, CopySemantics) {
    TXCellStyle style1;
    style1.setFont("Arial", 14)
          .setFontColor(ColorConstants::RED)
          .setHorizontalAlignment(HorizontalAlignment::Center);
    
    // 测试拷贝构造
    TXCellStyle style2(style1);
    EXPECT_EQ(style1, style2);
    EXPECT_EQ("Arial", style2.getFont().name);
    EXPECT_EQ(14U, style2.getFont().size);
    EXPECT_EQ(ColorConstants::RED, style2.getFont().color);
    
    // 测试拷贝赋值
    TXCellStyle style3;
    style3 = style1;
    EXPECT_EQ(style1, style3);
}

TEST_F(TXCellStyleTest, MoveSemantics) {
    TXCellStyle style1;
    style1.setFont("Arial", 14)
          .setFontColor(ColorConstants::RED);
    
    TXCellStyle style2(style1); // 备份用于比较
    
    // 测试移动构造
    TXCellStyle style3(std::move(style1));
    EXPECT_EQ(style2, style3);
    
    // 测试移动赋值
    TXCellStyle style4;
    style4 = std::move(style3);
    EXPECT_EQ(style2, style4);
}

TEST_F(TXCellStyleTest, ChainedMethods) {
    TXCellStyle style;
    style.setFont("Times New Roman", 16)
         .setFontColor(ColorConstants::BLUE)
         .setFontStyle(FontStyle::Bold)
         .setHorizontalAlignment(HorizontalAlignment::Center)
         .setVerticalAlignment(VerticalAlignment::Middle)
         .setBackgroundColor(ColorConstants::YELLOW)
         .setAllBorders(BorderStyle::Thick, ColorConstants::BLACK);
    
    // 验证字体
    const auto& font = style.getFont();
    EXPECT_EQ("Times New Roman", font.name);
    EXPECT_EQ(16U, font.size);
    EXPECT_EQ(ColorConstants::BLUE, font.color);
    EXPECT_TRUE(font.isBold());
    
    // 验证对齐
    const auto& alignment = style.getAlignment();
    EXPECT_EQ(HorizontalAlignment::Center, alignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Middle, alignment.vertical);
    
    // 验证填充
    const auto& fill = style.getFill();
    EXPECT_EQ(FillPattern::Solid, fill.pattern);
    EXPECT_EQ(ColorConstants::YELLOW, fill.foregroundColor);
    
    // 验证边框
    const auto& border = style.getBorder();
    EXPECT_EQ(BorderStyle::Thick, border.leftStyle);
    EXPECT_EQ(ColorConstants::BLACK, border.leftColor);
}

TEST_F(TXCellStyleTest, Reset) {
    TXCellStyle style;
    style.setFont("Arial", 16)
         .setFontColor(ColorConstants::RED)
         .setBackgroundColor(ColorConstants::BLUE);
    
    // 验证样式已设置
    EXPECT_EQ("Arial", style.getFont().name);
    EXPECT_EQ(ColorConstants::RED, style.getFont().color);
    
    // 重置样式
    style.reset();
    
    // 验证已恢复默认值
    EXPECT_EQ("Calibri", style.getFont().name);
    EXPECT_EQ(ColorConstants::BLACK, style.getFont().color);
    EXPECT_EQ(FillPattern::None, style.getFill().pattern);
}

// ==================== 预定义样式测试 ====================

class PredefinedStylesTest : public ::testing::Test {
protected:
    void SetUp() override {}
    void TearDown() override {}
};

TEST_F(PredefinedStylesTest, HeaderStyle) {
    auto style = Styles::createHeaderStyle();
    
    const auto& font = style.getFont();
    EXPECT_EQ("Calibri", font.name);
    EXPECT_EQ(14U, font.size);
    EXPECT_TRUE(font.isBold());
    
    const auto& alignment = style.getAlignment();
    EXPECT_EQ(HorizontalAlignment::Center, alignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Middle, alignment.vertical);
    
    const auto& fill = style.getFill();
    EXPECT_EQ(FillPattern::Solid, fill.pattern);
    EXPECT_EQ(ColorConstants::LIGHT_GRAY, fill.foregroundColor);
    
    const auto& border = style.getBorder();
    EXPECT_EQ(BorderStyle::Thin, border.leftStyle);
}

TEST_F(PredefinedStylesTest, DataStyle) {
    auto style = Styles::createDataStyle();
    
    const auto& font = style.getFont();
    EXPECT_EQ("Calibri", font.name);
    EXPECT_EQ(11U, font.size);
    EXPECT_FALSE(font.isBold());
    
    const auto& alignment = style.getAlignment();
    EXPECT_EQ(HorizontalAlignment::Left, alignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Middle, alignment.vertical);
}

TEST_F(PredefinedStylesTest, NumberStyle) {
    auto style = Styles::createNumberStyle();
    
    const auto& alignment = style.getAlignment();
    EXPECT_EQ(HorizontalAlignment::Right, alignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Middle, alignment.vertical);
}

TEST_F(PredefinedStylesTest, HighlightStyle) {
    auto style = Styles::createHighlightStyle(ColorConstants::GREEN);
    
    const auto& font = style.getFont();
    EXPECT_TRUE(font.isBold());
    
    const auto& fill = style.getFill();
    EXPECT_EQ(FillPattern::Solid, fill.pattern);
    EXPECT_EQ(ColorConstants::GREEN, fill.foregroundColor);
}

TEST_F(PredefinedStylesTest, TableStyle) {
    auto style = Styles::createTableStyle();
    
    const auto& border = style.getBorder();
    EXPECT_EQ(BorderStyle::Thin, border.leftStyle);
    EXPECT_EQ(BorderStyle::Thin, border.rightStyle);
    EXPECT_EQ(BorderStyle::Thin, border.topStyle);
    EXPECT_EQ(BorderStyle::Thin, border.bottomStyle);
    EXPECT_EQ(ColorConstants::GRAY, border.leftColor);
} 
