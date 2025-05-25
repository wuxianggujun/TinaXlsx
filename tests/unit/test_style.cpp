//
// TXStyle and TXTypes Unit Tests
// Comprehensive testing for style and type systems
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXTypes.hpp"
#include "TinaXlsx/TXStyle.hpp"

using namespace TinaXlsx;

// ==================== TXTypes 测试 ====================

class TXTypesTest : public ::testing::Test {
protected:
    void SetUp() override {}
    void TearDown() override {}
};

TEST_F(TXTypesTest, ColumnIndexToName) {
    // 测试基本转换
    EXPECT_EQ("A", TXTypes::colIndexToName(1));
    EXPECT_EQ("B", TXTypes::colIndexToName(2));
    EXPECT_EQ("Z", TXTypes::colIndexToName(26));
    EXPECT_EQ("AA", TXTypes::colIndexToName(27));
    EXPECT_EQ("AB", TXTypes::colIndexToName(28));
    EXPECT_EQ("AZ", TXTypes::colIndexToName(52));
    EXPECT_EQ("BA", TXTypes::colIndexToName(53));
    
    // 测试边界值
    EXPECT_EQ("", TXTypes::colIndexToName(0));
    EXPECT_EQ("", TXTypes::colIndexToName(TXTypes::MAX_COLS + 1));
    
    // 测试最大值
    EXPECT_NE("", TXTypes::colIndexToName(TXTypes::MAX_COLS));
}

TEST_F(TXTypesTest, ColumnNameToIndex) {
    // 测试基本转换
    EXPECT_EQ(1U, TXTypes::colNameToIndex("A"));
    EXPECT_EQ(2U, TXTypes::colNameToIndex("B"));
    EXPECT_EQ(26U, TXTypes::colNameToIndex("Z"));
    EXPECT_EQ(27U, TXTypes::colNameToIndex("AA"));
    EXPECT_EQ(28U, TXTypes::colNameToIndex("AB"));
    EXPECT_EQ(52U, TXTypes::colNameToIndex("AZ"));
    EXPECT_EQ(53U, TXTypes::colNameToIndex("BA"));
    
    // 测试大小写不敏感
    EXPECT_EQ(1U, TXTypes::colNameToIndex("a"));
    EXPECT_EQ(27U, TXTypes::colNameToIndex("aa"));
    EXPECT_EQ(27U, TXTypes::colNameToIndex("Aa"));
    
    // 测试无效输入
    EXPECT_EQ(TXTypes::INVALID_COL, TXTypes::colNameToIndex(""));
    EXPECT_EQ(TXTypes::INVALID_COL, TXTypes::colNameToIndex("1"));
    EXPECT_EQ(TXTypes::INVALID_COL, TXTypes::colNameToIndex("A1"));
    EXPECT_EQ(TXTypes::INVALID_COL, TXTypes::colNameToIndex("@"));
}

TEST_F(TXTypesTest, CoordinateToAddress) {
    // 测试基本转换
    EXPECT_EQ("A1", TXTypes::coordinateToAddress(1, 1));
    EXPECT_EQ("B5", TXTypes::coordinateToAddress(5, 2));
    EXPECT_EQ("Z10", TXTypes::coordinateToAddress(10, 26));
    EXPECT_EQ("AA100", TXTypes::coordinateToAddress(100, 27));
    
    // 测试无效输入
    EXPECT_EQ("", TXTypes::coordinateToAddress(0, 1));
    EXPECT_EQ("", TXTypes::coordinateToAddress(1, 0));
    EXPECT_EQ("", TXTypes::coordinateToAddress(TXTypes::MAX_ROWS + 1, 1));
    EXPECT_EQ("", TXTypes::coordinateToAddress(1, TXTypes::MAX_COLS + 1));
}

TEST_F(TXTypesTest, AddressToCoordinate) {
    // 测试基本转换
    auto coord1 = TXTypes::addressToCoordinate("A1");
    EXPECT_EQ(1U, coord1.first);
    EXPECT_EQ(1U, coord1.second);
    
    auto coord2 = TXTypes::addressToCoordinate("B5");
    EXPECT_EQ(5U, coord2.first);
    EXPECT_EQ(2U, coord2.second);
    
    auto coord3 = TXTypes::addressToCoordinate("AA100");
    EXPECT_EQ(100U, coord3.first);
    EXPECT_EQ(27U, coord3.second);
    
    // 测试无效输入
    auto invalid1 = TXTypes::addressToCoordinate("");
    EXPECT_EQ(TXTypes::INVALID_ROW, invalid1.first);
    EXPECT_EQ(TXTypes::INVALID_COL, invalid1.second);
    
    auto invalid2 = TXTypes::addressToCoordinate("1A");
    EXPECT_EQ(TXTypes::INVALID_ROW, invalid2.first);
    EXPECT_EQ(TXTypes::INVALID_COL, invalid2.second);
    
    auto invalid3 = TXTypes::addressToCoordinate("A");
    EXPECT_EQ(TXTypes::INVALID_ROW, invalid3.first);
    EXPECT_EQ(TXTypes::INVALID_COL, invalid3.second);
}

TEST_F(TXTypesTest, ValidityChecks) {
    // 测试行有效性
    EXPECT_FALSE(TXTypes::isValidRow(0));
    EXPECT_TRUE(TXTypes::isValidRow(1));
    EXPECT_TRUE(TXTypes::isValidRow(TXTypes::MAX_ROWS));
    EXPECT_FALSE(TXTypes::isValidRow(TXTypes::MAX_ROWS + 1));
    
    // 测试列有效性
    EXPECT_FALSE(TXTypes::isValidCol(0));
    EXPECT_TRUE(TXTypes::isValidCol(1));
    EXPECT_TRUE(TXTypes::isValidCol(TXTypes::MAX_COLS));
    EXPECT_FALSE(TXTypes::isValidCol(TXTypes::MAX_COLS + 1));
    
    // 测试坐标有效性
    EXPECT_TRUE(TXTypes::isValidCoordinate(1, 1));
    EXPECT_FALSE(TXTypes::isValidCoordinate(0, 1));
    EXPECT_FALSE(TXTypes::isValidCoordinate(1, 0));
    EXPECT_FALSE(TXTypes::isValidCoordinate(0, 0));
}

TEST_F(TXTypesTest, ColorOperations) {
    // 测试颜色创建
    auto red = TXTypes::createColor(255, 0, 0);
    EXPECT_EQ(Colors::RED, red);
    
    auto green = TXTypes::createColor(0, 255, 0);
    EXPECT_EQ(Colors::GREEN, green);
    
    auto blue = TXTypes::createColor(0, 0, 255);
    EXPECT_EQ(Colors::BLUE, blue);
    
    // 测试透明度
    auto transparent_red = TXTypes::createColor(255, 0, 0, 128);
    EXPECT_EQ(0x80FF0000U, transparent_red);
    
    // 测试从16进制创建
    EXPECT_EQ(Colors::RED, TXTypes::createColorFromHex("#FF0000"));
    EXPECT_EQ(Colors::RED, TXTypes::createColorFromHex("FF0000"));
    EXPECT_EQ(Colors::RED, TXTypes::createColorFromHex("#FFFF0000"));
    
    // 测试颜色分量提取
    auto components = TXTypes::extractColorComponents(Colors::RED);
    EXPECT_EQ(255, std::get<0>(components)); // R
    EXPECT_EQ(0, std::get<1>(components));   // G
    EXPECT_EQ(0, std::get<2>(components));   // B
    EXPECT_EQ(255, std::get<3>(components)); // A
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
    EXPECT_EQ(TXTypes::DEFAULT_FONT_SIZE, font.size);
    EXPECT_EQ(TXTypes::DEFAULT_COLOR, font.color);
    EXPECT_EQ(FontStyle::Normal, font.style);
}

TEST_F(TXFontTest, ParameterizedConstructor) {
    TXFont font("Arial", 12);
    EXPECT_EQ("Arial", font.name);
    EXPECT_EQ(12U, font.size);
    EXPECT_EQ(TXTypes::DEFAULT_COLOR, font.color);
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
        .setColor(Colors::BLUE)
        .setBold(true)
        .setItalic(true);
    
    EXPECT_EQ("Times New Roman", font.name);
    EXPECT_EQ(14U, font.size);
    EXPECT_EQ(Colors::BLUE, font.color);
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
    border.setAllBorders(BorderStyle::Thin, Colors::BLACK);
    
    EXPECT_EQ(BorderStyle::Thin, border.leftStyle);
    EXPECT_EQ(BorderStyle::Thin, border.rightStyle);
    EXPECT_EQ(BorderStyle::Thin, border.topStyle);
    EXPECT_EQ(BorderStyle::Thin, border.bottomStyle);
    EXPECT_EQ(Colors::BLACK, border.leftColor);
    EXPECT_EQ(Colors::BLACK, border.rightColor);
    EXPECT_EQ(Colors::BLACK, border.topColor);
    EXPECT_EQ(Colors::BLACK, border.bottomColor);
}

TEST_F(TXBorderTest, IndividualBorders) {
    TXBorder border;
    
    border.setLeftBorder(BorderStyle::Thick, Colors::RED)
          .setRightBorder(BorderStyle::Thin, Colors::BLUE)
          .setTopBorder(BorderStyle::Double, Colors::GREEN)
          .setBottomBorder(BorderStyle::Dotted, Colors::YELLOW);
    
    EXPECT_EQ(BorderStyle::Thick, border.leftStyle);
    EXPECT_EQ(BorderStyle::Thin, border.rightStyle);
    EXPECT_EQ(BorderStyle::Double, border.topStyle);
    EXPECT_EQ(BorderStyle::Dotted, border.bottomStyle);
    
    EXPECT_EQ(Colors::RED, border.leftColor);
    EXPECT_EQ(Colors::BLUE, border.rightColor);
    EXPECT_EQ(Colors::GREEN, border.topColor);
    EXPECT_EQ(Colors::YELLOW, border.bottomColor);
}

TEST_F(TXBorderTest, DiagonalBorder) {
    TXBorder border;
    border.setDiagonalBorder(BorderStyle::Medium, Colors::GRAY, true, false);
    
    EXPECT_EQ(BorderStyle::Medium, border.diagonalStyle);
    EXPECT_EQ(Colors::GRAY, border.diagonalColor);
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
    EXPECT_EQ(TXTypes::DEFAULT_COLOR, fill.foregroundColor);
    EXPECT_EQ(Colors::WHITE, fill.backgroundColor);
}

TEST_F(TXFillTest, ParameterizedConstructor) {
    TXFill fill(FillPattern::Solid, Colors::RED, Colors::BLUE);
    EXPECT_EQ(FillPattern::Solid, fill.pattern);
    EXPECT_EQ(Colors::RED, fill.foregroundColor);
    EXPECT_EQ(Colors::BLUE, fill.backgroundColor);
}

TEST_F(TXFillTest, SolidFill) {
    TXFill fill;
    fill.setSolidFill(Colors::GREEN);
    
    EXPECT_EQ(FillPattern::Solid, fill.pattern);
    EXPECT_EQ(Colors::GREEN, fill.foregroundColor);
    EXPECT_EQ(Colors::WHITE, fill.backgroundColor);
}

TEST_F(TXFillTest, ChainedMethods) {
    TXFill fill;
    fill.setPattern(FillPattern::Gray50)
        .setForegroundColor(Colors::BLUE)
        .setBackgroundColor(Colors::YELLOW);
    
    EXPECT_EQ(FillPattern::Gray50, fill.pattern);
    EXPECT_EQ(Colors::BLUE, fill.foregroundColor);
    EXPECT_EQ(Colors::YELLOW, fill.backgroundColor);
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
    EXPECT_EQ(TXTypes::DEFAULT_FONT_SIZE, font.size);
    
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
          .setFontColor(Colors::RED)
          .setHorizontalAlignment(HorizontalAlignment::Center);
    
    // 测试拷贝构造
    TXCellStyle style2(style1);
    EXPECT_EQ(style1, style2);
    EXPECT_EQ("Arial", style2.getFont().name);
    EXPECT_EQ(14U, style2.getFont().size);
    EXPECT_EQ(Colors::RED, style2.getFont().color);
    
    // 测试拷贝赋值
    TXCellStyle style3;
    style3 = style1;
    EXPECT_EQ(style1, style3);
}

TEST_F(TXCellStyleTest, MoveSemantics) {
    TXCellStyle style1;
    style1.setFont("Arial", 14)
          .setFontColor(Colors::RED);
    
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
         .setFontColor(Colors::BLUE)
         .setFontStyle(FontStyle::Bold)
         .setHorizontalAlignment(HorizontalAlignment::Center)
         .setVerticalAlignment(VerticalAlignment::Middle)
         .setBackgroundColor(Colors::YELLOW)
         .setAllBorders(BorderStyle::Thick, Colors::BLACK);
    
    // 验证字体
    const auto& font = style.getFont();
    EXPECT_EQ("Times New Roman", font.name);
    EXPECT_EQ(16U, font.size);
    EXPECT_EQ(Colors::BLUE, font.color);
    EXPECT_TRUE(font.isBold());
    
    // 验证对齐
    const auto& alignment = style.getAlignment();
    EXPECT_EQ(HorizontalAlignment::Center, alignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Middle, alignment.vertical);
    
    // 验证填充
    const auto& fill = style.getFill();
    EXPECT_EQ(FillPattern::Solid, fill.pattern);
    EXPECT_EQ(Colors::YELLOW, fill.foregroundColor);
    
    // 验证边框
    const auto& border = style.getBorder();
    EXPECT_EQ(BorderStyle::Thick, border.leftStyle);
    EXPECT_EQ(Colors::BLACK, border.leftColor);
}

TEST_F(TXCellStyleTest, Reset) {
    TXCellStyle style;
    style.setFont("Arial", 16)
         .setFontColor(Colors::RED)
         .setBackgroundColor(Colors::BLUE);
    
    // 验证样式已设置
    EXPECT_EQ("Arial", style.getFont().name);
    EXPECT_EQ(Colors::RED, style.getFont().color);
    
    // 重置样式
    style.reset();
    
    // 验证已恢复默认值
    EXPECT_EQ("Calibri", style.getFont().name);
    EXPECT_EQ(TXTypes::DEFAULT_COLOR, style.getFont().color);
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
    EXPECT_EQ(Colors::LIGHT_GRAY, fill.foregroundColor);
    
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
    auto style = Styles::createHighlightStyle(Colors::GREEN);
    
    const auto& font = style.getFont();
    EXPECT_TRUE(font.isBold());
    
    const auto& fill = style.getFill();
    EXPECT_EQ(FillPattern::Solid, fill.pattern);
    EXPECT_EQ(Colors::GREEN, fill.foregroundColor);
}

TEST_F(PredefinedStylesTest, TableStyle) {
    auto style = Styles::createTableStyle();
    
    const auto& border = style.getBorder();
    EXPECT_EQ(BorderStyle::Thin, border.leftStyle);
    EXPECT_EQ(BorderStyle::Thin, border.rightStyle);
    EXPECT_EQ(BorderStyle::Thin, border.topStyle);
    EXPECT_EQ(BorderStyle::Thin, border.bottomStyle);
    EXPECT_EQ(Colors::GRAY, border.leftColor);
} 