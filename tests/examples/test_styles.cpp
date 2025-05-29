//
// Styles Test
// 测试字体、对齐、边框、填充等样式功能
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXStyle.hpp"
#include "TinaXlsx/TXColor.hpp"
#include <filesystem>
#include <iostream>

using namespace TinaXlsx;

class StylesTest : public ::testing::Test {
protected:
    void SetUp() override {
        std::filesystem::create_directories("output");
    }

    void TearDown() override {
        // std::filesystem::remove("output/styles_test.xlsx");
    }
};

/**
 * @brief 测试字体样式系统
 */
TEST_F(StylesTest, FontStyleSystem) {
    std::cout << "\n=== 字体样式系统测试 ===\n";
    
    // 测试默认字体
    TXFont font;
    EXPECT_EQ("Calibri", font.name);
    EXPECT_EQ(11U, font.size);
    EXPECT_EQ(ColorConstants::BLACK, font.color.getValue());
    EXPECT_FALSE(font.isBold());
    EXPECT_FALSE(font.isItalic());
    
    // 测试字体设置
    font.setName("Arial")
        .setSize(14)
        .setColor(ColorConstants::BLUE)
        .setBold(true)
        .setItalic(true)
        .setUnderline(true);
    
    EXPECT_EQ("Arial", font.name);
    EXPECT_EQ(14U, font.size);
    EXPECT_EQ(ColorConstants::BLUE, font.color.getValue());
    EXPECT_TRUE(font.isBold());
    EXPECT_TRUE(font.isItalic());
    EXPECT_TRUE(font.hasUnderline());
    
    // 测试字体样式枚举
    TXFont font2;
    font2.setStyle(FontStyle::Bold);
    EXPECT_TRUE(font2.isBold());
    EXPECT_FALSE(font2.isItalic());
    
    font2.setStyle(FontStyle::BoldItalic);
    EXPECT_TRUE(font2.isBold());
    EXPECT_TRUE(font2.isItalic());
    
    std::cout << "字体样式系统测试通过！\n";
}

/**
 * @brief 测试对齐样式系统
 */
TEST_F(StylesTest, AlignmentStyleSystem) {
    std::cout << "\n=== 对齐样式系统测试 ===\n";
    
    // 测试默认对齐
    TXAlignment alignment;
    EXPECT_EQ(HorizontalAlignment::Left, alignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Bottom, alignment.vertical);
    EXPECT_FALSE(alignment.wrapText);
    EXPECT_FALSE(alignment.shrinkToFit);
    EXPECT_EQ(0U, alignment.textRotation);
    EXPECT_EQ(0U, alignment.indent);
    
    // 测试对齐设置
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
    
    // 测试所有对齐选项
    alignment.setHorizontal(HorizontalAlignment::Right);
    EXPECT_EQ(HorizontalAlignment::Right, alignment.horizontal);
    
    alignment.setVertical(VerticalAlignment::Top);
    EXPECT_EQ(VerticalAlignment::Top, alignment.vertical);
    
    std::cout << "对齐样式系统测试通过！\n";
}

/**
 * @brief 测试边框样式系统
 */
TEST_F(StylesTest, BorderStyleSystem) {
    std::cout << "\n=== 边框样式系统测试 ===\n";
    
    // 测试默认边框
    TXBorder border;
    EXPECT_EQ(BorderStyle::None, border.leftStyle);
    EXPECT_EQ(BorderStyle::None, border.rightStyle);
    EXPECT_EQ(BorderStyle::None, border.topStyle);
    EXPECT_EQ(BorderStyle::None, border.bottomStyle);
    EXPECT_FALSE(border.diagonalUp);
    EXPECT_FALSE(border.diagonalDown);
    
    // 测试设置所有边框
    border.setAllBorders(BorderStyle::Thick, TXColor(ColorConstants::RED));
    EXPECT_EQ(BorderStyle::Thick, border.leftStyle);
    EXPECT_EQ(BorderStyle::Thick, border.rightStyle);
    EXPECT_EQ(BorderStyle::Thick, border.topStyle);
    EXPECT_EQ(BorderStyle::Thick, border.bottomStyle);
    EXPECT_EQ(ColorConstants::RED, border.leftColor.getValue());
    
    // 测试单独设置边框
    border.setLeftBorder(BorderStyle::Dotted, TXColor(ColorConstants::GREEN))
          .setRightBorder(BorderStyle::Dashed, TXColor(ColorConstants::BLUE))
          .setTopBorder(BorderStyle::Double, TXColor(ColorConstants::YELLOW))
          .setBottomBorder(BorderStyle::Medium, TXColor(ColorConstants::MAGENTA));
    
    EXPECT_EQ(BorderStyle::Dotted, border.leftStyle);
    EXPECT_EQ(BorderStyle::Dashed, border.rightStyle);
    EXPECT_EQ(BorderStyle::Double, border.topStyle);
    EXPECT_EQ(BorderStyle::Medium, border.bottomStyle);
    
    // 测试对角线边框
    border.setDiagonalBorder(BorderStyle::Thin, TXColor(ColorConstants::CYAN), true, false);
    EXPECT_EQ(BorderStyle::Thin, border.diagonalStyle);
    EXPECT_TRUE(border.diagonalUp);
    EXPECT_FALSE(border.diagonalDown);
    
    std::cout << "边框样式系统测试通过！\n";
}

/**
 * @brief 测试填充样式系统
 */
TEST_F(StylesTest, FillStyleSystem) {
    std::cout << "\n=== 填充样式系统测试 ===\n";
    
    // 测试默认填充
    TXFill fill;
    EXPECT_EQ(FillPattern::None, fill.pattern);
    EXPECT_EQ(ColorConstants::BLACK, fill.foregroundColor.getValue());
    EXPECT_EQ(ColorConstants::WHITE, fill.backgroundColor.getValue());
    
    // 测试实心填充
    fill.setSolidFill(TXColor(ColorConstants::YELLOW));
    EXPECT_EQ(FillPattern::Solid, fill.pattern);
    EXPECT_EQ(ColorConstants::YELLOW, fill.foregroundColor.getValue());
    EXPECT_EQ(ColorConstants::WHITE, fill.backgroundColor.getValue());
    
    // 测试其他填充模式
    fill.setPattern(FillPattern::Gray50)
        .setForegroundColor(TXColor(ColorConstants::BLUE))
        .setBackgroundColor(TXColor(ColorConstants::RED));
    
    EXPECT_EQ(FillPattern::Gray50, fill.pattern);
    EXPECT_EQ(ColorConstants::BLUE, fill.foregroundColor.getValue());
    EXPECT_EQ(ColorConstants::RED, fill.backgroundColor.getValue());
    
    std::cout << "填充样式系统测试通过！\n";
}

/**
 * @brief 测试完整单元格样式系统
 */
TEST_F(StylesTest, CellStyleSystem) {
    std::cout << "\n=== 单元格样式系统测试 ===\n";
    
    // 测试默认样式
    TXCellStyle style;
    
    const auto& font = style.getFont();
    EXPECT_EQ("Calibri", font.name);
    EXPECT_EQ(11U, font.size);
    
    const auto& alignment = style.getAlignment();
    EXPECT_EQ(HorizontalAlignment::Left, alignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Bottom, alignment.vertical);
    
    const auto& border = style.getBorder();
    EXPECT_EQ(BorderStyle::None, border.leftStyle);
    
    const auto& fill = style.getFill();
    EXPECT_EQ(FillPattern::None, fill.pattern);
    
    // 测试链式设置样式
    style.setFont("Times New Roman", 16)
         .setFontColor(ColorConstants::BLUE)
         .setFontStyle(FontStyle::Bold)
         .setHorizontalAlignment(HorizontalAlignment::Center)
         .setVerticalAlignment(VerticalAlignment::Middle)
         .setBackgroundColor(ColorConstants::YELLOW)
         .setAllBorders(BorderStyle::Thick, TXColor(ColorConstants::BLACK));
    
    // 验证设置结果
    const auto& updated_font = style.getFont();
    EXPECT_EQ("Times New Roman", updated_font.name);
    EXPECT_EQ(16U, updated_font.size);
    EXPECT_EQ(ColorConstants::BLUE, updated_font.color.getValue());
    EXPECT_TRUE(updated_font.isBold());
    
    const auto& updated_alignment = style.getAlignment();
    EXPECT_EQ(HorizontalAlignment::Center, updated_alignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Middle, updated_alignment.vertical);
    
    const auto& updated_fill = style.getFill();
    EXPECT_EQ(FillPattern::Solid, updated_fill.pattern);
    EXPECT_EQ(ColorConstants::YELLOW, updated_fill.foregroundColor.getValue());
    
    const auto& updated_border = style.getBorder();
    EXPECT_EQ(BorderStyle::Thick, updated_border.leftStyle);
    EXPECT_EQ(ColorConstants::BLACK, updated_border.leftColor.getValue());
    
    std::cout << "单元格样式系统测试通过！\n";
}

/**
 * @brief 测试预定义样式
 */
TEST_F(StylesTest, PredefinedStyles) {
    std::cout << "\n=== 预定义样式测试 ===\n";
    
    // 测试标题样式
    auto headerStyle = Styles::createHeaderStyle();
    const auto& headerFont = headerStyle.getFont();
    EXPECT_EQ("Calibri", headerFont.name);
    EXPECT_EQ(14U, headerFont.size);
    EXPECT_TRUE(headerFont.isBold());
    
    const auto& headerAlignment = headerStyle.getAlignment();
    EXPECT_EQ(HorizontalAlignment::Center, headerAlignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Middle, headerAlignment.vertical);
    
    // 测试数据样式
    auto dataStyle = Styles::createDataStyle();
    const auto& dataFont = dataStyle.getFont();
    EXPECT_EQ("Calibri", dataFont.name);
    EXPECT_EQ(11U, dataFont.size);
    EXPECT_FALSE(dataFont.isBold());
    
    const auto& dataAlignment = dataStyle.getAlignment();
    EXPECT_EQ(HorizontalAlignment::Left, dataAlignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Middle, dataAlignment.vertical);
    
    // 测试数字样式
    auto numberStyle = Styles::createNumberStyle();
    const auto& numberAlignment = numberStyle.getAlignment();
    EXPECT_EQ(HorizontalAlignment::Right, numberAlignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Middle, numberAlignment.vertical);
    
    // 测试高亮样式
    auto highlightStyle = Styles::createHighlightStyle(TXColor(ColorConstants::GREEN));
    const auto& highlightFont = highlightStyle.getFont();
    EXPECT_TRUE(highlightFont.isBold());
    
    const auto& highlightFill = highlightStyle.getFill();
    EXPECT_EQ(FillPattern::Solid, highlightFill.pattern);
    EXPECT_EQ(ColorConstants::GREEN, highlightFill.foregroundColor.getValue());
    
    // 测试表格样式
    auto tableStyle = Styles::createTableStyle();
    const auto& tableBorder = tableStyle.getBorder();
    EXPECT_EQ(BorderStyle::Thin, tableBorder.leftStyle);
    
    std::cout << "预定义样式测试通过！\n";
}

/**
 * @brief 测试颜色系统
 */
TEST_F(StylesTest, ColorSystem) {
    std::cout << "\n=== 颜色系统测试 ===\n";
    
    // 测试颜色创建
    TXColor red(255, 0, 0);
    TXColor green(0, 255, 0);
    TXColor blue(0, 0, 255);
    TXColor transparent(128, 128, 128, 128);
    
    EXPECT_EQ(ColorConstants::RED, red.getValue());
    EXPECT_EQ(ColorConstants::GREEN, green.getValue());
    EXPECT_EQ(ColorConstants::BLUE, blue.getValue());
    EXPECT_EQ(0x80808080U, transparent.getValue());
    
    // 测试16进制颜色
    auto hexRed = TXColor::fromHex("#FF0000");
    auto hexGreen = TXColor::fromHex("00FF00");
    auto hexBlue = TXColor::fromHex("#0000FF");
    
    EXPECT_EQ(ColorConstants::RED, hexRed.getValue());
    EXPECT_EQ(ColorConstants::GREEN, hexGreen.getValue());
    EXPECT_EQ(ColorConstants::BLUE, hexBlue.getValue());
    
    // 测试颜色分量
    auto [r, g, b, a] = red.getComponents();
    EXPECT_EQ(255, r);
    EXPECT_EQ(0, g);
    EXPECT_EQ(0, b);
    EXPECT_EQ(255, a);
    
    std::cout << "颜色系统测试通过！\n";
}

/**
 * @brief 测试样式应用到工作表
 */
TEST_F(StylesTest, ApplyStylesToWorksheet) {
    std::cout << "\n=== 样式应用到工作表测试 ===\n";
    
    TXWorkbook workbook;
    auto* sheet = workbook.addSheet("样式测试");
    ASSERT_NE(sheet, nullptr);
    
    // 设置标题
    sheet->setCellValue("A1", std::string("样式测试报表"));
    
    // 创建并应用标题样式
    TXCellStyle titleStyle;
    titleStyle.setFont("Arial", 18)
              .setFontColor(ColorConstants::WHITE)
              .setFontStyle(FontStyle::Bold)
              .setHorizontalAlignment(HorizontalAlignment::Center)
              .setVerticalAlignment(VerticalAlignment::Middle)
              .setBackgroundColor(TXColor(70, 130, 180).getValue())
              .setAllBorders(BorderStyle::Medium, TXColor(ColorConstants::BLACK));
    
    // 应用样式（目前只是API测试，实际样式还需要在XML写入时处理）
    bool styleApplied = sheet->setCellStyle("A1", titleStyle);
    EXPECT_TRUE(styleApplied);
    
    // 设置表头数据
    sheet->setCellValue("A3", std::string("产品"));
    sheet->setCellValue("B3", std::string("数量"));
    sheet->setCellValue("C3", std::string("单价"));
    sheet->setCellValue("D3", std::string("总计"));
    
    // 应用表头样式
    auto headerStyle = Styles::createHeaderStyle();
    EXPECT_GT(sheet->setRangeStyle(TXRange::fromAddress("A3:D3"), headerStyle), 0);
    
    // 添加数据行
    sheet->setCellValue("A4", std::string("产品A"));
    sheet->setCellValue("B4", 100.0);
    sheet->setCellValue("C4", 25.50);
    sheet->setCellValue("D4", 2550.0);
    
    sheet->setCellValue("A5", std::string("产品B"));
    sheet->setCellValue("B5", 200.0);
    sheet->setCellValue("C5", 15.75);
    sheet->setCellValue("D5", 3150.0);
    
    // 应用数据样式
    auto dataStyle = Styles::createDataStyle();
    EXPECT_GT(sheet->setRangeStyle(TXRange::fromAddress("A4:A5"), dataStyle), 0);
    
    auto numberStyle = Styles::createNumberStyle();
    EXPECT_GT(sheet->setRangeStyle(TXRange::fromAddress("B4:D5"), numberStyle), 0);
    
    // 保存文件
    bool saved = workbook.saveToFile("output/styles_test.xlsx");
    EXPECT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    
    std::cout << "样式应用到工作表测试通过！\n";
} 
