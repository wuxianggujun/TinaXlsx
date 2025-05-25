//
// Styled Report Example Test
// Demonstrates TinaXlsx style system usage
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXStyle.hpp"
#include "TinaXlsx/TXTypes.hpp"
#include <filesystem>
#include <iostream>

using namespace TinaXlsx;

class StyledReportExampleTest : public ::testing::Test {
protected:
    void SetUp() override {
        std::filesystem::remove("StyledReport.xlsx");
    }

    void TearDown() override {
        std::filesystem::remove("StyledReport.xlsx");
    }
};

TEST_F(StyledReportExampleTest, CreateStyledFinancialReport) {
    std::cout << "=== Creating Styled Financial Report ===" << std::endl;
    
    TXWorkbook workbook;
    workbook.addSheet("Financial_Summary");
    
    auto sheet = workbook.getSheet("Financial_Summary");
    ASSERT_NE(sheet, nullptr);
    
    // 1. 创建标题样式 - 大字体、粗体、居中、蓝色背景
    TXCellStyle titleStyle;
    titleStyle.setFont("Arial", 18)
              .setFontColor(Colors::WHITE)
              .setFontStyle(FontStyle::Bold)
              .setHorizontalAlignment(HorizontalAlignment::Center)
              .setVerticalAlignment(VerticalAlignment::Middle)
              .setBackgroundColor(TXTypes::createColor(70, 130, 180)) // 钢蓝色
              .setAllBorders(BorderStyle::Medium, Colors::BLACK);
    
    // 2. 创建表头样式 - 中等字体、粗体、居中、灰色背景
    auto headerStyle = Styles::createHeaderStyle();
    
    // 3. 创建数据样式 - 左对齐
    auto dataStyle = Styles::createDataStyle();
    
    // 4. 创建数字样式 - 右对齐
    auto numberStyle = Styles::createNumberStyle();
    
    // 5. 创建强调样式 - 粗体、黄色背景
    auto highlightStyle = Styles::createHighlightStyle(Colors::YELLOW);
    
    // 6. 创建边框表格样式
    auto tableStyle = Styles::createTableStyle();
    
    // 设置主标题 (A1:E1 合并单元格)
    EXPECT_TRUE(sheet->setCellValue("A1", std::string("2024 Q3 Financial Report")));
    // 注意：这里暂时只设置A1，实际使用中需要合并单元格功能
    
    // 设置表头
    EXPECT_TRUE(sheet->setCellValue("A3", std::string("Category")));
    EXPECT_TRUE(sheet->setCellValue("B3", std::string("Q1 Amount")));
    EXPECT_TRUE(sheet->setCellValue("C3", std::string("Q2 Amount")));
    EXPECT_TRUE(sheet->setCellValue("D3", std::string("Q3 Amount")));
    EXPECT_TRUE(sheet->setCellValue("E3", std::string("Total")));
    
    // 设置数据行
    const std::vector<std::vector<std::string>> reportData = {
        {"Revenue", "1250000", "1380000", "1456000", "4086000"},
        {"Cost", "750000", "820000", "864000", "2434000"},
        {"Gross Profit", "500000", "560000", "592000", "1652000"},
        {"Operating Expense", "180000", "195000", "208000", "583000"},
        {"Net Profit", "320000", "365000", "384000", "1069000"}
    };
    
    for (size_t i = 0; i < reportData.size(); ++i) {
        TXTypes::RowIndex row = static_cast<TXTypes::RowIndex>(4 + i);
        
        // Project name (Column A)
        EXPECT_TRUE(sheet->setCellValue(TXTypes::coordinateToAddress(row, 1), reportData[i][0]));
        
        // Numeric data (Column B-E)
        for (size_t j = 1; j < reportData[i].size(); ++j) {
            TXTypes::ColIndex col = static_cast<TXTypes::ColIndex>(1 + j);
            double value = std::stod(reportData[i][j]);
            EXPECT_TRUE(sheet->setCellValue(TXTypes::coordinateToAddress(row, col), value));
        }
    }
    
    // 保存文件
    bool saved = workbook.saveToFile("StyledReport.xlsx");
    ASSERT_TRUE(saved) << "Failed to save styled report: " << workbook.getLastError();
    
    // 验证文件存在
    EXPECT_TRUE(std::filesystem::exists("StyledReport.xlsx"));
    
    std::cout << "Styled financial report created successfully!" << std::endl;
    
    // 验证数据完整性
    TXWorkbook verifyWorkbook;
    bool loaded = verifyWorkbook.loadFromFile("StyledReport.xlsx");
    ASSERT_TRUE(loaded) << "Failed to load styled report: " << verifyWorkbook.getLastError();
    
    auto verifySheet = verifyWorkbook.getSheet("Financial_Summary");
    ASSERT_NE(verifySheet, nullptr);
    
    // 验证标题
    auto titleValue = verifySheet->getCellValue("A1");
    EXPECT_FALSE(std::holds_alternative<std::monostate>(titleValue));
    
    // 验证数据
    auto revenueValue = verifySheet->getCellValue("B4"); // Q1 Revenue
    EXPECT_FALSE(std::holds_alternative<std::monostate>(revenueValue));
    
    std::cout << "Report verification completed!" << std::endl;
}

TEST_F(StyledReportExampleTest, StyleSystemAPI_Demo) {
    std::cout << "=== Style System API Demo ===" << std::endl;
    
    // 演示字体系统
    TXFont font;
    EXPECT_EQ("Calibri", font.name);
    EXPECT_EQ(TXTypes::DEFAULT_FONT_SIZE, font.size);
    
    font.setName("Arial")
        .setSize(14)
        .setColor(Colors::BLUE)
        .setBold(true)
        .setItalic(true)
        .setUnderline(true);
    
    EXPECT_EQ("Arial", font.name);
    EXPECT_EQ(14U, font.size);
    EXPECT_EQ(Colors::BLUE, font.color);
    EXPECT_TRUE(font.isBold());
    EXPECT_TRUE(font.isItalic());
    EXPECT_TRUE(font.hasUnderline());
    
    // 演示对齐系统
    TXAlignment alignment;
    alignment.setHorizontal(HorizontalAlignment::Center)
             .setVertical(VerticalAlignment::Middle)
             .setWrapText(true)
             .setTextRotation(45);
    
    EXPECT_EQ(HorizontalAlignment::Center, alignment.horizontal);
    EXPECT_EQ(VerticalAlignment::Middle, alignment.vertical);
    EXPECT_TRUE(alignment.wrapText);
    EXPECT_EQ(45U, alignment.textRotation);
    
    // 演示边框系统
    TXBorder border;
    border.setAllBorders(BorderStyle::Thick, Colors::RED)
          .setDiagonalBorder(BorderStyle::Dotted, Colors::GREEN);
    
    EXPECT_EQ(BorderStyle::Thick, border.leftStyle);
    EXPECT_EQ(BorderStyle::Thick, border.rightStyle);
    EXPECT_EQ(BorderStyle::Dotted, border.diagonalStyle);
    EXPECT_EQ(Colors::RED, border.leftColor);
    EXPECT_EQ(Colors::GREEN, border.diagonalColor);
    
    // 演示填充系统
    TXFill fill;
    fill.setSolidFill(Colors::YELLOW);
    
    EXPECT_EQ(FillPattern::Solid, fill.pattern);
    EXPECT_EQ(Colors::YELLOW, fill.foregroundColor);
    
    // 演示完整样式系统
    TXCellStyle style;
    style.setFont(font)
         .setAlignment(alignment)
         .setBorder(border)
         .setFill(fill);
    
    // 验证样式组件
    EXPECT_EQ("Arial", style.getFont().name);
    EXPECT_EQ(HorizontalAlignment::Center, style.getAlignment().horizontal);
    EXPECT_EQ(BorderStyle::Thick, style.getBorder().leftStyle);
    EXPECT_EQ(FillPattern::Solid, style.getFill().pattern);
    
    std::cout << "Style system API demo completed!" << std::endl;
}

TEST_F(StyledReportExampleTest, ColorSystemDemo) {
    std::cout << "=== Color System Demo ===" << std::endl;
    
    // 测试颜色创建和转换
    auto red = TXTypes::createColor(255, 0, 0);
    auto green = TXTypes::createColor(0, 255, 0);
    auto blue = TXTypes::createColor(0, 0, 255);
    auto transparent = TXTypes::createColor(128, 128, 128, 128);
    
    EXPECT_EQ(Colors::RED, red);
    EXPECT_EQ(Colors::GREEN, green);
    EXPECT_EQ(Colors::BLUE, blue);
    EXPECT_EQ(0x80808080U, transparent);
    
    // 测试16进制颜色解析
    auto hexRed = TXTypes::createColorFromHex("#FF0000");
    auto hexGreen = TXTypes::createColorFromHex("00FF00");
    auto hexBlue = TXTypes::createColorFromHex("#FF0000FF");
    
    EXPECT_EQ(Colors::RED, hexRed);
    EXPECT_EQ(Colors::GREEN, hexGreen);
    EXPECT_EQ(Colors::BLUE, hexBlue);
    
    // 测试颜色分量提取
    auto [r, g, b, a] = TXTypes::extractColorComponents(Colors::RED);
    EXPECT_EQ(255, r);
    EXPECT_EQ(0, g);
    EXPECT_EQ(0, b);
    EXPECT_EQ(255, a);
    
    // 测试坐标系统
    EXPECT_EQ("A1", TXTypes::coordinateToAddress(1, 1));
    EXPECT_EQ("Z26", TXTypes::coordinateToAddress(26, 26));
    EXPECT_EQ("AA27", TXTypes::coordinateToAddress(27, 27));
    
    auto [row, col] = TXTypes::addressToCoordinate("B5");
    EXPECT_EQ(5U, row);
    EXPECT_EQ(2U, col);
    
    // 测试列名转换
    EXPECT_EQ("A", TXTypes::colIndexToName(1));
    EXPECT_EQ("Z", TXTypes::colIndexToName(26));
    EXPECT_EQ("AA", TXTypes::colIndexToName(27));
    
    EXPECT_EQ(1U, TXTypes::colNameToIndex("A"));
    EXPECT_EQ(26U, TXTypes::colNameToIndex("Z"));
    EXPECT_EQ(27U, TXTypes::colNameToIndex("AA"));
    
    std::cout << "Color system demo completed!" << std::endl;
}

TEST_F(StyledReportExampleTest, PredefinedStylesDemo) {
    std::cout << "=== Predefined Styles Demo ===" << std::endl;
    
    TXWorkbook workbook;
    workbook.addSheet("Styles_Demo");
    
    auto sheet = workbook.getSheet("Styles_Demo");
    ASSERT_NE(sheet, nullptr);
    
    // 演示预定义样式
    EXPECT_TRUE(sheet->setCellValue("A1", std::string("Header Style")));
    EXPECT_TRUE(sheet->setCellValue("A2", std::string("Data Style")));
    EXPECT_TRUE(sheet->setCellValue("A3", std::string("Number Style")));
    EXPECT_TRUE(sheet->setCellValue("A4", std::string("Highlight Style")));
    EXPECT_TRUE(sheet->setCellValue("A5", std::string("Table Style")));
    
    EXPECT_TRUE(sheet->setCellValue("B1", std::string("Title Style Demonstration")));
    EXPECT_TRUE(sheet->setCellValue("B2", std::string("Data Style Demonstration")));
    EXPECT_TRUE(sheet->setCellValue("B3", static_cast<double>(12345.67)));
    EXPECT_TRUE(sheet->setCellValue("B4", std::string("Highlight Style Demonstration")));
    EXPECT_TRUE(sheet->setCellValue("B5", std::string("Table Style Demonstration")));
    
    // 验证预定义样式的属性
    auto headerStyle = Styles::createHeaderStyle();
    EXPECT_EQ(14U, headerStyle.getFont().size);
    EXPECT_TRUE(headerStyle.getFont().isBold());
    EXPECT_EQ(HorizontalAlignment::Center, headerStyle.getAlignment().horizontal);
    
    auto dataStyle = Styles::createDataStyle();
    EXPECT_EQ(11U, dataStyle.getFont().size);
    EXPECT_FALSE(dataStyle.getFont().isBold());
    EXPECT_EQ(HorizontalAlignment::Left, dataStyle.getAlignment().horizontal);
    
    auto numberStyle = Styles::createNumberStyle();
    EXPECT_EQ(HorizontalAlignment::Right, numberStyle.getAlignment().horizontal);
    
    auto highlightStyle = Styles::createHighlightStyle();
    EXPECT_TRUE(highlightStyle.getFont().isBold());
    EXPECT_EQ(FillPattern::Solid, highlightStyle.getFill().pattern);
    
    auto tableStyle = Styles::createTableStyle();
    EXPECT_EQ(BorderStyle::Thin, tableStyle.getBorder().leftStyle);
    
    // 保存演示文件
    bool saved = workbook.saveToFile("StyledReport.xlsx");
    EXPECT_TRUE(saved);
    
    std::cout << "Predefined styles demo completed!" << std::endl;
}

TEST_F(StyledReportExampleTest, StyleComparison) {
    std::cout << "=== Style Comparison Demo ===" << std::endl;
    
    // 测试样式相等性
    TXCellStyle style1;
    TXCellStyle style2;
    
    EXPECT_EQ(style1, style2); // 默认样式应该相等
    
    style1.setFont("Arial", 12)
          .setFontColor(Colors::RED);
    
    EXPECT_NE(style1, style2); // 修改后应该不等
    
    style2.setFont("Arial", 12)
          .setFontColor(Colors::RED);
    
    EXPECT_EQ(style1, style2); // 相同设置后应该相等
    
    // 测试拷贝和移动
    TXCellStyle style3(style1);
    EXPECT_EQ(style1, style3);
    
    TXCellStyle style4 = std::move(style3);
    EXPECT_EQ(style1, style4);
    
    // 测试重置
    style4.reset();
    EXPECT_EQ(style2, style1); // style1仍然保持原样
    EXPECT_NE(style4, style1); // style4已重置
    
    TXCellStyle defaultStyle;
    EXPECT_EQ(style4, defaultStyle); // 重置后应该等于默认样式
    
    std::cout << "Style comparison demo completed!" << std::endl;
} 
