#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <cstdio>

using namespace TinaXlsx;

class CellStyleTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("StyleTest");
        ASSERT_NE(sheet, nullptr);
    }

    void TearDown() override {
        workbook.reset();
        // 清理测试文件
        // std::remove("test_styles.xlsx");
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet;
};

// 测试基本字体样式
TEST_F(CellStyleTest, FontStyling) {
    // 创建字体样式
    TXFont font;
    font.setName("Arial");
    font.setSize(12);
    font.setBold(true);
    font.setItalic(true);
    font.setColor(TXColor(255, 0, 0)); // 红色
    
    // 创建单元格样式
    TXCellStyle style;
    style.setFont(font);
    
    // 设置单元格值和样式
    sheet->setCellValue(row_t(1), column_t(1), std::string("Bold Red Text"));
    EXPECT_TRUE(sheet->setCellStyle(row_t(1), column_t(1), style));
    
    // 验证样式是否应用
    auto* cell = sheet->getCell(row_t(1), column_t(1));
    ASSERT_NE(cell, nullptr);
    EXPECT_TRUE(cell->hasStyle());
    
    // 保存文件验证
    EXPECT_TRUE(workbook->saveToFile("test_styles.xlsx"));
}

// 测试边框样式
TEST_F(CellStyleTest, BorderStyling) {
    // 创建边框样式
    TXBorder border;
    border.setLeftBorder(BorderStyle::Thin, TXColor(0, 0, 255));   // 蓝色细线
    border.setRightBorder(BorderStyle::Thick, TXColor(0, 255, 0)); // 绿色粗线
    border.setTopBorder(BorderStyle::Double, TXColor(255, 0, 0));  // 红色双线
    border.setBottomBorder(BorderStyle::Dashed, TXColor(0, 0, 0)); // 黑色虚线
    
    // 创建单元格样式
    TXCellStyle style;
    style.setBorder(border);
    
    // 应用样式
    sheet->setCellValue(row_t(2), column_t(2), std::string("Bordered Cell"));
    EXPECT_TRUE(sheet->setCellStyle(row_t(2), column_t(2), style));
    
    // 验证样式
    auto* cell = sheet->getCell(row_t(2), column_t(2));
    ASSERT_NE(cell, nullptr);
    EXPECT_TRUE(cell->hasStyle());
    
    // 保存文件验证
    EXPECT_TRUE(workbook->saveToFile("test_border_styles.xlsx"));
}

// 测试填充样式
TEST_F(CellStyleTest, FillStyling) {
    // 创建填充样式
    TXFill fill;
    fill.setPattern(FillPattern::Solid);
    fill.setForegroundColor(TXColor(255, 255, 0)); // 黄色背景
    
    // 创建单元格样式
    TXCellStyle style;
    style.setFill(fill);
    
    // 应用样式
    sheet->setCellValue(row_t(3), column_t(3), std::string("Yellow Background"));
    EXPECT_TRUE(sheet->setCellStyle(row_t(3), column_t(3), style));
    
    // 验证样式
    auto* cell = sheet->getCell(row_t(3), column_t(3));
    ASSERT_NE(cell, nullptr);
    EXPECT_TRUE(cell->hasStyle());
    
    // 保存文件验证
    EXPECT_TRUE(workbook->saveToFile("test_fill_styles.xlsx"));
}

// 测试对齐样式
TEST_F(CellStyleTest, AlignmentStyling) {
    // 创建对齐样式
    TXAlignment alignment;
    alignment.setHorizontal(HorizontalAlignment::Center);
    alignment.setVertical(VerticalAlignment::Middle);
    alignment.setWrapText(true);
    
    // 创建单元格样式
    TXCellStyle style;
    style.setAlignment(alignment);
    
    // 应用样式
    sheet->setCellValue(row_t(4), column_t(4), std::string("Centered Text"));
    EXPECT_TRUE(sheet->setCellStyle(row_t(4), column_t(4), style));
    
    // 验证样式
    auto* cell = sheet->getCell(row_t(4), column_t(4));
    ASSERT_NE(cell, nullptr);
    EXPECT_TRUE(cell->hasStyle());
    
    // 保存文件验证
    EXPECT_TRUE(workbook->saveToFile("test_alignment_styles.xlsx"));
}

// 测试组合样式
TEST_F(CellStyleTest, CombinedStyling) {
    // 创建组合样式
    TXFont font;
    font.setName("Times New Roman");
    font.setSize(14);
    font.setBold(true);
    font.setColor(TXColor(255, 255, 255)); // 白色文字
    
    TXFill fill;
    fill.setPattern(FillPattern::Solid);
    fill.setForegroundColor(TXColor(0, 0, 128)); // 深蓝色背景
    
    TXBorder border;
    border.setAllBorders(BorderStyle::Medium, TXColor(0, 0, 0)); // 黑色中等边框
    
    TXAlignment alignment;
    alignment.setHorizontal(HorizontalAlignment::Center);
    alignment.setVertical(VerticalAlignment::Middle);
    
    // 组合样式
    TXCellStyle style;
    style.setFont(font);
    style.setFill(fill);
    style.setBorder(border);
    style.setAlignment(alignment);
    
    // 应用样式
    sheet->setCellValue(row_t(5), column_t(5), std::string("Styled Header"));
    EXPECT_TRUE(sheet->setCellStyle(row_t(5), column_t(5), style));
    
    // 验证样式
    auto* cell = sheet->getCell(row_t(5), column_t(5));
    ASSERT_NE(cell, nullptr);
    EXPECT_TRUE(cell->hasStyle());
    
    // 保存文件验证
    EXPECT_TRUE(workbook->saveToFile("test_combined_styles.xlsx"));
}

// 测试范围样式应用
TEST_F(CellStyleTest, RangeStyling) {
    // 创建简单样式
    TXFont font;
    font.setBold(true);
    font.setSize(10);
    
    TXCellStyle style;
    style.setFont(font);
    
    // 创建2x3范围
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(2), column_t(3)));
    
    // 填充数据
    for (row_t row = row_t(1); row <= row_t(2); ++row) {
        for (column_t col = column_t(1); col <= column_t(3); ++col) {
            sheet->setCellValue(row, col, std::string("Cell ") + 
                               std::to_string(row.index()) + "," + std::to_string(col.index()));
        }
    }
    
    // 应用范围样式
    std::size_t count = sheet->setRangeStyle(range, style);
    EXPECT_EQ(count, 6); // 2x3 = 6个单元格
    
    // 验证每个单元格都有样式
    for (row_t row = row_t(1); row <= row_t(2); ++row) {
        for (column_t col = column_t(1); col <= column_t(3); ++col) {
            auto* cell = sheet->getCell(row, col);
            ASSERT_NE(cell, nullptr);
            EXPECT_TRUE(cell->hasStyle());
        }
    }
}

// 测试批量样式应用
TEST_F(CellStyleTest, BatchStyling) {
    // 创建不同的样式
    TXCellStyle headerStyle;
    TXFont headerFont;
    headerFont.setBold(true);
    headerFont.setSize(12);
    headerStyle.setFont(headerFont);
    
    TXCellStyle dataStyle;
    TXFont dataFont;
    dataFont.setSize(10);
    dataStyle.setFont(dataFont);
    
    TXCellStyle totalStyle;
    TXFont totalFont;
    totalFont.setBold(true);
    totalFont.setItalic(true);
    totalStyle.setFont(totalFont);
    
    // 准备批量样式数据
    std::vector<std::pair<TXCoordinate, TXCellStyle>> styles = {
        {TXCoordinate(row_t(1), column_t(1)), headerStyle},  // 标题
        {TXCoordinate(row_t(1), column_t(2)), headerStyle},  // 标题
        {TXCoordinate(row_t(2), column_t(1)), dataStyle},    // 数据
        {TXCoordinate(row_t(2), column_t(2)), dataStyle},    // 数据
        {TXCoordinate(row_t(3), column_t(1)), totalStyle}    // 合计
    };
    
    // 设置单元格值
    sheet->setCellValue(row_t(1), column_t(1), std::string("Name"));
    sheet->setCellValue(row_t(1), column_t(2), std::string("Amount"));
    sheet->setCellValue(row_t(2), column_t(1), std::string("Item 1"));
    sheet->setCellValue(row_t(2), column_t(2), 100.0);
    sheet->setCellValue(row_t(3), column_t(1), std::string("Total"));
    
    // 批量应用样式
    std::size_t count = sheet->setCellStyles(styles);
    EXPECT_EQ(count, 5);
    
    // 验证样式应用
    for (const auto& styleEntry : styles) {
        auto* cell = sheet->getCell(styleEntry.first);
        ASSERT_NE(cell, nullptr);
        EXPECT_TRUE(cell->hasStyle());
    }
}

// 测试使用A1格式的样式应用
TEST_F(CellStyleTest, A1FormatStyling) {
    // 创建样式
    TXFont font;
    font.setName("Calibri");
    font.setSize(11);
    font.setColor(TXColor(128, 0, 128)); // 紫色
    
    TXCellStyle style;
    style.setFont(font);
    
    // 使用A1格式应用样式
    sheet->setCellValue("A1", std::string("A1 Cell"));
    sheet->setCellValue("B2", std::string("B2 Cell"));
    sheet->setCellValue("C3", std::string("C3 Cell"));
    
    EXPECT_TRUE(sheet->setCellStyle("A1", style));
    EXPECT_TRUE(sheet->setCellStyle("B2", style));
    EXPECT_TRUE(sheet->setCellStyle("C3", style));
    
    // 验证样式
    auto* cellA1 = sheet->getCell("A1");
    auto* cellB2 = sheet->getCell("B2");
    auto* cellC3 = sheet->getCell("C3");
    
    ASSERT_NE(cellA1, nullptr);
    ASSERT_NE(cellB2, nullptr);
    ASSERT_NE(cellC3, nullptr);
    
    EXPECT_TRUE(cellA1->hasStyle());
    EXPECT_TRUE(cellB2->hasStyle());
    EXPECT_TRUE(cellC3->hasStyle());
} 
