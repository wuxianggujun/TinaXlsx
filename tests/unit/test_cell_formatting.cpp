#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <cstdio>

using namespace TinaXlsx;

class CellFormattingTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("FormatTest");
        ASSERT_NE(sheet, nullptr);
    }

    void TearDown() override {
        workbook.reset();
        // 清理测试文件
        // std::remove("test_formatting.xlsx");
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet;
};

// 测试数字格式化
TEST_F(CellFormattingTest, NumberFormatting) {
    // 设置不同的数字值
    sheet->setCellValue(row_t(1), column_t(1), 1234.567);
    sheet->setCellValue(row_t(2), column_t(1), -9876.543);
    sheet->setCellValue(row_t(3), column_t(1), 0.75);
    
    // 应用数字格式 - 2位小数
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(1), column_t(1), TXCell::NumberFormat::Number, 2));
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(2), column_t(1), TXCell::NumberFormat::Number, 2));
    
    // 检查格式化后的值
    std::string formatted1 = sheet->getCellFormattedValue(row_t(1), column_t(1));
    std::string formatted2 = sheet->getCellFormattedValue(row_t(2), column_t(1));
    
    EXPECT_FALSE(formatted1.empty());
    EXPECT_FALSE(formatted2.empty());
    
    // 保存验证
    EXPECT_TRUE(workbook->saveToFile("test_formatting.xlsx"));
}

// 测试百分比格式化
TEST_F(CellFormattingTest, PercentageFormatting) {
    // 设置百分比值
    sheet->setCellValue(row_t(1), column_t(1), 0.25);  // 25%
    sheet->setCellValue(row_t(2), column_t(1), 0.856); // 85.6%
    sheet->setCellValue(row_t(3), column_t(1), 1.2);   // 120%
    
    // 应用百分比格式
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(1), column_t(1), TXCell::NumberFormat::Percentage, 1));
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(2), column_t(1), TXCell::NumberFormat::Percentage, 1));
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(3), column_t(1), TXCell::NumberFormat::Percentage, 0));
    
    // 检查格式化后的值
    std::string formatted1 = sheet->getCellFormattedValue(row_t(1), column_t(1));
    std::string formatted2 = sheet->getCellFormattedValue(row_t(2), column_t(1));
    std::string formatted3 = sheet->getCellFormattedValue(row_t(3), column_t(1));
    
    EXPECT_FALSE(formatted1.empty());
    EXPECT_FALSE(formatted2.empty());  
    EXPECT_FALSE(formatted3.empty());
}

// 测试货币格式化
TEST_F(CellFormattingTest, CurrencyFormatting) {
    // 设置货币值
    sheet->setCellValue(row_t(1), column_t(1), 1234.56);
    sheet->setCellValue(row_t(2), column_t(1), -567.89);
    
    // 应用货币格式
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(1), column_t(1), TXCell::NumberFormat::Currency, 2));
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(2), column_t(1), TXCell::NumberFormat::Currency, 2));
    
    // 检查格式化后的值
    std::string formatted1 = sheet->getCellFormattedValue(row_t(1), column_t(1));
    std::string formatted2 = sheet->getCellFormattedValue(row_t(2), column_t(1));
    
    EXPECT_FALSE(formatted1.empty());
    EXPECT_FALSE(formatted2.empty());
}

// 测试自定义格式
TEST_F(CellFormattingTest, CustomFormatting) {
    // 设置数值
    sheet->setCellValue(row_t(1), column_t(1), 1234567.89);
    
    // 应用自定义格式
    EXPECT_TRUE(sheet->setCellCustomFormat(row_t(1), column_t(1), "#,##0.00"));
    
    // 检查自定义格式
    auto* cell = sheet->getCell(row_t(1), column_t(1));
    ASSERT_NE(cell, nullptr);
    EXPECT_EQ(cell->getCustomFormat(), "#,##0.00");
}

// 测试范围格式化
TEST_F(CellFormattingTest, RangeFormatting) {
    // 设置一个2x2的数据范围
    sheet->setCellValue(row_t(1), column_t(1), 100.1);
    sheet->setCellValue(row_t(1), column_t(2), 200.2);
    sheet->setCellValue(row_t(2), column_t(1), 300.3);
    sheet->setCellValue(row_t(2), column_t(2), 400.4);
    
    // 创建范围
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(2), column_t(2)));
    
    // 应用范围格式
    std::size_t count = sheet->setRangeNumberFormat(range, TXCell::NumberFormat::Number, 1);
    EXPECT_EQ(count, 4);  // 应该格式化4个单元格
    
    // 验证每个单元格都被格式化了
    for (row_t row = row_t(1); row <= row_t(2); ++row) {
        for (column_t col = column_t(1); col <= column_t(2); ++col) {
            auto* cell = sheet->getCell(row, col);
            ASSERT_NE(cell, nullptr);
            EXPECT_EQ(cell->getNumberFormat(), TXCell::NumberFormat::Number);
        }
    }
}

// 测试TXNumberFormat对象
TEST_F(CellFormattingTest, TXNumberFormatObject) {
    // 创建不同类型的格式对象
    auto numberFormat = TXNumberFormat::createNumberFormat(3, true);
    auto currencyFormat = TXNumberFormat::createCurrencyFormat("¥", 2);
    auto percentageFormat = TXNumberFormat::createPercentageFormat(1);
    
    // 验证格式类型
    EXPECT_EQ(numberFormat.getFormatType(), TXNumberFormat::FormatType::Number);
    EXPECT_EQ(currencyFormat.getFormatType(), TXNumberFormat::FormatType::Currency);
    EXPECT_EQ(percentageFormat.getFormatType(), TXNumberFormat::FormatType::Percentage);
    
    // 测试格式化功能
    double testValue = 1234.5678;
    
    std::string numberStr = numberFormat.formatNumber(testValue);
    std::string currencyStr = currencyFormat.formatCurrency(testValue);
    std::string percentStr = percentageFormat.formatPercentage(testValue);
    
    EXPECT_FALSE(numberStr.empty());
    EXPECT_FALSE(currencyStr.empty());
    EXPECT_FALSE(percentStr.empty());
    
    // 验证货币符号
    EXPECT_TRUE(currencyStr.find("¥") != std::string::npos);
}

// 测试批量格式化
TEST_F(CellFormattingTest, BatchFormatting) {
    // 设置多个单元格的值
    std::vector<std::pair<TXCoordinate, TXCell::NumberFormat>> formats = {
        {TXCoordinate(row_t(1), column_t(1)), TXCell::NumberFormat::Number},
        {TXCoordinate(row_t(1), column_t(2)), TXCell::NumberFormat::Currency}, 
        {TXCoordinate(row_t(1), column_t(3)), TXCell::NumberFormat::Percentage}
    };
    
    // 先设置数值
    sheet->setCellValue(row_t(1), column_t(1), 123.45);
    sheet->setCellValue(row_t(1), column_t(2), 678.90);
    sheet->setCellValue(row_t(1), column_t(3), 0.85);
    
    // 批量应用格式
    std::size_t count = sheet->setCellFormats(formats);
    EXPECT_EQ(count, 3);
    
    // 验证格式是否正确应用
    auto* cell1 = sheet->getCell(row_t(1), column_t(1));
    auto* cell2 = sheet->getCell(row_t(1), column_t(2));
    auto* cell3 = sheet->getCell(row_t(1), column_t(3));
    
    ASSERT_NE(cell1, nullptr);
    ASSERT_NE(cell2, nullptr);
    ASSERT_NE(cell3, nullptr);
    
    EXPECT_EQ(cell1->getNumberFormat(), TXCell::NumberFormat::Number);
    EXPECT_EQ(cell2->getNumberFormat(), TXCell::NumberFormat::Currency);
    EXPECT_EQ(cell3->getNumberFormat(), TXCell::NumberFormat::Percentage);
} 
