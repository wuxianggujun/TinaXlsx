#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <cstdio>

using namespace TinaXlsx;
using cell_value_t = TinaXlsx::cell_value_t;

class NewArchitectureTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("TestSheet");
        ASSERT_NE(sheet, nullptr);
    }

    void TearDown() override {
        workbook.reset();
        std::remove("test_new_architecture.xlsx");
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet;
};

// 测试新的数字格式架构
TEST_F(NewArchitectureTest, NumberFormatArchitecture) {
    // 设置一些数据
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{1234.567});
    sheet->setCellValue(row_t(2), column_t(1), cell_value_t{0.75});
    sheet->setCellValue(row_t(3), column_t(1), cell_value_t{50000.0});

    // 使用新架构设置数字格式
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(1), column_t(1), TXNumberFormat::FormatType::Number, 2));
    EXPECT_TRUE(sheet->setCellNumberFormat(row_t(2), column_t(1), TXNumberFormat::FormatType::Percentage, 1));
    EXPECT_TRUE(sheet->setCellCustomFormat(row_t(3), column_t(1), "#,##0.00_);[红色](#,##0.00)"));

    // 验证格式化值
    std::string formatted1 = sheet->getCellFormattedValue(row_t(1), column_t(1));
    std::string formatted2 = sheet->getCellFormattedValue(row_t(2), column_t(1));
    std::string formatted3 = sheet->getCellFormattedValue(row_t(3), column_t(1));

    EXPECT_FALSE(formatted1.empty());
    EXPECT_FALSE(formatted2.empty());
    EXPECT_FALSE(formatted3.empty());

    // 保存文件验证
    EXPECT_TRUE(workbook->saveToFile("test_new_architecture.xlsx"));
}

// 测试完整样式设置
TEST_F(NewArchitectureTest, FullStyleArchitecture) {
    // 创建完整样式
    TXCellStyle style;
    style.setFontName("Arial")
         .setFontSize(14)
         .setFontBold(true)
         .setNumberFormat(TXNumberFormat::FormatType::Currency, 2, true, "¥")
         .setHorizontalAlignment(HorizontalAlignment::Center);

    // 设置数据和样式
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{12345.67});
    EXPECT_TRUE(sheet->setCellStyle(row_t(1), column_t(1), style));

    // 验证保存
    EXPECT_TRUE(workbook->saveToFile("test_new_architecture.xlsx"));
}

// 测试NumberFormatDefinition
TEST_F(NewArchitectureTest, NumberFormatDefinition) {
    // 测试基本格式定义
    TXCellStyle::NumberFormatDefinition def1(TXNumberFormat::FormatType::Number, 3, true);
    EXPECT_FALSE(def1.isGeneral());
    
    std::string code1 = def1.generateExcelFormatCode();
    EXPECT_FALSE(code1.empty());
    
    // 测试自定义格式定义
    TXCellStyle::NumberFormatDefinition def2("0.000%");
    EXPECT_FALSE(def2.isGeneral());
    EXPECT_EQ(def2.generateExcelFormatCode(), "0.000%");
    
    // 测试常规格式
    TXCellStyle::NumberFormatDefinition def3;
    EXPECT_TRUE(def3.isGeneral());
    EXPECT_EQ(def3.generateExcelFormatCode(), "General");
}
