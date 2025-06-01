#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "test_file_generator.hpp"
#include <memory>
#include <iostream>

using namespace TinaXlsx;

class ColumnWidthDecimalTest : public TestWithFileGeneration<ColumnWidthDecimalTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<ColumnWidthDecimalTest>::SetUp();
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("小数列宽测试");
    }

    void TearDown() override {
        workbook.reset();
        TestWithFileGeneration<ColumnWidthDecimalTest>::TearDown();
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
};

TEST_F(ColumnWidthDecimalTest, DecimalColumnWidths) {
    std::cout << "\n=== 小数列宽测试 ===" << std::endl;
    
    // 测试各种小数列宽
    double widths[] = {8.5, 12.25, 15.75, 20.1, 25.99, 30.0, 35.123};
    
    for (size_t i = 0; i < sizeof(widths)/sizeof(widths[0]); ++i) {
        column_t col(i + 1);
        double width = widths[i];
        
        std::cout << "设置列" << (i + 1) << "宽度为: " << width << std::endl;
        bool result = sheet->setColumnWidth(col, width);
        EXPECT_TRUE(result);
        
        double actualWidth = sheet->getColumnWidth(col);
        std::cout << "实际获取的宽度: " << actualWidth << std::endl;
        EXPECT_DOUBLE_EQ(actualWidth, width);
    }
    
    // 生成测试文件
    addTestInfo(sheet, "DecimalColumnWidths", "测试小数列宽设置和精度保持");
    
    // 添加测试数据
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"列"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"设置宽度"});
    sheet->setCellValue(row_t(7), column_t(3), cell_value_t{"实际宽度"});
    sheet->setCellValue(row_t(7), column_t(4), cell_value_t{"内容示例"});
    
    for (size_t i = 0; i < sizeof(widths)/sizeof(widths[0]); ++i) {
        column_t col(i + 1);
        double width = widths[i];
        double actualWidth = sheet->getColumnWidth(col);
        
        row_t row(8 + i);
        sheet->setCellValue(row, column_t(1), cell_value_t{static_cast<char>('A' + i)});
        sheet->setCellValue(row, column_t(2), cell_value_t{width});
        sheet->setCellValue(row, column_t(3), cell_value_t{actualWidth});
        
        // 根据列宽添加不同长度的内容
        std::string content;
        if (width < 10) {
            content = "短";
        } else if (width < 20) {
            content = "中等长度内容";
        } else if (width < 30) {
            content = "这是较长的内容示例";
        } else {
            content = "这是一个很长的内容示例，用于测试宽列的显示效果";
        }
        sheet->setCellValue(row, column_t(4), cell_value_t{content});
    }
    
    // 添加说明
    sheet->setCellValue(row_t(16), column_t(1), cell_value_t{"测试说明:"});
    sheet->setCellValue(row_t(17), column_t(1), cell_value_t{"1. 测试各种小数列宽设置"});
    sheet->setCellValue(row_t(18), column_t(1), cell_value_t{"2. 验证小数精度是否正确保持"});
    sheet->setCellValue(row_t(19), column_t(1), cell_value_t{"3. 检查XML中的列宽格式"});
    
    saveWorkbook(workbook, "DecimalColumnWidths");
    
    std::cout << "=== 小数列宽测试完成 ===" << std::endl;
}
