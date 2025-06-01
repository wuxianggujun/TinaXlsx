#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "test_file_generator.hpp"
#include <memory>
#include <iostream>

using namespace TinaXlsx;

class ColumnWidthDebugTest : public TestWithFileGeneration<ColumnWidthDebugTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<ColumnWidthDebugTest>::SetUp();
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("列宽调试测试");
    }

    void TearDown() override {
        workbook.reset();
        TestWithFileGeneration<ColumnWidthDebugTest>::TearDown();
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
};

TEST_F(ColumnWidthDebugTest, DebugColumnWidthSetting) {
    std::cout << "\n=== 列宽设置调试测试 ===" << std::endl;
    
    // 检查初始默认列宽
    double defaultWidth1 = sheet->getColumnWidth(column_t(1));
    double defaultWidth2 = sheet->getColumnWidth(column_t(2));
    double defaultWidth3 = sheet->getColumnWidth(column_t(3));
    
    std::cout << "初始列宽:" << std::endl;
    std::cout << "  A列: " << defaultWidth1 << std::endl;
    std::cout << "  B列: " << defaultWidth2 << std::endl;
    std::cout << "  C列: " << defaultWidth3 << std::endl;
    
    // 设置列宽
    std::cout << "\n设置列宽..." << std::endl;
    bool result1 = sheet->setColumnWidth(column_t(1), 8.0);
    bool result2 = sheet->setColumnWidth(column_t(2), 15.0);
    bool result3 = sheet->setColumnWidth(column_t(3), 25.0);
    
    std::cout << "设置结果:" << std::endl;
    std::cout << "  A列设置为8.0: " << (result1 ? "成功" : "失败") << std::endl;
    std::cout << "  B列设置为15.0: " << (result2 ? "成功" : "失败") << std::endl;
    std::cout << "  C列设置为25.0: " << (result3 ? "成功" : "失败") << std::endl;
    
    // 验证设置后的列宽
    double newWidth1 = sheet->getColumnWidth(column_t(1));
    double newWidth2 = sheet->getColumnWidth(column_t(2));
    double newWidth3 = sheet->getColumnWidth(column_t(3));
    
    std::cout << "\n设置后的列宽:" << std::endl;
    std::cout << "  A列: " << newWidth1 << " (期望: 8.0)" << std::endl;
    std::cout << "  B列: " << newWidth2 << " (期望: 15.0)" << std::endl;
    std::cout << "  C列: " << newWidth3 << " (期望: 25.0)" << std::endl;
    
    // 断言验证
    EXPECT_TRUE(result1);
    EXPECT_TRUE(result2);
    EXPECT_TRUE(result3);
    EXPECT_DOUBLE_EQ(newWidth1, 8.0);
    EXPECT_DOUBLE_EQ(newWidth2, 15.0);
    EXPECT_DOUBLE_EQ(newWidth3, 25.0);
    
    // 添加一些内容
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"A列内容"});
    sheet->setCellValue(row_t(1), column_t(2), cell_value_t{"B列内容"});
    sheet->setCellValue(row_t(1), column_t(3), cell_value_t{"C列内容"});
    
    // 再次检查列宽（确保添加内容后没有自动调整）
    double finalWidth1 = sheet->getColumnWidth(column_t(1));
    double finalWidth2 = sheet->getColumnWidth(column_t(2));
    double finalWidth3 = sheet->getColumnWidth(column_t(3));
    
    std::cout << "\n添加内容后的列宽:" << std::endl;
    std::cout << "  A列: " << finalWidth1 << " (期望: 8.0)" << std::endl;
    std::cout << "  B列: " << finalWidth2 << " (期望: 15.0)" << std::endl;
    std::cout << "  C列: " << finalWidth3 << " (期望: 25.0)" << std::endl;
    
    EXPECT_DOUBLE_EQ(finalWidth1, 8.0);
    EXPECT_DOUBLE_EQ(finalWidth2, 15.0);
    EXPECT_DOUBLE_EQ(finalWidth3, 25.0);
    
    // 生成测试文件
    addTestInfo(sheet, "DebugColumnWidthSetting", "调试列宽设置功能");
    
    // 添加调试信息到文件
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"列"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"设置宽度"});
    sheet->setCellValue(row_t(7), column_t(3), cell_value_t{"实际宽度"});
    sheet->setCellValue(row_t(7), column_t(4), cell_value_t{"状态"});
    
    sheet->setCellValue(row_t(8), column_t(1), cell_value_t{"A"});
    sheet->setCellValue(row_t(8), column_t(2), cell_value_t{8.0});
    sheet->setCellValue(row_t(8), column_t(3), cell_value_t{finalWidth1});
    sheet->setCellValue(row_t(8), column_t(4), cell_value_t{finalWidth1 == 8.0 ? "正确" : "错误"});
    
    sheet->setCellValue(row_t(9), column_t(1), cell_value_t{"B"});
    sheet->setCellValue(row_t(9), column_t(2), cell_value_t{15.0});
    sheet->setCellValue(row_t(9), column_t(3), cell_value_t{finalWidth2});
    sheet->setCellValue(row_t(9), column_t(4), cell_value_t{finalWidth2 == 15.0 ? "正确" : "错误"});
    
    sheet->setCellValue(row_t(10), column_t(1), cell_value_t{"C"});
    sheet->setCellValue(row_t(10), column_t(2), cell_value_t{25.0});
    sheet->setCellValue(row_t(10), column_t(3), cell_value_t{finalWidth3});
    sheet->setCellValue(row_t(10), column_t(4), cell_value_t{finalWidth3 == 25.0 ? "正确" : "错误"});
    
    // 添加一些测试内容来验证列宽效果
    sheet->setCellValue(row_t(12), column_t(1), cell_value_t{"短"});
    sheet->setCellValue(row_t(12), column_t(2), cell_value_t{"中等长度内容"});
    sheet->setCellValue(row_t(12), column_t(3), cell_value_t{"这是一个很长的内容，用于测试宽列的显示效果"});
    
    saveWorkbook(workbook, "DebugColumnWidthSetting");
    
    std::cout << "\n=== 调试测试完成 ===" << std::endl;
}
