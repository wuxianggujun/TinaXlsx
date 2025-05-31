#include <gtest/gtest.h>
#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <memory>

using namespace TinaXlsx;

class ColumnWidthRowHeightTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("测试工作表");
    }

    void TearDown() override {
        workbook.reset();
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
};

TEST_F(ColumnWidthRowHeightTest, SetAndGetColumnWidth) {
    // 测试设置列宽
    EXPECT_TRUE(sheet->setColumnWidth(column_t(1), 15.0));
    EXPECT_DOUBLE_EQ(sheet->getColumnWidth(column_t(1)), 15.0);
    
    // 测试默认列宽
    EXPECT_DOUBLE_EQ(sheet->getColumnWidth(column_t(2)), 8.43);
    
    // 测试边界值
    EXPECT_TRUE(sheet->setColumnWidth(column_t(1), 1.0));   // 最小值
    EXPECT_TRUE(sheet->setColumnWidth(column_t(1), 255.0)); // 最大值
    
    // 测试无效值
    EXPECT_FALSE(sheet->setColumnWidth(column_t(1), -1.0));  // 负值
    EXPECT_FALSE(sheet->setColumnWidth(column_t(1), 256.0)); // 超出最大值
    EXPECT_FALSE(sheet->setColumnWidth(column_t(static_cast<u32>(0)), 10.0));  // 无效列号
}

TEST_F(ColumnWidthRowHeightTest, SetAndGetRowHeight) {
    // 测试设置行高
    EXPECT_TRUE(sheet->setRowHeight(row_t(1), 20.0));
    EXPECT_DOUBLE_EQ(sheet->getRowHeight(row_t(1)), 20.0);
    
    // 测试默认行高
    EXPECT_DOUBLE_EQ(sheet->getRowHeight(row_t(2)), 15.0);
    
    // 测试边界值
    EXPECT_TRUE(sheet->setRowHeight(row_t(1), 1.0));   // 最小值
    EXPECT_TRUE(sheet->setRowHeight(row_t(1), 409.0)); // 最大值
    
    // 测试无效值
    EXPECT_FALSE(sheet->setRowHeight(row_t(1), -1.0));  // 负值
    EXPECT_FALSE(sheet->setRowHeight(row_t(1), 410.0)); // 超出最大值
    EXPECT_FALSE(sheet->setRowHeight(row_t(0), 10.0));  // 无效行号
}

TEST_F(ColumnWidthRowHeightTest, AutoFitColumnWidth) {
    // 添加测试数据
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"短文本"});
    sheet->setCellValue(row_t(2), column_t(1), cell_value_t{"这是一个很长的文本内容用于测试自动调整列宽功能"});
    sheet->setCellValue(row_t(3), column_t(1), cell_value_t{"中等长度的文本内容"});
    
    // 测试自动调整列宽
    double originalWidth = sheet->getColumnWidth(column_t(1));
    double newWidth = sheet->autoFitColumnWidth(column_t(1));
    
    EXPECT_GT(newWidth, originalWidth); // 新宽度应该大于原宽度
    EXPECT_GE(newWidth, 1.0);          // 不小于最小宽度
    EXPECT_LE(newWidth, 255.0);        // 不大于最大宽度
    
    // 验证宽度已被设置
    EXPECT_DOUBLE_EQ(sheet->getColumnWidth(column_t(1)), newWidth);
}

TEST_F(ColumnWidthRowHeightTest, AutoFitRowHeight) {
    // 添加测试数据
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"单行文本"});
    
    // 测试自动调整行高
    double originalHeight = sheet->getRowHeight(row_t(1));
    double newHeight = sheet->autoFitRowHeight(row_t(1));
    
    EXPECT_GE(newHeight, 12.0);   // 不小于最小高度
    EXPECT_LE(newHeight, 409.0);  // 不大于最大高度
    
    // 验证高度已被设置
    EXPECT_DOUBLE_EQ(sheet->getRowHeight(row_t(1)), newHeight);
}

TEST_F(ColumnWidthRowHeightTest, AutoFitAllColumns) {
    // 添加多列数据
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"列A数据"});
    sheet->setCellValue(row_t(1), column_t(2), cell_value_t{"列B的长数据内容"});
    sheet->setCellValue(row_t(1), column_t(3), cell_value_t{"列C"});
    
    // 测试自动调整所有列宽
    std::size_t adjustedCount = sheet->autoFitAllColumnWidths();
    EXPECT_EQ(adjustedCount, 3); // 应该调整了3列
    
    // 验证各列宽度都被调整
    EXPECT_GT(sheet->getColumnWidth(column_t(1)), 8.43);
    EXPECT_GT(sheet->getColumnWidth(column_t(2)), 8.43);
    EXPECT_GT(sheet->getColumnWidth(column_t(3)), 8.43);
}

TEST_F(ColumnWidthRowHeightTest, AutoFitAllRows) {
    // 添加多行数据
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"行1数据"});
    sheet->setCellValue(row_t(2), column_t(1), cell_value_t{"行2数据"});
    sheet->setCellValue(row_t(3), column_t(1), cell_value_t{"行3数据"});
    
    // 测试自动调整所有行高
    std::size_t adjustedCount = sheet->autoFitAllRowHeights();
    EXPECT_EQ(adjustedCount, 3); // 应该调整了3行
}

TEST_F(ColumnWidthRowHeightTest, AutoFitWithCustomLimits) {
    // 添加测试数据
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"测试数据"});
    
    // 测试自定义最小/最大宽度
    double width = sheet->autoFitColumnWidth(column_t(1), 5.0, 10.0);
    EXPECT_GE(width, 5.0);
    EXPECT_LE(width, 10.0);
    
    // 测试自定义最小/最大高度
    double height = sheet->autoFitRowHeight(row_t(1), 20.0, 30.0);
    EXPECT_GE(height, 20.0);
    EXPECT_LE(height, 30.0);
}
