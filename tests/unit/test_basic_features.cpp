#include <gtest/gtest.h>
#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <memory>

using namespace TinaXlsx;

// 基本功能测试
class BasicFeaturesTest : public ::testing::Test {
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

TEST_F(BasicFeaturesTest, CreateWorkbookAndSheet) {
    ASSERT_NE(workbook, nullptr);
    ASSERT_NE(sheet, nullptr);
    EXPECT_EQ(sheet->getName(), "测试工作表");
}

TEST_F(BasicFeaturesTest, SetAndGetCellValue) {
    // 测试设置和获取单元格值
    EXPECT_TRUE(sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"测试数据"}));
    
    auto value = sheet->getCellValue(row_t(1), column_t(1));
    EXPECT_TRUE(std::holds_alternative<std::string>(value));
    EXPECT_EQ(std::get<std::string>(value), "测试数据");
}

TEST_F(BasicFeaturesTest, ColumnWidthBasic) {
    // 测试基本的列宽设置
    EXPECT_TRUE(sheet->setColumnWidth(column_t(1), 15.0));
    EXPECT_DOUBLE_EQ(sheet->getColumnWidth(column_t(1)), 15.0);
    
    // 测试默认列宽
    EXPECT_DOUBLE_EQ(sheet->getColumnWidth(column_t(2)), 8.43);
}

TEST_F(BasicFeaturesTest, RowHeightBasic) {
    // 测试基本的行高设置
    EXPECT_TRUE(sheet->setRowHeight(row_t(1), 20.0));
    EXPECT_DOUBLE_EQ(sheet->getRowHeight(row_t(1)), 20.0);
    
    // 测试默认行高
    EXPECT_DOUBLE_EQ(sheet->getRowHeight(row_t(2)), 15.0);
}

TEST_F(BasicFeaturesTest, CellLockingBasic) {
    // 创建单元格
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"测试数据"});
    
    // 测试锁定功能
    EXPECT_TRUE(sheet->setCellLocked(row_t(1), column_t(1), false));
    EXPECT_FALSE(sheet->isCellLocked(row_t(1), column_t(1)));
    
    EXPECT_TRUE(sheet->setCellLocked(row_t(1), column_t(1), true));
    EXPECT_TRUE(sheet->isCellLocked(row_t(1), column_t(1)));
}

TEST_F(BasicFeaturesTest, SheetProtectionBasic) {
    // 测试基本保护功能
    EXPECT_FALSE(sheet->isSheetProtected());
    
    EXPECT_TRUE(sheet->protectSheet("test123"));
    EXPECT_TRUE(sheet->isSheetProtected());
    
    EXPECT_TRUE(sheet->unprotectSheet("test123"));
    EXPECT_FALSE(sheet->isSheetProtected());
}