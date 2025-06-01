#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "test_file_generator.hpp"
#include <memory>

using namespace TinaXlsx;

// 基本功能测试
class BasicFeaturesTest : public TestWithFileGeneration<BasicFeaturesTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<BasicFeaturesTest>::SetUp();
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("测试工作表");
    }

    void TearDown() override {
        workbook.reset();
        TestWithFileGeneration<BasicFeaturesTest>::TearDown();
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
};

TEST_F(BasicFeaturesTest, CreateWorkbookAndSheet) {
    ASSERT_NE(workbook, nullptr);
    ASSERT_NE(sheet, nullptr);
    EXPECT_EQ(sheet->getName(), "测试工作表");

    // 生成测试文件
    addTestInfo(sheet, "CreateWorkbookAndSheet", "测试工作簿和工作表的创建功能");
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"工作簿名称:"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"TinaXlsx测试工作簿"});
    sheet->setCellValue(row_t(8), column_t(1), cell_value_t{"工作表名称:"});
    sheet->setCellValue(row_t(8), column_t(2), cell_value_t{sheet->getName()});

    saveWorkbook(workbook, "01_CreateWorkbookAndSheet");
}

TEST_F(BasicFeaturesTest, SetAndGetCellValue) {
    // 测试设置和获取单元格值
    EXPECT_TRUE(sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"测试数据"}));

    auto value = sheet->getCellValue(row_t(1), column_t(1));
    EXPECT_TRUE(std::holds_alternative<std::string>(value));
    EXPECT_EQ(std::get<std::string>(value), "测试数据");

    // 生成测试文件
    addTestInfo(sheet, "SetAndGetCellValue", "测试单元格数据的设置和获取功能");

    // 添加各种数据类型的示例
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"数据类型"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"示例值"});

    sheet->setCellValue(row_t(8), column_t(1), cell_value_t{"字符串"});
    sheet->setCellValue(row_t(8), column_t(2), cell_value_t{"Hello TinaXlsx"});

    sheet->setCellValue(row_t(9), column_t(1), cell_value_t{"整数"});
    sheet->setCellValue(row_t(9), column_t(2), cell_value_t{42});

    sheet->setCellValue(row_t(10), column_t(1), cell_value_t{"浮点数"});
    sheet->setCellValue(row_t(10), column_t(2), cell_value_t{3.14159});

    sheet->setCellValue(row_t(11), column_t(1), cell_value_t{"布尔值"});
    sheet->setCellValue(row_t(11), column_t(2), cell_value_t{true});

    sheet->setCellValue(row_t(12), column_t(1), cell_value_t{"中文测试"});
    sheet->setCellValue(row_t(12), column_t(2), cell_value_t{"这是中文内容测试"});

    saveWorkbook(workbook, "02_SetAndGetCellValue");
}

TEST_F(BasicFeaturesTest, ColumnWidthBasic) {
    // 测试基本的列宽设置
    EXPECT_TRUE(sheet->setColumnWidth(column_t(1), 15.0));
    EXPECT_DOUBLE_EQ(sheet->getColumnWidth(column_t(1)), 15.0);

    // 测试默认列宽
    EXPECT_DOUBLE_EQ(sheet->getColumnWidth(column_t(2)), 8.43);

    // 生成测试文件
    addTestInfo(sheet, "ColumnWidthBasic", "测试列宽设置和获取功能");

    // 设置不同的列宽并添加内容
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"列"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"宽度"});
    sheet->setCellValue(row_t(7), column_t(3), cell_value_t{"内容示例"});

    sheet->setColumnWidth(column_t(1), 8.0);
    sheet->setCellValue(row_t(8), column_t(1), cell_value_t{"A"});
    sheet->setCellValue(row_t(8), column_t(2), cell_value_t{"8.0"});
    sheet->setCellValue(row_t(8), column_t(3), cell_value_t{"窄列"});

    sheet->setColumnWidth(column_t(2), 15.0);
    sheet->setCellValue(row_t(9), column_t(1), cell_value_t{"B"});
    sheet->setCellValue(row_t(9), column_t(2), cell_value_t{"15.0"});
    sheet->setCellValue(row_t(9), column_t(3), cell_value_t{"中等宽度列"});

    sheet->setColumnWidth(column_t(3), 25.0);
    sheet->setCellValue(row_t(10), column_t(1), cell_value_t{"C"});
    sheet->setCellValue(row_t(10), column_t(2), cell_value_t{"25.0"});
    sheet->setCellValue(row_t(10), column_t(3), cell_value_t{"这是一个比较宽的列，用于显示更多内容"});

    saveWorkbook(workbook, "03_ColumnWidthBasic");
}

TEST_F(BasicFeaturesTest, RowHeightBasic) {
    // 测试基本的行高设置
    EXPECT_TRUE(sheet->setRowHeight(row_t(1), 20.0));
    EXPECT_DOUBLE_EQ(sheet->getRowHeight(row_t(1)), 20.0);

    // 测试默认行高
    EXPECT_DOUBLE_EQ(sheet->getRowHeight(row_t(2)), 15.0);

    // 生成测试文件
    addTestInfo(sheet, "RowHeightBasic", "测试行高设置和获取功能");

    // 设置不同的行高并添加内容
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"行号"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"高度"});
    sheet->setCellValue(row_t(7), column_t(3), cell_value_t{"内容示例"});

    sheet->setRowHeight(row_t(8), 15.0);
    sheet->setCellValue(row_t(8), column_t(1), cell_value_t{"8"});
    sheet->setCellValue(row_t(8), column_t(2), cell_value_t{"15.0"});
    sheet->setCellValue(row_t(8), column_t(3), cell_value_t{"标准高度行"});

    sheet->setRowHeight(row_t(9), 25.0);
    sheet->setCellValue(row_t(9), column_t(1), cell_value_t{"9"});
    sheet->setCellValue(row_t(9), column_t(2), cell_value_t{"25.0"});
    sheet->setCellValue(row_t(9), column_t(3), cell_value_t{"较高的行"});

    sheet->setRowHeight(row_t(10), 35.0);
    sheet->setCellValue(row_t(10), column_t(1), cell_value_t{"10"});
    sheet->setCellValue(row_t(10), column_t(2), cell_value_t{"35.0"});
    sheet->setCellValue(row_t(10), column_t(3), cell_value_t{"很高的行，可以容纳更多内容"});

    saveWorkbook(workbook, "04_RowHeightBasic");
}

TEST_F(BasicFeaturesTest, CellLockingBasic) {
    // 创建单元格
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"测试数据"});

    // 测试锁定功能
    EXPECT_TRUE(sheet->setCellLocked(row_t(1), column_t(1), false));
    EXPECT_FALSE(sheet->isCellLocked(row_t(1), column_t(1)));

    EXPECT_TRUE(sheet->setCellLocked(row_t(1), column_t(1), true));
    EXPECT_TRUE(sheet->isCellLocked(row_t(1), column_t(1)));

    // 生成测试文件
    addTestInfo(sheet, "CellLockingBasic", "测试单元格锁定功能");

    // 创建锁定和未锁定的单元格示例
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"单元格"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"锁定状态"});
    sheet->setCellValue(row_t(7), column_t(3), cell_value_t{"内容"});

    sheet->setCellValue(row_t(8), column_t(1), cell_value_t{"A8"});
    sheet->setCellValue(row_t(8), column_t(2), cell_value_t{"已锁定"});
    sheet->setCellValue(row_t(8), column_t(3), cell_value_t{"这个单元格已锁定"});
    sheet->setCellLocked(row_t(8), column_t(3), true);

    sheet->setCellValue(row_t(9), column_t(1), cell_value_t{"A9"});
    sheet->setCellValue(row_t(9), column_t(2), cell_value_t{"未锁定"});
    sheet->setCellValue(row_t(9), column_t(3), cell_value_t{"这个单元格未锁定"});
    sheet->setCellLocked(row_t(9), column_t(3), false);

    saveWorkbook(workbook, "05_CellLockingBasic");
}

TEST_F(BasicFeaturesTest, SheetProtectionBasic) {
    // 测试基本保护功能
    EXPECT_FALSE(sheet->isSheetProtected());

    EXPECT_TRUE(sheet->protectSheet("test123"));
    EXPECT_TRUE(sheet->isSheetProtected());

    EXPECT_TRUE(sheet->unprotectSheet("test123"));
    EXPECT_FALSE(sheet->isSheetProtected());

    // 生成测试文件
    addTestInfo(sheet, "SheetProtectionBasic", "测试工作表保护功能");

    // 添加保护相关的信息
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"保护功能"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"状态"});

    sheet->setCellValue(row_t(8), column_t(1), cell_value_t{"工作表保护"});
    sheet->setCellValue(row_t(8), column_t(2), cell_value_t{"已测试"});

    sheet->setCellValue(row_t(9), column_t(1), cell_value_t{"密码保护"});
    sheet->setCellValue(row_t(9), column_t(2), cell_value_t{"支持"});

    sheet->setCellValue(row_t(10), column_t(1), cell_value_t{"测试密码"});
    sheet->setCellValue(row_t(10), column_t(2), cell_value_t{"test123"});

    // 重新保护工作表用于演示
    sheet->protectSheet("test123");

    saveWorkbook(workbook, "06_SheetProtectionBasic");
}