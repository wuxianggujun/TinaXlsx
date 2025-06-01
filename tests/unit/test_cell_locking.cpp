#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "test_file_generator.hpp"
#include <memory>

using namespace TinaXlsx;

class CellLockingTest : public TestWithFileGeneration<CellLockingTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<CellLockingTest>::SetUp();
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("锁定测试");
    }

    void TearDown() override {
        workbook.reset();
        TestWithFileGeneration<CellLockingTest>::TearDown();
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
};

TEST_F(CellLockingTest, DefaultLockingState) {
    // 创建新单元格
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"测试数据"});
    
    TXCell* cell = sheet->getCell(row_t(1), column_t(1));
    ASSERT_NE(cell, nullptr);
    
    // 新单元格默认应该是锁定的
    EXPECT_TRUE(cell->isLocked());
    
    // 通过工作表接口也应该返回锁定状态
    EXPECT_TRUE(sheet->isCellLocked(row_t(1), column_t(1)));
}

TEST_F(CellLockingTest, SetCellLocking) {
    // 创建单元格
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"测试数据"});
    TXCell* cell = sheet->getCell(row_t(1), column_t(1));
    ASSERT_NE(cell, nullptr);
    
    // 测试解锁
    cell->setLocked(false);
    EXPECT_FALSE(cell->isLocked());
    EXPECT_FALSE(sheet->isCellLocked(row_t(1), column_t(1)));
    
    // 测试重新锁定
    cell->setLocked(true);
    EXPECT_TRUE(cell->isLocked());
    EXPECT_TRUE(sheet->isCellLocked(row_t(1), column_t(1)));
}

TEST_F(CellLockingTest, SetCellLockingViaSheet) {
    // 通过工作表接口设置锁定状态
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"测试数据"});
    
    // 测试解锁
    EXPECT_TRUE(sheet->setCellLocked(row_t(1), column_t(1), false));
    EXPECT_FALSE(sheet->isCellLocked(row_t(1), column_t(1)));
    
    // 测试重新锁定
    EXPECT_TRUE(sheet->setCellLocked(row_t(1), column_t(1), true));
    EXPECT_TRUE(sheet->isCellLocked(row_t(1), column_t(1)));

    // 生成测试文件
    addTestInfo(sheet, "SetCellLockingViaSheet", "测试通过工作表接口设置单元格锁定状态");

    // 创建锁定状态演示
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"单元格"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"锁定状态"});
    sheet->setCellValue(row_t(7), column_t(3), cell_value_t{"内容"});
    sheet->setCellValue(row_t(7), column_t(4), cell_value_t{"说明"});

    // 锁定的单元格
    sheet->setCellValue(row_t(8), column_t(1), cell_value_t{"A8"});
    sheet->setCellValue(row_t(8), column_t(2), cell_value_t{"锁定"});
    sheet->setCellValue(row_t(8), column_t(3), cell_value_t{"重要数据"});
    sheet->setCellValue(row_t(8), column_t(4), cell_value_t{"此单元格被锁定，保护时无法编辑"});
    sheet->setCellLocked(row_t(8), column_t(3), true);

    // 未锁定的单元格
    sheet->setCellValue(row_t(9), column_t(1), cell_value_t{"A9"});
    sheet->setCellValue(row_t(9), column_t(2), cell_value_t{"未锁定"});
    sheet->setCellValue(row_t(9), column_t(3), cell_value_t{"可编辑数据"});
    sheet->setCellValue(row_t(9), column_t(4), cell_value_t{"此单元格未锁定，保护时仍可编辑"});
    sheet->setCellLocked(row_t(9), column_t(3), false);

    // 混合状态的行
    sheet->setCellValue(row_t(10), column_t(1), cell_value_t{"A10"});
    sheet->setCellValue(row_t(10), column_t(2), cell_value_t{"锁定"});
    sheet->setCellValue(row_t(10), column_t(3), cell_value_t{"标题"});
    sheet->setCellValue(row_t(10), column_t(4), cell_value_t{"标题通常需要锁定"});
    sheet->setCellLocked(row_t(10), column_t(3), true);

    sheet->setCellValue(row_t(11), column_t(1), cell_value_t{"B10"});
    sheet->setCellValue(row_t(11), column_t(2), cell_value_t{"未锁定"});
    sheet->setCellValue(row_t(11), column_t(3), cell_value_t{"输入区域"});
    sheet->setCellValue(row_t(11), column_t(4), cell_value_t{"输入区域通常不锁定"});
    sheet->setCellLocked(row_t(11), column_t(3), false);

    // 添加保护说明
    sheet->setCellValue(row_t(13), column_t(1), cell_value_t{"重要说明:"});
    sheet->setCellValue(row_t(14), column_t(1), cell_value_t{"1. 单元格锁定只有在工作表保护时才生效"});
    sheet->setCellValue(row_t(15), column_t(1), cell_value_t{"2. 此工作表已启用保护，密码为: test123"});
    sheet->setCellValue(row_t(16), column_t(1), cell_value_t{"3. 锁定的单元格无法编辑，未锁定的可以编辑"});

    // 保护工作表以使锁定生效
    sheet->protectSheet("test123");

    saveWorkbook(workbook, "SetCellLockingViaSheet");
}

TEST_F(CellLockingTest, AutoCreateCellForLocking) {
    // 对不存在的单元格设置锁定状态应该自动创建单元格
    EXPECT_TRUE(sheet->setCellLocked(row_t(5), column_t(5), false));
    
    // 验证单元格被创建并且锁定状态正确
    TXCell* cell = sheet->getCell(row_t(5), column_t(5));
    ASSERT_NE(cell, nullptr);
    EXPECT_FALSE(cell->isLocked());
}

TEST_F(CellLockingTest, NonExistentCellDefaultLocked) {
    // 不存在的单元格应该返回默认锁定状态
    EXPECT_TRUE(sheet->isCellLocked(row_t(10), column_t(10)));
}

TEST_F(CellLockingTest, CellCopyPreservesLocking) {
    // 创建原始单元格并设置为解锁
    sheet->setCellValue(row_t(1), column_t(1), cell_value_t{"原始数据"});
    TXCell* originalCell = sheet->getCell(row_t(1), column_t(1));
    originalCell->setLocked(false);
    
    // 测试拷贝构造
    TXCell copiedCell(*originalCell);
    EXPECT_FALSE(copiedCell.isLocked());
    EXPECT_EQ(copiedCell.getStringValue(), "原始数据");
    
    // 测试拷贝赋值
    TXCell assignedCell;
    assignedCell = *originalCell;
    EXPECT_FALSE(assignedCell.isLocked());
    EXPECT_EQ(assignedCell.getStringValue(), "原始数据");
}

TEST_F(CellLockingTest, CellMovePreservesLocking) {
    // 创建原始单元格并设置为解锁
    TXCell originalCell(cell_value_t{"移动数据"});
    originalCell.setLocked(false);
    
    // 测试移动构造
    TXCell movedCell(std::move(originalCell));
    EXPECT_FALSE(movedCell.isLocked());
    EXPECT_EQ(movedCell.getStringValue(), "移动数据");
    
    // 测试移动赋值
    TXCell anotherCell(cell_value_t{"另一个数据"});
    anotherCell.setLocked(false);
    
    TXCell assignedCell;
    assignedCell = std::move(anotherCell);
    EXPECT_FALSE(assignedCell.isLocked());
    EXPECT_EQ(assignedCell.getStringValue(), "另一个数据");
}

TEST_F(CellLockingTest, HasFormulaMethod) {
    // 创建普通单元格
    TXCell normalCell(cell_value_t{"普通数据"});
    EXPECT_FALSE(normalCell.hasFormula());
    
    // 创建公式单元格
    TXCell formulaCell;
    formulaCell.setFormula("A1+B1");
    EXPECT_TRUE(formulaCell.hasFormula());
    
    // 清除公式
    formulaCell.setFormula("");
    EXPECT_FALSE(formulaCell.hasFormula());
}