#include <gtest/gtest.h>
#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <memory>

using namespace TinaXlsx;

class CellLockingTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("锁定测试");
    }

    void TearDown() override {
        workbook.reset();
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