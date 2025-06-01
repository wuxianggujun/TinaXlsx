#include <gtest/gtest.h>
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXStyle.hpp"
#include "test_file_generator.hpp"

using namespace TinaXlsx;

class TXSheetRefactoredIntegrationTest : public TestWithFileGeneration<TXSheetRefactoredIntegrationTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<TXSheetRefactoredIntegrationTest>::SetUp();
        workbook = std::make_unique<TXWorkbook>();
        sheet = std::make_unique<TXSheet>("TestSheet", workbook.get());
    }

    void TearDown() override {
        sheet.reset();
        workbook.reset();
        TestWithFileGeneration<TXSheetRefactoredIntegrationTest>::TearDown();
    }

    std::unique_ptr<TXWorkbook> workbook;
    std::unique_ptr<TXSheet> sheet;
};

// ==================== 基本功能测试 ====================

TEST_F(TXSheetRefactoredIntegrationTest, BasicProperties) {
    EXPECT_EQ(sheet->getName(), "TestSheet");
    EXPECT_EQ(sheet->getWorkbook(), workbook.get());
    EXPECT_TRUE(sheet->getLastError().empty());
}

TEST_F(TXSheetRefactoredIntegrationTest, BasicCellOperations) {
    // 设置单元格值
    EXPECT_TRUE(sheet->setCellValue(row_t(1), column_t(1), std::string("Hello")));
    EXPECT_TRUE(sheet->setCellValue(row_t(1), column_t(2), 123.45));
    EXPECT_TRUE(sheet->setCellValue(row_t(1), column_t(3), true));
    
    // 获取单元格值
    auto value1 = sheet->getCellValue(row_t(1), column_t(1));
    auto value2 = sheet->getCellValue(row_t(1), column_t(2));
    auto value3 = sheet->getCellValue(row_t(1), column_t(3));
    
    EXPECT_EQ(std::get<std::string>(value1), "Hello");
    EXPECT_DOUBLE_EQ(std::get<double>(value2), 123.45);
    EXPECT_EQ(std::get<bool>(value3), true);
    
    // 获取单元格对象
    auto* cell = sheet->getCell(row_t(1), column_t(1));
    ASSERT_NE(cell, nullptr);
    auto cellValue = cell->getValue();
    EXPECT_EQ(std::get<std::string>(cellValue), "Hello");
}

TEST_F(TXSheetRefactoredIntegrationTest, CoordinateOperations) {
    TXCoordinate coord(row_t(2), column_t(3));
    
    // 使用坐标设置值
    EXPECT_TRUE(sheet->setCellValue(coord, std::string("C2")));
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(coord)), "C2");
}

// ==================== 行列操作测试 ====================

TEST_F(TXSheetRefactoredIntegrationTest, RowOperations) {
    // 设置初始数据
    sheet->setCellValue(row_t(1), column_t(1), std::string("A1"));
    sheet->setCellValue(row_t(2), column_t(1), std::string("A2"));
    sheet->setCellValue(row_t(3), column_t(1), std::string("A3"));
    
    // 插入行
    EXPECT_TRUE(sheet->insertRows(row_t(2), row_t(1)));
    
    // 验证数据移动
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(1), column_t(1))), "A1");
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(3), column_t(1))), "A2");
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(4), column_t(1))), "A3");
    
    // 删除行
    EXPECT_TRUE(sheet->deleteRows(row_t(2), row_t(1)));
    
    // 验证数据移动
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(1), column_t(1))), "A1");
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(2), column_t(1))), "A2");
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(3), column_t(1))), "A3");
}

TEST_F(TXSheetRefactoredIntegrationTest, ColumnOperations) {
    // 设置初始数据
    sheet->setCellValue(row_t(1), column_t(1), std::string("A1"));
    sheet->setCellValue(row_t(1), column_t(2), std::string("B1"));
    sheet->setCellValue(row_t(1), column_t(3), std::string("C1"));
    
    // 插入列
    EXPECT_TRUE(sheet->insertColumns(column_t(2), column_t(1)));
    
    // 验证数据移动
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(1), column_t(1))), "A1");
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(1), column_t(3))), "B1");
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(1), column_t(4))), "C1");
    
    // 删除列
    EXPECT_TRUE(sheet->deleteColumns(column_t(2), column_t(1)));
    
    // 验证数据移动
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(1), column_t(1))), "A1");
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(1), column_t(2))), "B1");
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(1), column_t(3))), "C1");
}

TEST_F(TXSheetRefactoredIntegrationTest, RowColumnSizing) {
    // 设置行高
    EXPECT_TRUE(sheet->setRowHeight(row_t(1), 25.0));
    EXPECT_DOUBLE_EQ(sheet->getRowHeight(row_t(1)), 25.0);
    
    // 设置列宽
    EXPECT_TRUE(sheet->setColumnWidth(column_t(1), 15.0));
    EXPECT_DOUBLE_EQ(sheet->getColumnWidth(column_t(1)), 15.0);
    
    // 自动调整
    sheet->setCellValue(row_t(1), column_t(1), std::string("Very long text content"));
    double newWidth = sheet->autoFitColumnWidth(column_t(1));
    EXPECT_GT(newWidth, 8.43);  // 应该大于默认宽度
}

// ==================== 工作表保护测试 ====================

TEST_F(TXSheetRefactoredIntegrationTest, SheetProtection) {
    // 初始状态未保护
    EXPECT_FALSE(sheet->isSheetProtected());
    
    // 保护工作表
    EXPECT_TRUE(sheet->protectSheet("password123"));
    EXPECT_TRUE(sheet->isSheetProtected());
    
    // 保护状态下的操作应该被阻止
    EXPECT_FALSE(sheet->insertRows(row_t(1), row_t(1)));
    EXPECT_FALSE(sheet->deleteRows(row_t(1), row_t(1)));
    EXPECT_FALSE(sheet->setRowHeight(row_t(1), 25.0));
    
    // 取消保护
    EXPECT_TRUE(sheet->unprotectSheet("password123"));
    EXPECT_FALSE(sheet->isSheetProtected());
    
    // 现在操作应该被允许
    EXPECT_TRUE(sheet->setRowHeight(row_t(1), 25.0));
}

TEST_F(TXSheetRefactoredIntegrationTest, CellLocking) {
    // 默认锁定状态
    EXPECT_TRUE(sheet->isCellLocked(row_t(1), column_t(1)));
    
    // 解锁单元格
    EXPECT_TRUE(sheet->setCellLocked(row_t(1), column_t(1), false));
    EXPECT_FALSE(sheet->isCellLocked(row_t(1), column_t(1)));
    
    // 重新锁定
    EXPECT_TRUE(sheet->setCellLocked(row_t(1), column_t(1), true));
    EXPECT_TRUE(sheet->isCellLocked(row_t(1), column_t(1)));
}

// ==================== 公式操作测试 ====================

TEST_F(TXSheetRefactoredIntegrationTest, FormulaOperations) {
    // 设置公式
    EXPECT_TRUE(sheet->setCellFormula(row_t(3), column_t(1), "=1+2"));
    EXPECT_EQ(sheet->getCellFormula(row_t(3), column_t(1)), "=1+2");
    
    // 计算公式
    std::size_t count = sheet->calculateAllFormulas();
    EXPECT_GT(count, 0);
    
    // 验证公式存在
    auto* cell = sheet->getCell(row_t(3), column_t(1));
    ASSERT_NE(cell, nullptr);
    EXPECT_TRUE(cell->hasFormula());
}

TEST_F(TXSheetRefactoredIntegrationTest, NamedRanges) {
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(5), column_t(1)));
    
    // 添加命名范围
    EXPECT_TRUE(sheet->addNamedRange("TestRange", range));
    
    // 获取命名范围
    auto retrievedRange = sheet->getNamedRange("TestRange");
    EXPECT_TRUE(retrievedRange.isValid());
    EXPECT_EQ(retrievedRange.getStart().getRow(), row_t(1));
    EXPECT_EQ(retrievedRange.getEnd().getRow(), row_t(5));
    
    // 删除命名范围
    EXPECT_TRUE(sheet->removeNamedRange("TestRange"));
    auto emptyRange = sheet->getNamedRange("TestRange");
    EXPECT_FALSE(emptyRange.isValid());
}

// ==================== 合并单元格测试 ====================

TEST_F(TXSheetRefactoredIntegrationTest, MergeCells) {
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(2), column_t(2)));
    
    // 合并单元格
    EXPECT_TRUE(sheet->mergeCells(range));
    
    // 验证合并状态
    EXPECT_TRUE(sheet->isCellMerged(row_t(1), column_t(1)));
    EXPECT_TRUE(sheet->isCellMerged(row_t(2), column_t(2)));
    
    // 获取合并区域
    auto mergeRegion = sheet->getMergeRegion(row_t(1), column_t(2));
    EXPECT_TRUE(mergeRegion.isValid());
    EXPECT_EQ(mergeRegion.getStart().getRow(), row_t(1));
    EXPECT_EQ(mergeRegion.getEnd().getRow(), row_t(2));
    
    // 取消合并
    EXPECT_TRUE(sheet->unmergeCells(row_t(1), column_t(1)));
    EXPECT_FALSE(sheet->isCellMerged(row_t(1), column_t(1)));
}

TEST_F(TXSheetRefactoredIntegrationTest, MergeCellsWithCoordinates) {
    // 使用坐标合并
    EXPECT_TRUE(sheet->mergeCells(row_t(3), column_t(1), row_t(4), column_t(3)));
    
    // 验证合并
    EXPECT_TRUE(sheet->isCellMerged(row_t(3), column_t(2)));
    EXPECT_TRUE(sheet->isCellMerged(row_t(4), column_t(3)));
    
    // 获取合并数量
    EXPECT_EQ(sheet->getMergeCount(), 1);
    
    // 获取所有合并区域
    auto allRegions = sheet->getAllMergeRegions();
    EXPECT_EQ(allRegions.size(), 1);
}

// ==================== 批量操作测试 ====================

TEST_F(TXSheetRefactoredIntegrationTest, BatchOperations) {
    // 批量设置值
    std::vector<std::pair<TXCoordinate, cell_value_t>> values = {
        {TXCoordinate(row_t(1), column_t(1)), std::string("A1")},
        {TXCoordinate(row_t(1), column_t(2)), std::string("B1")},
        {TXCoordinate(row_t(2), column_t(1)), 123.0},
        {TXCoordinate(row_t(2), column_t(2)), true}
    };
    
    std::size_t count = sheet->setCellValues(values);
    EXPECT_EQ(count, 4);
    
    // 批量获取值
    std::vector<TXCoordinate> coords = {
        TXCoordinate(row_t(1), column_t(1)),
        TXCoordinate(row_t(1), column_t(2)),
        TXCoordinate(row_t(2), column_t(1)),
        TXCoordinate(row_t(2), column_t(2))
    };
    
    auto result = sheet->getCellValues(coords);
    EXPECT_EQ(result.size(), 4);

    // 生成测试文件
    // 注意：这里需要使用workbook中的sheet，而不是独立的sheet
    auto* workbookSheet = workbook->addSheet("集成测试");
    addTestInfo(workbookSheet, "BatchOperations", "测试TXSheet重构后的批量操作功能");

    // 复制测试数据到工作簿的工作表中
    workbookSheet->setCellValue(row_t(7), column_t(1), cell_value_t{"坐标"});
    workbookSheet->setCellValue(row_t(7), column_t(2), cell_value_t{"数据类型"});
    workbookSheet->setCellValue(row_t(7), column_t(3), cell_value_t{"值"});
    workbookSheet->setCellValue(row_t(7), column_t(4), cell_value_t{"说明"});

    workbookSheet->setCellValue(row_t(8), column_t(1), cell_value_t{"A1"});
    workbookSheet->setCellValue(row_t(8), column_t(2), cell_value_t{"字符串"});
    workbookSheet->setCellValue(row_t(8), column_t(3), cell_value_t{"A1"});
    workbookSheet->setCellValue(row_t(8), column_t(4), cell_value_t{"批量设置的字符串值"});

    workbookSheet->setCellValue(row_t(9), column_t(1), cell_value_t{"B1"});
    workbookSheet->setCellValue(row_t(9), column_t(2), cell_value_t{"字符串"});
    workbookSheet->setCellValue(row_t(9), column_t(3), cell_value_t{"B1"});
    workbookSheet->setCellValue(row_t(9), column_t(4), cell_value_t{"批量设置的字符串值"});

    workbookSheet->setCellValue(row_t(10), column_t(1), cell_value_t{"A2"});
    workbookSheet->setCellValue(row_t(10), column_t(2), cell_value_t{"数字"});
    workbookSheet->setCellValue(row_t(10), column_t(3), cell_value_t{123.0});
    workbookSheet->setCellValue(row_t(10), column_t(4), cell_value_t{"批量设置的数字值"});

    workbookSheet->setCellValue(row_t(11), column_t(1), cell_value_t{"B2"});
    workbookSheet->setCellValue(row_t(11), column_t(2), cell_value_t{"布尔值"});
    workbookSheet->setCellValue(row_t(11), column_t(3), cell_value_t{true});
    workbookSheet->setCellValue(row_t(11), column_t(4), cell_value_t{"批量设置的布尔值"});

    workbookSheet->setCellValue(row_t(13), column_t(1), cell_value_t{"批量操作统计:"});
    workbookSheet->setCellValue(row_t(13), column_t(2), cell_value_t{"成功设置"});
    workbookSheet->setCellValue(row_t(13), column_t(3), cell_value_t{static_cast<double>(count)});
    workbookSheet->setCellValue(row_t(13), column_t(4), cell_value_t{"个单元格"});

    workbookSheet->setCellValue(row_t(14), column_t(2), cell_value_t{"成功获取"});
    workbookSheet->setCellValue(row_t(14), column_t(3), cell_value_t{static_cast<double>(result.size())});
    workbookSheet->setCellValue(row_t(14), column_t(4), cell_value_t{"个单元格值"});

    saveWorkbook(workbook, "BatchOperations");
}

// ==================== 范围操作测试 ====================

TEST_F(TXSheetRefactoredIntegrationTest, RangeOperations) {
    // 设置范围值
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(2), column_t(2)));
    std::vector<std::vector<cell_value_t>> values = {
        {std::string("A1"), std::string("B1")},
        {123.0, 456.0}
    };
    
    EXPECT_TRUE(sheet->setRangeValues(range, values));
    
    // 获取范围值
    auto retrievedValues = sheet->getRangeValues(range);
    EXPECT_EQ(retrievedValues.size(), 2);
    EXPECT_EQ(retrievedValues[0].size(), 2);
    EXPECT_EQ(std::get<std::string>(retrievedValues[0][0]), "A1");
    EXPECT_EQ(std::get<std::string>(retrievedValues[0][1]), "B1");
    EXPECT_DOUBLE_EQ(std::get<double>(retrievedValues[1][0]), 123.0);
    EXPECT_DOUBLE_EQ(std::get<double>(retrievedValues[1][1]), 456.0);
}

// ==================== 查询操作测试 ====================

TEST_F(TXSheetRefactoredIntegrationTest, QueryOperations) {
    // 设置一些数据
    sheet->setCellValue(row_t(1), column_t(1), std::string("A1"));
    sheet->setCellValue(row_t(3), column_t(2), std::string("B3"));
    sheet->setCellValue(row_t(5), column_t(4), 123.0);
    
    // 获取使用范围
    auto usedRange = sheet->getUsedRange();
    EXPECT_TRUE(usedRange.isValid());
    EXPECT_EQ(usedRange.getStart().getRow(), row_t(1));
    EXPECT_EQ(usedRange.getStart().getCol(), column_t(1));
    EXPECT_EQ(usedRange.getEnd().getRow(), row_t(5));
    EXPECT_EQ(usedRange.getEnd().getCol(), column_t(4));
    
    // 获取使用的行列数
    EXPECT_EQ(sheet->getUsedRowCount(), row_t(5));
    EXPECT_EQ(sheet->getUsedColumnCount(), column_t(4));
}

// ==================== 管理器访问测试 ====================

TEST_F(TXSheetRefactoredIntegrationTest, ManagerAccess) {
    // 测试管理器访问
    auto& cellManager = sheet->getCellManager();
    auto& rowColManager = sheet->getRowColumnManager();
    auto& protectionManager = sheet->getProtectionManager();
    auto& formulaManager = sheet->getFormulaManager();
    auto& mergedCells = sheet->getMergedCells();
    
    // 验证管理器可以直接使用
    EXPECT_TRUE(cellManager.setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("Direct")));
    EXPECT_TRUE(rowColManager.setRowHeight(row_t(1), 20.0));
    EXPECT_FALSE(protectionManager.isSheetProtected());
    EXPECT_EQ(formulaManager.getAllNamedRanges().size(), 0);
    EXPECT_EQ(mergedCells.getMergeCount(), 0);
}

// ==================== 清空操作测试 ====================

TEST_F(TXSheetRefactoredIntegrationTest, ClearOperations) {
    // 设置一些数据
    sheet->setCellValue(row_t(1), column_t(1), std::string("Test"));
    sheet->setRowHeight(row_t(1), 25.0);
    sheet->protectSheet("password");
    sheet->addNamedRange("TestRange", TXRange(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(5), column_t(1))));
    
    // 验证数据存在
    EXPECT_TRUE(sheet->getCell(row_t(1), column_t(1)) != nullptr);
    EXPECT_DOUBLE_EQ(sheet->getRowHeight(row_t(1)), 25.0);
    EXPECT_TRUE(sheet->isSheetProtected());
    EXPECT_TRUE(sheet->getNamedRange("TestRange").isValid());
    
    // 清空
    sheet->clear();
    
    // 验证清空结果
    EXPECT_EQ(sheet->getCell(row_t(1), column_t(1)), nullptr);
    EXPECT_DOUBLE_EQ(sheet->getRowHeight(row_t(1)), 15.0);  // 默认行高
    EXPECT_FALSE(sheet->isSheetProtected());
    EXPECT_FALSE(sheet->getNamedRange("TestRange").isValid());
}
