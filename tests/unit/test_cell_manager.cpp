#include <gtest/gtest.h>
#include "TinaXlsx/TXCellManager.hpp"
#include "TinaXlsx/TXCoordinate.hpp"
#include "TinaXlsx/TXRange.hpp"
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "test_file_generator.hpp"

using namespace TinaXlsx;

class TXCellManagerTest : public TestWithFileGeneration<TXCellManagerTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<TXCellManagerTest>::SetUp();
        cellManager = std::make_unique<TXCellManager>();
        // 为文件生成创建工作簿和工作表
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("CellManager测试");
    }

    void TearDown() override {
        cellManager.reset();
        workbook.reset();
        TestWithFileGeneration<TXCellManagerTest>::TearDown();
    }

    std::unique_ptr<TXCellManager> cellManager;
    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
};

// ==================== 基本单元格操作测试 ====================

TEST_F(TXCellManagerTest, BasicCellOperations) {
    TXCoordinate coord(row_t(1), column_t(1));
    
    // 测试获取不存在的单元格
    EXPECT_EQ(cellManager->getCell(coord), nullptr);
    
    // 测试设置单元格值
    EXPECT_TRUE(cellManager->setCellValue(coord, std::string("Hello")));
    
    // 测试获取单元格值
    auto value = cellManager->getCellValue(coord);
    EXPECT_EQ(std::get<std::string>(value), "Hello");
    
    // 测试单元格存在性
    EXPECT_TRUE(cellManager->hasCell(coord));
    
    // 测试获取单元格对象
    auto* cell = cellManager->getCell(coord);
    ASSERT_NE(cell, nullptr);
    auto cellValue = cell->getValue();
    EXPECT_EQ(std::get<std::string>(cellValue), "Hello");
}

TEST_F(TXCellManagerTest, DifferentDataTypes) {
    // 测试字符串
    TXCoordinate coord1(row_t(1), column_t(1));
    EXPECT_TRUE(cellManager->setCellValue(coord1, std::string("Text")));
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(coord1)), "Text");
    
    // 测试数字
    TXCoordinate coord2(row_t(1), column_t(2));
    EXPECT_TRUE(cellManager->setCellValue(coord2, 123.45));
    EXPECT_DOUBLE_EQ(std::get<double>(cellManager->getCellValue(coord2)), 123.45);
    
    // 测试布尔值
    TXCoordinate coord3(row_t(1), column_t(3));
    EXPECT_TRUE(cellManager->setCellValue(coord3, true));
    EXPECT_EQ(std::get<bool>(cellManager->getCellValue(coord3)), true);
}

TEST_F(TXCellManagerTest, InvalidCoordinates) {
    TXCoordinate invalidCoord(row_t(static_cast<row_t::index_t>(0)), column_t(static_cast<column_t::index_t>(0))); // 0,0是无效坐标

    // 测试无效坐标操作
    EXPECT_FALSE(cellManager->setCellValue(invalidCoord, std::string("Test")));
    EXPECT_EQ(cellManager->getCell(invalidCoord), nullptr);
}

TEST_F(TXCellManagerTest, RemoveCell) {
    TXCoordinate coord(row_t(1), column_t(1));
    
    // 设置单元格值
    EXPECT_TRUE(cellManager->setCellValue(coord, std::string("Test")));
    EXPECT_TRUE(cellManager->hasCell(coord));
    
    // 删除单元格
    EXPECT_TRUE(cellManager->removeCell(coord));
    EXPECT_FALSE(cellManager->hasCell(coord));
    
    // 删除不存在的单元格
    EXPECT_FALSE(cellManager->removeCell(coord));
}

// ==================== 批量操作测试 ====================

TEST_F(TXCellManagerTest, BatchOperations) {
    std::vector<std::pair<TXCoordinate, cell_value_t>> values = {
        {TXCoordinate(row_t(1), column_t(1)), std::string("A1")},
        {TXCoordinate(row_t(1), column_t(2)), std::string("B1")},
        {TXCoordinate(row_t(2), column_t(1)), 123.0},
        {TXCoordinate(row_t(2), column_t(2)), true}
    };
    
    // 批量设置值
    std::size_t count = cellManager->setCellValues(values);
    EXPECT_EQ(count, 4);
    
    // 验证设置的值
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(1), column_t(1)))), "A1");
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(1), column_t(2)))), "B1");
    EXPECT_DOUBLE_EQ(std::get<double>(cellManager->getCellValue(TXCoordinate(row_t(2), column_t(1)))), 123.0);
    EXPECT_EQ(std::get<bool>(cellManager->getCellValue(TXCoordinate(row_t(2), column_t(2)))), true);

    // 生成测试文件
    addTestInfo(sheet, "BatchOperations", "测试TXCellManager批量操作功能");

    // 将测试数据复制到工作表中进行演示
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"坐标"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"数据类型"});
    sheet->setCellValue(row_t(7), column_t(3), cell_value_t{"值"});
    sheet->setCellValue(row_t(7), column_t(4), cell_value_t{"说明"});

    sheet->setCellValue(row_t(8), column_t(1), cell_value_t{"A1"});
    sheet->setCellValue(row_t(8), column_t(2), cell_value_t{"字符串"});
    sheet->setCellValue(row_t(8), column_t(3), cell_value_t{"A1"});
    sheet->setCellValue(row_t(8), column_t(4), cell_value_t{"批量设置的字符串值"});

    sheet->setCellValue(row_t(9), column_t(1), cell_value_t{"B1"});
    sheet->setCellValue(row_t(9), column_t(2), cell_value_t{"字符串"});
    sheet->setCellValue(row_t(9), column_t(3), cell_value_t{"B1"});
    sheet->setCellValue(row_t(9), column_t(4), cell_value_t{"批量设置的字符串值"});

    sheet->setCellValue(row_t(10), column_t(1), cell_value_t{"A2"});
    sheet->setCellValue(row_t(10), column_t(2), cell_value_t{"数字"});
    sheet->setCellValue(row_t(10), column_t(3), cell_value_t{123.0});
    sheet->setCellValue(row_t(10), column_t(4), cell_value_t{"批量设置的数字值"});

    sheet->setCellValue(row_t(11), column_t(1), cell_value_t{"B2"});
    sheet->setCellValue(row_t(11), column_t(2), cell_value_t{"布尔值"});
    sheet->setCellValue(row_t(11), column_t(3), cell_value_t{true});
    sheet->setCellValue(row_t(11), column_t(4), cell_value_t{"批量设置的布尔值"});

    sheet->setCellValue(row_t(13), column_t(1), cell_value_t{"批量操作统计:"});
    sheet->setCellValue(row_t(13), column_t(2), cell_value_t{"成功设置"});
    sheet->setCellValue(row_t(13), column_t(3), cell_value_t{static_cast<double>(count)});
    sheet->setCellValue(row_t(13), column_t(4), cell_value_t{"个单元格"});

    saveWorkbook(workbook, "BatchOperations");
}

TEST_F(TXCellManagerTest, BatchGetValues) {
    // 先设置一些值
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("A1"));
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(2)), 42.0);
    cellManager->setCellValue(TXCoordinate(row_t(2), column_t(1)), false);
    
    std::vector<TXCoordinate> coords = {
        TXCoordinate(row_t(1), column_t(1)),
        TXCoordinate(row_t(1), column_t(2)),
        TXCoordinate(row_t(2), column_t(1)),
        TXCoordinate(row_t(3), column_t(1))  // 不存在的单元格
    };
    
    auto result = cellManager->getCellValues(coords);
    EXPECT_EQ(result.size(), 4);
    
    EXPECT_EQ(std::get<std::string>(result[0].second), "A1");
    EXPECT_DOUBLE_EQ(std::get<double>(result[1].second), 42.0);
    EXPECT_EQ(std::get<bool>(result[2].second), false);
    EXPECT_EQ(std::get<std::string>(result[3].second), ""); // 默认空字符串
}

// ==================== 范围操作测试 ====================

TEST_F(TXCellManagerTest, UsedRange) {
    // 空的管理器应该返回无效范围
    auto emptyRange = cellManager->getUsedRange();
    EXPECT_FALSE(emptyRange.isValid());
    
    // 设置一些单元格
    cellManager->setCellValue(TXCoordinate(row_t(2), column_t(2)), std::string("B2"));
    cellManager->setCellValue(TXCoordinate(row_t(5), column_t(4)), std::string("D5"));
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("A1"));
    
    auto usedRange = cellManager->getUsedRange();
    EXPECT_TRUE(usedRange.isValid());
    EXPECT_EQ(usedRange.getStart().getRow(), row_t(1));
    EXPECT_EQ(usedRange.getStart().getCol(), column_t(1));
    EXPECT_EQ(usedRange.getEnd().getRow(), row_t(5));
    EXPECT_EQ(usedRange.getEnd().getCol(), column_t(4));
}

TEST_F(TXCellManagerTest, MaxUsedRowColumn) {
    // 设置一些单元格
    cellManager->setCellValue(TXCoordinate(row_t(3), column_t(2)), std::string("B3"));
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(5)), std::string("E1"));
    cellManager->setCellValue(TXCoordinate(row_t(7), column_t(1)), std::string("A7"));
    
    EXPECT_EQ(cellManager->getMaxUsedRow(), row_t(7));
    EXPECT_EQ(cellManager->getMaxUsedColumn(), column_t(5));
}

TEST_F(TXCellManagerTest, CellCount) {
    EXPECT_EQ(cellManager->getCellCount(), 0);
    EXPECT_EQ(cellManager->getNonEmptyCellCount(), 0);
    
    // 添加一些单元格
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("A1"));
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(2)), std::string(""));  // 空字符串
    cellManager->setCellValue(TXCoordinate(row_t(2), column_t(1)), 42.0);
    
    EXPECT_EQ(cellManager->getCellCount(), 3);
    EXPECT_EQ(cellManager->getNonEmptyCellCount(), 2);  // 空字符串被认为是空的
}

// ==================== 坐标变换测试 ====================

TEST_F(TXCellManagerTest, TransformCells) {
    // 设置一些初始单元格
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("A1"));
    cellManager->setCellValue(TXCoordinate(row_t(2), column_t(1)), std::string("A2"));
    cellManager->setCellValue(TXCoordinate(row_t(3), column_t(1)), std::string("A3"));
    
    // 定义变换函数：所有行向下移动2行
    auto transform = [](const TXCoordinate& coord) -> TXCoordinate {
        return TXCoordinate(row_t(coord.getRow().index() + 2), coord.getCol());
    };
    
    cellManager->transformCells(transform);
    
    // 验证变换结果：原位置应该为空
    EXPECT_FALSE(cellManager->hasCell(TXCoordinate(row_t(1), column_t(1))));
    EXPECT_FALSE(cellManager->hasCell(TXCoordinate(row_t(2), column_t(1))));

    // 新位置应该有数据
    EXPECT_TRUE(cellManager->hasCell(TXCoordinate(row_t(3), column_t(1))));
    EXPECT_TRUE(cellManager->hasCell(TXCoordinate(row_t(4), column_t(1))));
    EXPECT_TRUE(cellManager->hasCell(TXCoordinate(row_t(5), column_t(1))));

    // 验证数据内容
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(3), column_t(1)))), "A1");
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(4), column_t(1)))), "A2");
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(5), column_t(1)))), "A3");
}

TEST_F(TXCellManagerTest, RemoveCellsInRange) {
    // 设置一些单元格
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("A1"));
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(2)), std::string("B1"));
    cellManager->setCellValue(TXCoordinate(row_t(2), column_t(1)), std::string("A2"));
    cellManager->setCellValue(TXCoordinate(row_t(2), column_t(2)), std::string("B2"));
    cellManager->setCellValue(TXCoordinate(row_t(3), column_t(1)), std::string("A3"));
    
    EXPECT_EQ(cellManager->getCellCount(), 5);
    
    // 删除A1:B2范围内的单元格
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(2), column_t(2)));
    std::size_t removedCount = cellManager->removeCellsInRange(range);
    
    EXPECT_EQ(removedCount, 4);
    EXPECT_EQ(cellManager->getCellCount(), 1);
    EXPECT_TRUE(cellManager->hasCell(TXCoordinate(row_t(3), column_t(1))));
}

// ==================== 清空操作测试 ====================

TEST_F(TXCellManagerTest, Clear) {
    // 添加一些单元格
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("A1"));
    cellManager->setCellValue(TXCoordinate(row_t(2), column_t(2)), 42.0);
    
    EXPECT_EQ(cellManager->getCellCount(), 2);
    
    // 清空
    cellManager->clear();
    
    EXPECT_EQ(cellManager->getCellCount(), 0);
    EXPECT_FALSE(cellManager->hasCell(TXCoordinate(row_t(1), column_t(1))));
    EXPECT_FALSE(cellManager->hasCell(TXCoordinate(row_t(2), column_t(2))));
}

// ==================== 迭代器测试 ====================

TEST_F(TXCellManagerTest, Iterators) {
    // 添加一些单元格
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("A1"));
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(2)), std::string("B1"));
    cellManager->setCellValue(TXCoordinate(row_t(2), column_t(1)), 42.0);
    
    // 测试迭代器
    std::size_t count = 0;
    for (auto it = cellManager->begin(); it != cellManager->end(); ++it) {
        ++count;
        EXPECT_TRUE(it->first.isValid());
    }
    EXPECT_EQ(count, 3);
    
    // 测试const迭代器
    const auto& constCellManager = *cellManager;
    count = 0;
    for (auto it = constCellManager.begin(); it != constCellManager.end(); ++it) {
        ++count;
    }
    EXPECT_EQ(count, 3);
}
