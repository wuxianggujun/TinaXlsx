#include <gtest/gtest.h>
#include "TinaXlsx/TXRowColumnManager.hpp"
#include "TinaXlsx/TXCellManager.hpp"
#include "TinaXlsx/TXCoordinate.hpp"
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "test_file_generator.hpp"

using namespace TinaXlsx;

class TXRowColumnManagerTest : public TestWithFileGeneration<TXRowColumnManagerTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<TXRowColumnManagerTest>::SetUp();
        rowColManager = std::make_unique<TXRowColumnManager>();
        cellManager = std::make_unique<TXCellManager>();

        // 为文件生成创建工作簿和工作表
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("RowColumn管理器测试");

        // 设置一些测试数据
        cellManager->setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("A1"));
        cellManager->setCellValue(TXCoordinate(row_t(1), column_t(2)), std::string("B1"));
        cellManager->setCellValue(TXCoordinate(row_t(2), column_t(1)), std::string("A2"));
        cellManager->setCellValue(TXCoordinate(row_t(2), column_t(2)), std::string("B2"));
        cellManager->setCellValue(TXCoordinate(row_t(3), column_t(1)), std::string("A3"));
    }

    void TearDown() override {
        rowColManager.reset();
        cellManager.reset();
        workbook.reset();
        TestWithFileGeneration<TXRowColumnManagerTest>::TearDown();
    }

    std::unique_ptr<TXRowColumnManager> rowColManager;
    std::unique_ptr<TXCellManager> cellManager;
    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
};

// ==================== 行操作测试 ====================

TEST_F(TXRowColumnManagerTest, InsertRows) {
    // 在第2行插入1行
    EXPECT_TRUE(rowColManager->insertRows(row_t(2), row_t(1), *cellManager));
    
    // 验证单元格移动
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(1), column_t(1)))), "A1");
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(1), column_t(2)))), "B1");
    
    // 原来的第2行现在应该在第3行
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(3), column_t(1)))), "A2");
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(3), column_t(2)))), "B2");
    
    // 原来的第3行现在应该在第4行
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(4), column_t(1)))), "A3");
    
    // 第2行应该是空的
    EXPECT_FALSE(cellManager->hasCell(TXCoordinate(row_t(2), column_t(1))));
}

TEST_F(TXRowColumnManagerTest, DeleteRows) {
    // 删除第2行
    EXPECT_TRUE(rowColManager->deleteRows(row_t(2), row_t(1), *cellManager));
    
    // 验证单元格移动
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(1), column_t(1)))), "A1");
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(1), column_t(2)))), "B1");
    
    // 原来的第3行现在应该在第2行
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(2), column_t(1)))), "A3");
    
    // 原来的第2行应该被删除
    EXPECT_FALSE(cellManager->hasCell(TXCoordinate(row_t(3), column_t(1))));
}

TEST_F(TXRowColumnManagerTest, RowHeight) {
    // 测试默认行高
    EXPECT_DOUBLE_EQ(rowColManager->getRowHeight(row_t(1)), 15.0);
    
    // 设置行高
    EXPECT_TRUE(rowColManager->setRowHeight(row_t(1), 25.0));
    EXPECT_DOUBLE_EQ(rowColManager->getRowHeight(row_t(1)), 25.0);
    
    // 测试无效行高
    EXPECT_FALSE(rowColManager->setRowHeight(row_t(1), -5.0));
    EXPECT_FALSE(rowColManager->setRowHeight(row_t(1), 500.0));
}

TEST_F(TXRowColumnManagerTest, RowHidden) {
    // 测试默认隐藏状态
    EXPECT_FALSE(rowColManager->isRowHidden(row_t(1)));
    
    // 隐藏行
    EXPECT_TRUE(rowColManager->setRowHidden(row_t(1), true));
    EXPECT_TRUE(rowColManager->isRowHidden(row_t(1)));
    
    // 显示行
    EXPECT_TRUE(rowColManager->setRowHidden(row_t(1), false));
    EXPECT_FALSE(rowColManager->isRowHidden(row_t(1)));
}

// ==================== 列操作测试 ====================

TEST_F(TXRowColumnManagerTest, InsertColumns) {
    // 在第2列插入1列
    EXPECT_TRUE(rowColManager->insertColumns(column_t(2), column_t(1), *cellManager));
    
    // 验证单元格移动
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(1), column_t(1)))), "A1");
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(2), column_t(1)))), "A2");
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(3), column_t(1)))), "A3");
    
    // 原来的第2列现在应该在第3列
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(1), column_t(3)))), "B1");
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(2), column_t(3)))), "B2");
    
    // 第2列应该是空的
    EXPECT_FALSE(cellManager->hasCell(TXCoordinate(row_t(1), column_t(2))));
}

TEST_F(TXRowColumnManagerTest, DeleteColumns) {
    // 删除第2列
    EXPECT_TRUE(rowColManager->deleteColumns(column_t(2), column_t(1), *cellManager));
    
    // 验证单元格移动
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(1), column_t(1)))), "A1");
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(2), column_t(1)))), "A2");
    EXPECT_EQ(std::get<std::string>(cellManager->getCellValue(TXCoordinate(row_t(3), column_t(1)))), "A3");
    
    // 原来的第2列应该被删除
    EXPECT_FALSE(cellManager->hasCell(TXCoordinate(row_t(1), column_t(2))));
    EXPECT_FALSE(cellManager->hasCell(TXCoordinate(row_t(2), column_t(2))));
}

TEST_F(TXRowColumnManagerTest, ColumnWidth) {
    // 测试默认列宽
    EXPECT_DOUBLE_EQ(rowColManager->getColumnWidth(column_t(1)), 8.43);
    
    // 设置列宽
    EXPECT_TRUE(rowColManager->setColumnWidth(column_t(1), 15.0));
    EXPECT_DOUBLE_EQ(rowColManager->getColumnWidth(column_t(1)), 15.0);
    
    // 测试无效列宽
    EXPECT_FALSE(rowColManager->setColumnWidth(column_t(1), -5.0));
    EXPECT_FALSE(rowColManager->setColumnWidth(column_t(1), 300.0));
}

TEST_F(TXRowColumnManagerTest, ColumnHidden) {
    // 测试默认隐藏状态
    EXPECT_FALSE(rowColManager->isColumnHidden(column_t(1)));
    
    // 隐藏列
    EXPECT_TRUE(rowColManager->setColumnHidden(column_t(1), true));
    EXPECT_TRUE(rowColManager->isColumnHidden(column_t(1)));
    
    // 显示列
    EXPECT_TRUE(rowColManager->setColumnHidden(column_t(1), false));
    EXPECT_FALSE(rowColManager->isColumnHidden(column_t(1)));
}

// ==================== 自动调整测试 ====================

TEST_F(TXRowColumnManagerTest, AutoFitColumnWidth) {
    // 设置一些有内容的单元格
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("Short"));
    cellManager->setCellValue(TXCoordinate(row_t(2), column_t(1)), std::string("This is a very long text content"));
    
    // 自动调整列宽
    double newWidth = rowColManager->autoFitColumnWidth(column_t(1), *cellManager);
    
    // 新宽度应该大于默认宽度
    EXPECT_GT(newWidth, 8.43);
    EXPECT_EQ(rowColManager->getColumnWidth(column_t(1)), newWidth);
}

TEST_F(TXRowColumnManagerTest, AutoFitRowHeight) {
    // 设置一些有内容的单元格
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("Normal text"));
    
    // 自动调整行高
    double newHeight = rowColManager->autoFitRowHeight(row_t(1), *cellManager);
    
    // 新高度应该合理
    EXPECT_GE(newHeight, 12.0);
    EXPECT_LE(newHeight, 409.0);
    EXPECT_EQ(rowColManager->getRowHeight(row_t(1)), newHeight);
}

TEST_F(TXRowColumnManagerTest, AutoFitAllColumns) {
    // 设置多列内容
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("Column 1"));
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(2)), std::string("Very long column 2 content"));
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(3)), std::string("Col3"));
    
    // 自动调整所有列宽
    std::size_t adjustedCount = rowColManager->autoFitAllColumnWidths(*cellManager);
    
    EXPECT_EQ(adjustedCount, 3);
    
    // 验证列宽被调整
    EXPECT_GT(rowColManager->getColumnWidth(column_t(1)), 8.43);
    EXPECT_GT(rowColManager->getColumnWidth(column_t(2)), 8.43);
    EXPECT_GT(rowColManager->getColumnWidth(column_t(3)), 8.43);
}

TEST_F(TXRowColumnManagerTest, AutoFitAllRows) {
    // 设置多行内容
    cellManager->setCellValue(TXCoordinate(row_t(1), column_t(1)), std::string("Row 1"));
    cellManager->setCellValue(TXCoordinate(row_t(2), column_t(1)), std::string("Row 2"));
    cellManager->setCellValue(TXCoordinate(row_t(3), column_t(1)), std::string("Row 3"));
    
    // 自动调整所有行高
    std::size_t adjustedCount = rowColManager->autoFitAllRowHeights(*cellManager);
    
    EXPECT_EQ(adjustedCount, 3);
}

// ==================== 批量操作测试 ====================

TEST_F(TXRowColumnManagerTest, BatchSetRowHeights) {
    std::vector<std::pair<row_t, double>> heights = {
        {row_t(1), 20.0},
        {row_t(2), 25.0},
        {row_t(3), 30.0}
    };
    
    std::size_t count = rowColManager->setRowHeights(heights);
    EXPECT_EQ(count, 3);
    
    EXPECT_DOUBLE_EQ(rowColManager->getRowHeight(row_t(1)), 20.0);
    EXPECT_DOUBLE_EQ(rowColManager->getRowHeight(row_t(2)), 25.0);
    EXPECT_DOUBLE_EQ(rowColManager->getRowHeight(row_t(3)), 30.0);
}

TEST_F(TXRowColumnManagerTest, BatchSetColumnWidths) {
    std::vector<std::pair<column_t, double>> widths = {
        {column_t(1), 10.0},
        {column_t(2), 15.0},
        {column_t(3), 20.0}
    };
    
    std::size_t count = rowColManager->setColumnWidths(widths);
    EXPECT_EQ(count, 3);
    
    EXPECT_DOUBLE_EQ(rowColManager->getColumnWidth(column_t(1)), 10.0);
    EXPECT_DOUBLE_EQ(rowColManager->getColumnWidth(column_t(2)), 15.0);
    EXPECT_DOUBLE_EQ(rowColManager->getColumnWidth(column_t(3)), 20.0);

    // 生成测试文件
    addTestInfo(sheet, "BatchSetColumnWidths", "测试批量设置列宽功能");

    // 演示不同列宽的效果
    sheet->setCellValue(row_t(7), column_t(1), cell_value_t{"列"});
    sheet->setCellValue(row_t(7), column_t(2), cell_value_t{"宽度"});
    sheet->setCellValue(row_t(7), column_t(3), cell_value_t{"内容示例"});

    // 设置实际的列宽并添加内容
    sheet->setColumnWidth(column_t(1), 10.0);
    sheet->setCellValue(row_t(8), column_t(1), cell_value_t{"A"});
    sheet->setCellValue(row_t(8), column_t(2), cell_value_t{"10.0"});
    sheet->setCellValue(row_t(8), column_t(3), cell_value_t{"窄列内容"});

    sheet->setColumnWidth(column_t(2), 15.0);
    sheet->setCellValue(row_t(9), column_t(1), cell_value_t{"B"});
    sheet->setCellValue(row_t(9), column_t(2), cell_value_t{"15.0"});
    sheet->setCellValue(row_t(9), column_t(3), cell_value_t{"中等宽度列内容"});

    sheet->setColumnWidth(column_t(3), 20.0);
    sheet->setCellValue(row_t(10), column_t(1), cell_value_t{"C"});
    sheet->setCellValue(row_t(10), column_t(2), cell_value_t{"20.0"});
    sheet->setCellValue(row_t(10), column_t(3), cell_value_t{"较宽列可以容纳更多内容"});

    // 添加批量操作统计
    sheet->setCellValue(row_t(12), column_t(1), cell_value_t{"批量操作统计:"});
    sheet->setCellValue(row_t(12), column_t(2), cell_value_t{"成功设置"});
    sheet->setCellValue(row_t(12), column_t(3), cell_value_t{static_cast<double>(count)});
    sheet->setCellValue(row_t(12), column_t(4), cell_value_t{"列的宽度"});

    saveWorkbook(workbook, "BatchSetColumnWidths");
}

// ==================== 边界条件测试 ====================

TEST_F(TXRowColumnManagerTest, InvalidOperations) {
    // 测试无效行号
    EXPECT_FALSE(rowColManager->insertRows(row_t(0), row_t(1), *cellManager));
    EXPECT_FALSE(rowColManager->deleteRows(row_t(0), row_t(1), *cellManager));
    EXPECT_FALSE(rowColManager->setRowHeight(row_t(0), 20.0));
    
    // 测试无效列号
    EXPECT_FALSE(rowColManager->insertColumns(column_t(static_cast<column_t::index_t>(0)), column_t(static_cast<column_t::index_t>(1)), *cellManager));
    EXPECT_FALSE(rowColManager->deleteColumns(column_t(static_cast<column_t::index_t>(0)), column_t(static_cast<column_t::index_t>(1)), *cellManager));
    EXPECT_FALSE(rowColManager->setColumnWidth(column_t(static_cast<column_t::index_t>(0)), 20.0));

    // 测试零计数
    EXPECT_FALSE(rowColManager->insertRows(row_t(1), row_t(static_cast<row_t::index_t>(0)), *cellManager));
    EXPECT_FALSE(rowColManager->deleteRows(row_t(1), row_t(static_cast<row_t::index_t>(0)), *cellManager));
    EXPECT_FALSE(rowColManager->insertColumns(column_t(1), column_t(static_cast<column_t::index_t>(0)), *cellManager));
    EXPECT_FALSE(rowColManager->deleteColumns(column_t(1), column_t(static_cast<column_t::index_t>(0)), *cellManager));
}

// ==================== 清空操作测试 ====================

TEST_F(TXRowColumnManagerTest, Clear) {
    // 设置一些行高和列宽
    rowColManager->setRowHeight(row_t(1), 25.0);
    rowColManager->setColumnWidth(column_t(1), 15.0);
    rowColManager->setRowHidden(row_t(2), true);
    rowColManager->setColumnHidden(column_t(2), true);
    
    // 验证设置成功
    EXPECT_DOUBLE_EQ(rowColManager->getRowHeight(row_t(1)), 25.0);
    EXPECT_DOUBLE_EQ(rowColManager->getColumnWidth(column_t(1)), 15.0);
    EXPECT_TRUE(rowColManager->isRowHidden(row_t(2)));
    EXPECT_TRUE(rowColManager->isColumnHidden(column_t(2)));
    
    // 清空
    rowColManager->clear();
    
    // 验证恢复默认值
    EXPECT_DOUBLE_EQ(rowColManager->getRowHeight(row_t(1)), 15.0);
    EXPECT_DOUBLE_EQ(rowColManager->getColumnWidth(column_t(1)), 8.43);
    EXPECT_FALSE(rowColManager->isRowHidden(row_t(2)));
    EXPECT_FALSE(rowColManager->isColumnHidden(column_t(2)));
}

// ==================== 查询方法测试 ====================

TEST_F(TXRowColumnManagerTest, QueryMethods) {
    // 设置一些自定义尺寸
    rowColManager->setRowHeight(row_t(1), 20.0);
    rowColManager->setRowHeight(row_t(3), 30.0);
    rowColManager->setColumnWidth(column_t(2), 12.0);
    rowColManager->setColumnWidth(column_t(4), 18.0);
    
    // 获取自定义行高
    const auto& customRowHeights = rowColManager->getCustomRowHeights();
    EXPECT_EQ(customRowHeights.size(), 2);
    EXPECT_EQ(customRowHeights.at(1), 20.0);
    EXPECT_EQ(customRowHeights.at(3), 30.0);
    
    // 获取自定义列宽
    const auto& customColumnWidths = rowColManager->getCustomColumnWidths();
    EXPECT_EQ(customColumnWidths.size(), 2);
    EXPECT_EQ(customColumnWidths.at(2), 12.0);
    EXPECT_EQ(customColumnWidths.at(4), 18.0);
}
