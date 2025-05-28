//
// Merged Cells Test
// 测试合并单元格功能
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXMergedCells.hpp"
#include "TinaXlsx/TXRange.hpp"
#include <filesystem>
#include <iostream>

class MergedCellsTest : public ::testing::Test {
protected:
    void SetUp() override {
        std::filesystem::create_directories("output");
    }

    void TearDown() override {
        std::filesystem::remove("output/merged_cells_test.xlsx");
    }
};

/**
 * @brief 测试基本合并单元格功能
 */
TEST_F(MergedCellsTest, BasicMergeOperations) {
    std::cout << "\n=== 基础合并单元格测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("合并测试");
    ASSERT_NE(sheet, nullptr);
    
    // 设置标题
    sheet->setCellValue("A1", std::string("销售报表"));
    
    // 合并标题单元格 A1:D1
    bool success = sheet->mergeCells(TinaXlsx::row_t(1), TinaXlsx::column_t(1), 
                                    TinaXlsx::row_t(1), TinaXlsx::column_t(4));
    EXPECT_TRUE(success);
    
    // 检查合并状态
    EXPECT_TRUE(sheet->isCellMerged(TinaXlsx::row_t(1), TinaXlsx::column_t(1)));
    EXPECT_TRUE(sheet->isCellMerged(TinaXlsx::row_t(1), TinaXlsx::column_t(2))); // B1也应该被合并
    EXPECT_TRUE(sheet->isCellMerged(TinaXlsx::row_t(1), TinaXlsx::column_t(3))); // C1也应该被合并
    EXPECT_TRUE(sheet->isCellMerged(TinaXlsx::row_t(1), TinaXlsx::column_t(4))); // D1也应该被合并
    
    // 检查非合并区域
    EXPECT_FALSE(sheet->isCellMerged(TinaXlsx::row_t(2), TinaXlsx::column_t(1)));
    
    // 获取合并区域
    auto mergeRegion = sheet->getMergeRegion(TinaXlsx::row_t(1), TinaXlsx::column_t(1));
    EXPECT_TRUE(mergeRegion.isValid());
    EXPECT_EQ("A1:D1", mergeRegion.toAddress());
    
    std::cout << "基础合并单元格测试通过！\n";
}

/**
 * @brief 测试多个合并区域
 */
TEST_F(MergedCellsTest, MultipleMergeRegions) {
    std::cout << "\n=== 多个合并区域测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("多合并区域");
    ASSERT_NE(sheet, nullptr);
    
    // 第一个合并区域：A1:C1
    sheet->setCellValue("A1", std::string("标题1"));
    EXPECT_TRUE(sheet->mergeCells("A1:C1"));
    
    // 第二个合并区域：A3:B4  
    sheet->setCellValue("A3", std::string("合并区域2"));
    EXPECT_TRUE(sheet->mergeCells(TinaXlsx::row_t(3), TinaXlsx::column_t(1),
                                 TinaXlsx::row_t(4), TinaXlsx::column_t(2)));
    
    // 第三个合并区域：D3:F3
    sheet->setCellValue("D3", std::string("标题3"));
    TinaXlsx::TXRange range3(TinaXlsx::TXCoordinate(TinaXlsx::row_t(3),TinaXlsx::column_t(4)),
                            TinaXlsx::TXCoordinate(TinaXlsx::row_t(3),TinaXlsx::column_t(6)));
    EXPECT_TRUE(sheet->mergeCells(range3));
    
    // 验证所有合并区域
    auto allRegions = sheet->getAllMergeRegions();
    EXPECT_EQ(3, allRegions.size());
    EXPECT_EQ(3, sheet->getMergeCount());
    
    // 验证每个区域
    EXPECT_TRUE(sheet->isCellMerged(TinaXlsx::row_t(1), TinaXlsx::column_t(2))); // B1在第一个区域
    EXPECT_TRUE(sheet->isCellMerged(TinaXlsx::row_t(4), TinaXlsx::column_t(1))); // A4在第二个区域
    EXPECT_TRUE(sheet->isCellMerged(TinaXlsx::row_t(3), TinaXlsx::column_t(5))); // E3在第三个区域
    
    // 保存文件
    bool saved = workbook.saveToFile("output/merged_cells_test.xlsx");
    EXPECT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    
    std::cout << "多个合并区域测试通过！\n";
}

/**
 * @brief 测试合并单元格拆分
 */
TEST_F(MergedCellsTest, UnmergeCells) {
    std::cout << "\n=== 合并单元格拆分测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("拆分测试");
    ASSERT_NE(sheet, nullptr);
    
    // 创建几个合并区域
    sheet->setCellValue("A1", std::string("区域1"));
    EXPECT_TRUE(sheet->mergeCells("A1:C1"));
    
    sheet->setCellValue("A3", std::string("区域2"));
    EXPECT_TRUE(sheet->mergeCells("A3:B4"));
    
    sheet->setCellValue("D1", std::string("区域3"));
    EXPECT_TRUE(sheet->mergeCells("D1:F2"));
    
    // 验证初始状态
    EXPECT_EQ(3, sheet->getMergeCount());
    
    // 拆分第一个区域
    EXPECT_TRUE(sheet->unmergeCells(TinaXlsx::row_t(1), TinaXlsx::column_t(2))); // 使用区域内任意单元格
    EXPECT_EQ(2, sheet->getMergeCount());
    EXPECT_FALSE(sheet->isCellMerged(TinaXlsx::row_t(1), TinaXlsx::column_t(1)));
    
    // 拆分一个范围内的所有合并区域
    
    TinaXlsx::TXRange range(TinaXlsx::TXCoordinate(TinaXlsx::row_t(1),TinaXlsx::column_t(1)),
                        TinaXlsx::TXCoordinate(TinaXlsx::row_t(5),TinaXlsx::column_t(5)));
    
    std::size_t unmerged_count = sheet->unmergeCellsInRange(range);
    EXPECT_EQ(1, unmerged_count); // 应该拆分了区域2
    EXPECT_EQ(1, sheet->getMergeCount()); // 只剩下区域3
    
    // 验证区域3仍然存在
    EXPECT_TRUE(sheet->isCellMerged(TinaXlsx::row_t(1), TinaXlsx::column_t(4)));
    
    std::cout << "合并单元格拆分测试通过！\n";
}

/**
 * @brief 测试合并单元格数据完整性
 */
TEST_F(MergedCellsTest, MergedCellsDataIntegrity) {
    std::cout << "\n=== 合并单元格数据完整性测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("数据完整性");
    ASSERT_NE(sheet, nullptr);
    
    // 在合并前设置数据
    sheet->setCellValue("A1", std::string("主单元格数据"));
    sheet->setCellValue("B1", std::string("将被合并的数据"));
    sheet->setCellValue("C1", 123.45);
    
    // 合并单元格
    EXPECT_TRUE(sheet->mergeCells("A1:C1"));
    
    // 验证主单元格数据保持不变
    auto main_value = sheet->getCellValue("A1");
    EXPECT_TRUE(std::holds_alternative<std::string>(main_value));
    if (std::holds_alternative<std::string>(main_value)) {
        EXPECT_EQ("主单元格数据", std::get<std::string>(main_value));
    }
    
    // 创建实际报表结构
    sheet->setCellValue("A3", std::string("产品"));
    sheet->setCellValue("B3", std::string("Q1"));
    sheet->setCellValue("C3", std::string("Q2"));
    sheet->setCellValue("D3", std::string("总计"));
    
    // 添加数据行
    sheet->setCellValue("A4", std::string("产品A"));
    sheet->setCellValue("B4", 1000.0);
    sheet->setCellValue("C4", 1200.0);
    sheet->setCellValue("D4", 2200.0);
    
    sheet->setCellValue("A5", std::string("产品B"));
    sheet->setCellValue("B5", 800.0);
    sheet->setCellValue("C5", 900.0);
    sheet->setCellValue("D5", 1700.0);
    
    // 合并小计行
    sheet->setCellValue("A6", std::string("小计"));
    EXPECT_TRUE(sheet->mergeCells("A6:B6"));
    sheet->setCellValue("C6", 2100.0);
    sheet->setCellValue("D6", 3900.0);
    
    // 验证最终状态
    EXPECT_EQ(2, sheet->getMergeCount()); // A1:C1 和 A6:B6
    
    // 保存文件
    bool saved = workbook.saveToFile("output/merged_cells_test.xlsx");
    EXPECT_TRUE(saved);
    
    std::cout << "合并单元格数据完整性测试通过！\n";
}

/**
 * @brief 测试批量合并操作
 */
TEST_F(MergedCellsTest, BatchMergeOperations) {
    std::cout << "\n=== 批量合并操作测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("批量合并");
    ASSERT_NE(sheet, nullptr);
    
    // 创建一个表格结构，然后批量合并一些区域
    std::vector<TinaXlsx::TXMergedCells::MergeRegion> mergeRegions;
    
    // 标题行合并
    mergeRegions.emplace_back(TinaXlsx::row_t(1), TinaXlsx::column_t(1), 
                             TinaXlsx::row_t(1), TinaXlsx::column_t(6));
    
    // 每隔几行合并一些单元格
    for (int i = 3; i <= 10; i += 2) {
        mergeRegions.emplace_back(TinaXlsx::row_t(i), TinaXlsx::column_t(1),
                                 TinaXlsx::row_t(i), TinaXlsx::column_t(2));
    }
    
    // 使用TXMergedCells进行批量合并
    TinaXlsx::TXMergedCells mergedCells;
    std::size_t mergeCount = mergedCells.batchMergeCells(mergeRegions);
    EXPECT_EQ(mergeRegions.size(), mergeCount);
    
    // 设置一些数据来验证
    sheet->setCellValue("A1", std::string("批量合并测试报表"));
    sheet->setCellValue("A3", std::string("项目1"));
    sheet->setCellValue("A5", std::string("项目2"));
    sheet->setCellValue("A7", std::string("项目3"));
    sheet->setCellValue("A9", std::string("项目4"));
    
    // 手动应用合并到工作表
    for (const auto& region : mergeRegions) {
        EXPECT_TRUE(sheet->mergeCells(region.startRow, region.startCol, 
                                     region.endRow, region.endCol));
    }
    
    // 验证合并结果
    EXPECT_EQ(5, sheet->getMergeCount()); // 1个标题 + 4个项目行
    
    bool saved = workbook.saveToFile("output/merged_cells_test.xlsx");
    EXPECT_TRUE(saved);
    
    std::cout << "批量合并操作测试通过！\n";
} 
