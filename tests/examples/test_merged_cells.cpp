//
// Merged Cells Test
// 测试合并单元格功能
//

// tests/examples/test_merged_cells.cpp

#include <gtest/gtest.h>
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXMergedCells.hpp"
#include "TinaXlsx/TXRange.hpp"
#include <filesystem>
#include <iostream>

class MergedCellsTest : public ::testing::Test {
protected:
    // 将 workbook 声明为 static，以便在所有测试用例中共享
    // 或者在 SetUpTestSuite 中初始化，在 TearDownTestSuite 中保存和清理
    // 为简单起见，这里假设每个TEST_F创建自己的workbook并保存为不同文件或只测试逻辑

    // 如果希望所有 TEST_F 操作同一个 workbook，则需要更复杂的共享机制或在 TestSuite 级别处理
    // 当前的 Test Fixture 设计更适合每个 TEST_F 是独立的操作单元

    void SetUp() override {
        std::filesystem::create_directories("output");
        // 不要在 SetUp 中重新创建 workbook，除非每个 TEST_F 确实需要独立的 workbook
        // workbook = std::make_unique<TinaXlsx::TXWorkbook>(); // 如果每个测试用例独立
    }

    void TearDown() override {
        // 不要在这里保存文件，除非每个 TEST_F 的目标就是生成一个独立文件
        // if (workbook) {
        //     // workbook->saveToFile("output/merged_cells_test.xlsx"); // 移除或注释掉
        // }
        // 清理操作也应谨慎，确保不会影响其他测试
    }
    // std::unique_ptr<TinaXlsx::TXWorkbook> workbook; // 如果每个测试用例独立
};

// ... (你的 TEST_F 定义)

TEST_F(MergedCellsTest, BasicMergeOperations) {
    std::cout << "\n=== 基础合并单元格测试 ===\n";
    TinaXlsx::TXWorkbook workbook; // 每个测试创建自己的workbook实例
    auto* sheet = workbook.addSheet("合并测试_Basic");
    ASSERT_NE(sheet, nullptr);

    sheet->setCellValue("A1", std::string("销售报表"));
    bool success = sheet->mergeCells(TinaXlsx::row_t(1), TinaXlsx::column_t(1),
                                    TinaXlsx::row_t(1), TinaXlsx::column_t(4));
    EXPECT_TRUE(success);
    EXPECT_TRUE(sheet->isCellMerged(TinaXlsx::row_t(1), TinaXlsx::column_t(1)));

    // 为这个测试用例单独保存文件
    bool saved = workbook.saveToFile("output/merged_cells_basic_test.xlsx");
    EXPECT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    std::cout << "基础合并单元格测试完成！\n";
}

TEST_F(MergedCellsTest, MultipleMergeRegions) {
    std::cout << "\n=== 多个合并区域测试 ===\n";
    TinaXlsx::TXWorkbook workbook; // 每个测试创建自己的workbook实例
    auto* sheet = workbook.addSheet("多合并区域_Multiple");
    ASSERT_NE(sheet, nullptr);

    sheet->setCellValue("A1", std::string("标题1"));
    EXPECT_TRUE(sheet->mergeCells("A1:C1"));
    sheet->setCellValue("A3", std::string("合并区域2"));
    EXPECT_TRUE(sheet->mergeCells(TinaXlsx::row_t(3), TinaXlsx::column_t(1),
                                 TinaXlsx::row_t(4), TinaXlsx::column_t(2)));

    // 为这个测试用例单独保存文件
    bool saved = workbook.saveToFile("output/merged_cells_multiple_test.xlsx");
    EXPECT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    std::cout << "多个合并区域测试完成！\n";
}

TEST_F(MergedCellsTest, UnmergeCells) {
    std::cout << "\n=== 合并单元格拆分测试 ===\n";
    TinaXlsx::TXWorkbook workbook; // 每个测试创建自己的workbook实例
    auto* sheet = workbook.addSheet("拆分测试_Unmerge");
    ASSERT_NE(sheet, nullptr);

    sheet->setCellValue("A1", std::string("区域1"));
    EXPECT_TRUE(sheet->mergeCells("A1:C1"));
    EXPECT_TRUE(sheet->unmergeCells(TinaXlsx::row_t(1), TinaXlsx::column_t(2)));
    EXPECT_FALSE(sheet->isCellMerged(TinaXlsx::row_t(1), TinaXlsx::column_t(1)));

    // 为这个测试用例单独保存文件
    bool saved = workbook.saveToFile("output/merged_cells_unmerge_test.xlsx");
    EXPECT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    std::cout << "合并单元格拆分测试完成！\n";
}

TEST_F(MergedCellsTest, MergedCellsDataIntegrity) {
    std::cout << "\n=== 合并单元格数据完整性测试 ===\n";
    TinaXlsx::TXWorkbook workbook; // 每个测试创建自己的workbook实例
    auto* sheet = workbook.addSheet("数据完整性_Integrity");
    ASSERT_NE(sheet, nullptr);

    sheet->setCellValue("A1", std::string("主单元格数据"));
    sheet->setCellValue("B1", std::string("将被合并的数据"));
    EXPECT_TRUE(sheet->mergeCells("A1:C1"));
    auto main_value = sheet->getCellValue("A1");
    EXPECT_TRUE(std::holds_alternative<std::string>(main_value));
    if (std::holds_alternative<std::string>(main_value)) {
        EXPECT_EQ("主单元格数据", std::get<std::string>(main_value));
    }

    // 为这个测试用例单独保存文件
    bool saved = workbook.saveToFile("output/merged_cells_integrity_test.xlsx");
    EXPECT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    std::cout << "合并单元格数据完整性测试完成！\n";
}

TEST_F(MergedCellsTest, BatchMergeOperations) {
    std::cout << "\n=== 批量合并操作测试 ===\n";
    TinaXlsx::TXWorkbook workbook; // 每个测试创建自己的workbook实例
    auto* sheet = workbook.addSheet("批量合并_Batch");
    ASSERT_NE(sheet, nullptr);

    std::vector<TinaXlsx::TXMergedCells::MergeRegion> mergeRegions;
    mergeRegions.emplace_back(TinaXlsx::row_t(1), TinaXlsx::column_t(1),
                             TinaXlsx::row_t(1), TinaXlsx::column_t(6));
    sheet->setCellValue("A1", std::string("批量合并测试报表"));
     // 手动应用合并到工作表
    for (const auto& region : mergeRegions) {
        EXPECT_TRUE(sheet->mergeCells(region.startRow, region.startCol,
                                     region.endRow, region.endCol));
    }
    EXPECT_EQ(1, sheet->getMergeCount());

    // 为这个测试用例单独保存文件
    bool saved = workbook.saveToFile("output/merged_cells_batch_test.xlsx");
    EXPECT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    std::cout << "批量合并操作测试完成！\n";
}
