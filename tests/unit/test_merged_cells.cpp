#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <cstdio>

using namespace TinaXlsx;

class MergedCellsTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("MergeTest");
        ASSERT_NE(sheet, nullptr);
    }

    void TearDown() override {
        workbook.reset();
        // 清理测试文件
        std::remove("test_merged.xlsx");
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet;
};

// 测试基本合并功能
TEST_F(MergedCellsTest, BasicMerge) {
    // 设置主单元格的值
    sheet->setCellValue(row_t(1), column_t(1), std::string("Merged Cell"));
    
    // 合并A1:C3区域
    EXPECT_TRUE(sheet->mergeCells(row_t(1), column_t(1), row_t(3), column_t(3)));
    
    // 验证合并状态
    EXPECT_TRUE(sheet->isCellMerged(row_t(1), column_t(1))); // 主单元格
    EXPECT_TRUE(sheet->isCellMerged(row_t(2), column_t(2))); // 从属单元格
    EXPECT_TRUE(sheet->isCellMerged(row_t(3), column_t(3))); // 从属单元格
    
    // 验证合并区域
    auto mergeRegion = sheet->getMergeRegion(row_t(2), column_t(2));
    EXPECT_TRUE(mergeRegion.isValid());
    EXPECT_EQ(mergeRegion.getStart().getRow(), row_t(1));
    EXPECT_EQ(mergeRegion.getStart().getCol(), column_t(1));
    EXPECT_EQ(mergeRegion.getEnd().getRow(), row_t(3));
    EXPECT_EQ(mergeRegion.getEnd().getCol(), column_t(3));
    
    // 验证合并区域数量
    EXPECT_EQ(sheet->getMergeCount(), 1);
    
    // 保存文件验证
    EXPECT_TRUE(workbook->saveToFile("test_merged.xlsx"));
}

// 测试使用Range对象合并
TEST_F(MergedCellsTest, RangeMerge) {
    // 创建范围对象
    TXRange range(TXCoordinate(row_t(2), column_t(2)), TXCoordinate(row_t(4), column_t(5)));
    
    // 设置主单元格的值
    sheet->setCellValue(row_t(2), column_t(2), std::string("Range Merged"));
    
    // 使用Range合并
    EXPECT_TRUE(sheet->mergeCells(range));
    
    // 验证合并状态
    EXPECT_TRUE(sheet->isCellMerged(row_t(2), column_t(2)));
    EXPECT_TRUE(sheet->isCellMerged(row_t(3), column_t(4)));
    EXPECT_TRUE(sheet->isCellMerged(row_t(4), column_t(5)));
    
    // 验证合并区域
    auto mergeRegion = sheet->getMergeRegion(row_t(3), column_t(3));
    EXPECT_TRUE(mergeRegion.isValid());
    EXPECT_EQ(mergeRegion.getRowCount(), row_t(3)); // 2-4 = 3行
    EXPECT_EQ(mergeRegion.getColCount(), column_t(4)); // 2-5 = 4列
}

// 测试使用A1格式合并
TEST_F(MergedCellsTest, A1FormatMerge) {
    // 设置值
    sheet->setCellValue("B2", std::string("A1 Format Merge"));
    
    // 使用A1格式合并
    EXPECT_TRUE(sheet->mergeCells("B2:D4"));
    
    // 验证合并状态
    EXPECT_TRUE(sheet->isCellMerged(row_t(2), column_t(2))); // B2
    EXPECT_TRUE(sheet->isCellMerged(row_t(3), column_t(3))); // C3
    EXPECT_TRUE(sheet->isCellMerged(row_t(4), column_t(4))); // D4
    
    // 验证合并区域数量
    EXPECT_EQ(sheet->getMergeCount(), 1);
}

// 测试多个合并区域
TEST_F(MergedCellsTest, MultipleMerges) {
    // 第一个合并区域
    sheet->setCellValue(row_t(1), column_t(1), std::string("First Merge"));
    EXPECT_TRUE(sheet->mergeCells(row_t(1), column_t(1), row_t(2), column_t(2)));
    
    // 第二个合并区域
    sheet->setCellValue(row_t(4), column_t(4), std::string("Second Merge"));
    EXPECT_TRUE(sheet->mergeCells(row_t(4), column_t(4), row_t(5), column_t(6)));
    
    // 第三个合并区域
    sheet->setCellValue(row_t(1), column_t(8), std::string("Third Merge"));
    EXPECT_TRUE(sheet->mergeCells("H1:J2"));
    
    // 验证合并区域数量
    EXPECT_EQ(sheet->getMergeCount(), 3);
    
    // 获取所有合并区域
    auto allRegions = sheet->getAllMergeRegions();
    EXPECT_EQ(allRegions.size(), 3);
    
    // 验证各个区域的合并状态
    EXPECT_TRUE(sheet->isCellMerged(row_t(1), column_t(1)));
    EXPECT_TRUE(sheet->isCellMerged(row_t(5), column_t(5)));
    EXPECT_TRUE(sheet->isCellMerged(row_t(2), column_t(9)));
}

// 测试拆分合并区域
TEST_F(MergedCellsTest, UnmergeCells) {
    // 先合并一个区域
    sheet->setCellValue(row_t(1), column_t(1), std::string("To be unmerged"));
    EXPECT_TRUE(sheet->mergeCells(row_t(1), column_t(1), row_t(3), column_t(3)));
    
    // 验证合并状态
    EXPECT_TRUE(sheet->isCellMerged(row_t(2), column_t(2)));
    EXPECT_EQ(sheet->getMergeCount(), 1);
    
    // 拆分合并区域
    EXPECT_TRUE(sheet->unmergeCells(row_t(2), column_t(2))); // 使用从属单元格拆分
    
    // 验证拆分后状态
    EXPECT_FALSE(sheet->isCellMerged(row_t(1), column_t(1)));
    EXPECT_FALSE(sheet->isCellMerged(row_t(2), column_t(2)));
    EXPECT_FALSE(sheet->isCellMerged(row_t(3), column_t(3)));
    EXPECT_EQ(sheet->getMergeCount(), 0);
}

// 测试范围内拆分
TEST_F(MergedCellsTest, UnmergeInRange) {
    // 创建多个合并区域
    sheet->mergeCells(row_t(1), column_t(1), row_t(2), column_t(2)); // A1:B2
    sheet->mergeCells(row_t(3), column_t(1), row_t(4), column_t(2)); // A3:B4
    sheet->mergeCells(row_t(1), column_t(4), row_t(2), column_t(5)); // D1:E2
    sheet->mergeCells(row_t(6), column_t(6), row_t(7), column_t(7)); // F6:G7 (超出范围)
    
    EXPECT_EQ(sheet->getMergeCount(), 4);
    
    // 定义拆分范围 A1:C4
    TXRange unmergeRange(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(4), column_t(3)));
    
    // 拆分范围内的合并区域
    std::size_t unmergedCount = sheet->unmergeCellsInRange(unmergeRange);
    EXPECT_EQ(unmergedCount, 2); // 只有A1:B2和A3:B4在范围内
    
    // 验证剩余合并区域
    EXPECT_EQ(sheet->getMergeCount(), 2); // D1:E2和F6:G7应该还存在
    EXPECT_TRUE(sheet->isCellMerged(row_t(1), column_t(4))); // D1还是合并的
    EXPECT_TRUE(sheet->isCellMerged(row_t(6), column_t(6))); // F6还是合并的
    EXPECT_FALSE(sheet->isCellMerged(row_t(1), column_t(1))); // A1已拆分
    EXPECT_FALSE(sheet->isCellMerged(row_t(3), column_t(1))); // A3已拆分
}

// 测试重叠检测
TEST_F(MergedCellsTest, OverlapDetection) {
    // 先合并一个区域
    EXPECT_TRUE(sheet->mergeCells(row_t(1), column_t(1), row_t(3), column_t(3)));
    
    // 尝试合并重叠的区域（应该失败）
    EXPECT_FALSE(sheet->mergeCells(row_t(2), column_t(2), row_t(4), column_t(4))); // 重叠
    EXPECT_FALSE(sheet->mergeCells(row_t(1), column_t(1), row_t(2), column_t(2))); // 包含在内
    EXPECT_FALSE(sheet->mergeCells(row_t(1), column_t(1), row_t(4), column_t(4))); // 包含原区域
    
    // 合并不重叠的区域（应该成功）
    EXPECT_TRUE(sheet->mergeCells(row_t(5), column_t(5), row_t(6), column_t(6))); // 不重叠
    
    // 验证合并区域数量
    EXPECT_EQ(sheet->getMergeCount(), 2);
}

// 测试单元格的主从关系
TEST_F(MergedCellsTest, MasterSlaveRelationship) {
    // 合并区域
    sheet->setCellValue(row_t(1), column_t(1), std::string("Master Cell"));
    EXPECT_TRUE(sheet->mergeCells(row_t(1), column_t(1), row_t(2), column_t(3)));
    
    // 检查主单元格
    auto* masterCell = sheet->getCell(row_t(1), column_t(1));
    ASSERT_NE(masterCell, nullptr);
    EXPECT_TRUE(masterCell->isMerged());
    EXPECT_TRUE(masterCell->isMasterCell());
    
    // 检查从属单元格
    auto* slaveCell = sheet->getCell(row_t(2), column_t(2));
    ASSERT_NE(slaveCell, nullptr);
    EXPECT_TRUE(slaveCell->isMerged());
    EXPECT_FALSE(slaveCell->isMasterCell());
    
    // 检查主单元格位置
    auto masterPos = slaveCell->getMasterCellPosition();
    EXPECT_EQ(masterPos.first, 1);  // row
    EXPECT_EQ(masterPos.second, 1); // col
}

// 测试边界情况
TEST_F(MergedCellsTest, EdgeCases) {
    // 单个单元格"合并"（应该失败或被忽略）
    EXPECT_FALSE(sheet->mergeCells(row_t(1), column_t(1), row_t(1), column_t(1)));
    
    // 无效范围（应该失败）
    EXPECT_FALSE(sheet->mergeCells(row_t(3), column_t(3), row_t(1), column_t(1))); // 起始>结束
    
    // 在空单元格上拆分（应该失败）
    EXPECT_FALSE(sheet->unmergeCells(row_t(10), column_t(10)));
    
    // 验证没有创建任何合并区域
    EXPECT_EQ(sheet->getMergeCount(), 0);
} 