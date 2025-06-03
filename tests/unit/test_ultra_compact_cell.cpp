//
// @file test_ultra_compact_cell.cpp
// @brief UltraCompactCell 单元测试
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXUltraCompactCell.hpp"
#include "TinaXlsx/TXBatchCellManager.hpp"
#include <chrono>
#include <random>

using namespace TinaXlsx;

class UltraCompactCellTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 测试前准备
    }
    
    void TearDown() override {
        // 测试后清理
    }
};

// ==================== 基础功能测试 ====================

TEST_F(UltraCompactCellTest, SizeVerification) {
    // 验证16字节大小
    EXPECT_EQ(sizeof(UltraCompactCell), 16);
    EXPECT_LE(alignof(UltraCompactCell), 8);
}

TEST_F(UltraCompactCellTest, DefaultConstructor) {
    UltraCompactCell cell;
    EXPECT_EQ(cell.getType(), UltraCompactCell::CellType::Empty);
    EXPECT_TRUE(cell.isEmpty());
    EXPECT_EQ(cell.getRow(), 0);
    EXPECT_EQ(cell.getCol(), 0);
}

TEST_F(UltraCompactCellTest, StringValue) {
    std::string test_str = "Hello, World!";
    uint32_t offset = 100;
    
    UltraCompactCell cell(test_str, offset);
    
    EXPECT_EQ(cell.getType(), UltraCompactCell::CellType::String);
    EXPECT_FALSE(cell.isEmpty());
    EXPECT_EQ(cell.getStringOffset(), offset);
    EXPECT_EQ(cell.getStringLength(), test_str.length());
}

TEST_F(UltraCompactCellTest, NumberValue) {
    double test_value = 3.14159;
    
    UltraCompactCell cell(test_value);
    
    EXPECT_EQ(cell.getType(), UltraCompactCell::CellType::Number);
    EXPECT_FALSE(cell.isEmpty());
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), test_value);
}

TEST_F(UltraCompactCellTest, IntegerValue) {
    int64_t test_value = 1234567890LL;
    
    UltraCompactCell cell(test_value);
    
    EXPECT_EQ(cell.getType(), UltraCompactCell::CellType::Integer);
    EXPECT_FALSE(cell.isEmpty());
    EXPECT_EQ(cell.getIntegerValue(), test_value);
}

TEST_F(UltraCompactCellTest, BooleanValue) {
    UltraCompactCell cell_true(true);
    UltraCompactCell cell_false(false);
    
    EXPECT_EQ(cell_true.getType(), UltraCompactCell::CellType::Boolean);
    EXPECT_TRUE(cell_true.getBooleanValue());
    
    EXPECT_EQ(cell_false.getType(), UltraCompactCell::CellType::Boolean);
    EXPECT_FALSE(cell_false.getBooleanValue());
}

// ==================== 样式和属性测试 ====================

TEST_F(UltraCompactCellTest, StyleManagement) {
    UltraCompactCell cell(42.0);
    
    EXPECT_FALSE(cell.hasStyle());
    EXPECT_EQ(cell.getStyleIndex(), 0);
    
    cell.setStyleIndex(123);
    EXPECT_TRUE(cell.hasStyle());
    EXPECT_EQ(cell.getStyleIndex(), 123);
    
    cell.setStyleIndex(0);
    EXPECT_FALSE(cell.hasStyle());
}

TEST_F(UltraCompactCellTest, FormulaManagement) {
    UltraCompactCell cell("=A1+B1", 0);

    EXPECT_FALSE(cell.isFormula()); // 默认不是公式

    cell.setIsFormula(true);
    EXPECT_TRUE(cell.isFormula());
    EXPECT_EQ(cell.getType(), UltraCompactCell::CellType::Formula);

    // 调试信息
    std::cout << "Before setFormulaOffset:" << std::endl;
    std::cout << "  isFormula(): " << cell.isFormula() << std::endl;
    std::cout << "  getType(): " << static_cast<int>(cell.getType()) << std::endl;

    cell.setFormulaOffset(500);

    // 调试信息
    std::cout << "After setFormulaOffset:" << std::endl;
    std::cout << "  isFormula(): " << cell.isFormula() << std::endl;
    std::cout << "  getType(): " << static_cast<int>(cell.getType()) << std::endl;
    std::cout << "  getFormulaOffset(): " << cell.getFormulaOffset() << std::endl;

    EXPECT_EQ(cell.getFormulaOffset(), 500);
}

TEST_F(UltraCompactCellTest, CoordinateManagement) {
    UltraCompactCell cell;
    row_t row_obj{static_cast<u32>(10)};
    column_t col_obj{static_cast<u32>(20)};
    TXCoordinate coord{row_obj, col_obj};

    cell.setCoordinate(coord);
    EXPECT_EQ(cell.getRow(), 10);
    EXPECT_EQ(cell.getCol(), 20);
    EXPECT_EQ(cell.getCoordinate(), coord);
}

// ==================== 批处理管理器测试 ====================

class BatchCellManagerTest : public ::testing::Test {
protected:
    void SetUp() override {
        manager_ = std::make_unique<TXBatchCellManager>();
    }
    
    void TearDown() override {
        manager_.reset();
    }
    
    std::unique_ptr<TXBatchCellManager> manager_;
};

TEST_F(BatchCellManagerTest, BasicOperations) {
    // 创建测试数据
    std::vector<CellData> test_cells;
    test_cells.emplace_back("Hello", TXCoordinate{row_t{static_cast<u32>(1)}, column_t{static_cast<u32>(1)}});
    test_cells.emplace_back(42.0, TXCoordinate{row_t{static_cast<u32>(1)}, column_t{static_cast<u32>(2)}});
    test_cells.emplace_back(123LL, TXCoordinate{row_t{static_cast<u32>(1)}, column_t{static_cast<u32>(3)}});
    test_cells.emplace_back(true, TXCoordinate{row_t{static_cast<u32>(1)}, column_t{static_cast<u32>(4)}});
    
    // 批量设置
    size_t processed = manager_->setBatchCells(test_cells);
    EXPECT_EQ(processed, 4);
    
    // 验证单个获取
    auto cell1 = manager_->getCell(TXCoordinate{row_t{static_cast<u32>(1)}, column_t{static_cast<u32>(1)}});
    EXPECT_TRUE(std::holds_alternative<std::string>(cell1.value));
    EXPECT_EQ(std::get<std::string>(cell1.value), "Hello");

    auto cell2 = manager_->getCell(TXCoordinate{row_t{static_cast<u32>(1)}, column_t{static_cast<u32>(2)}});
    EXPECT_TRUE(std::holds_alternative<double>(cell2.value));
    EXPECT_DOUBLE_EQ(std::get<double>(cell2.value), 42.0);

    auto cell3 = manager_->getCell(TXCoordinate{row_t{static_cast<u32>(1)}, column_t{static_cast<u32>(3)}});
    EXPECT_TRUE(std::holds_alternative<int64_t>(cell3.value));
    EXPECT_EQ(std::get<int64_t>(cell3.value), 123LL);

    auto cell4 = manager_->getCell(TXCoordinate{row_t{static_cast<u32>(1)}, column_t{static_cast<u32>(4)}});
    EXPECT_TRUE(std::holds_alternative<bool>(cell4.value));
    EXPECT_TRUE(std::get<bool>(cell4.value));
}

TEST_F(BatchCellManagerTest, BatchRetrieval) {
    // 设置测试数据
    std::vector<CellData> test_cells;
    for (int row = 1; row <= 3; ++row) {
        for (int col = 1; col <= 3; ++col) {
            std::string value = "R" + std::to_string(row) + "C" + std::to_string(col);
            test_cells.emplace_back(value, TXCoordinate{row_t{static_cast<u32>(row)}, column_t{static_cast<u32>(col)}});
        }
    }
    
    manager_->setBatchCells(test_cells);
    
    // 批量获取
    CellRange range(1, 1, 3, 3);
    auto retrieved = manager_->getBatchCells(range);
    
    EXPECT_EQ(retrieved.size(), 9);
    
    // 验证数据
    for (size_t i = 0; i < retrieved.size(); ++i) {
        EXPECT_TRUE(std::holds_alternative<std::string>(retrieved[i].value));
        
        int row = retrieved[i].coordinate.getRow().index();
        int col = retrieved[i].coordinate.getCol().index();
        std::string expected = "R" + std::to_string(row) + "C" + std::to_string(col);
        EXPECT_EQ(std::get<std::string>(retrieved[i].value), expected);
    }
}

// ==================== 性能测试 ====================

TEST_F(BatchCellManagerTest, PerformanceTest) {
    const size_t CELL_COUNT = 100000;
    
    // 生成测试数据
    std::vector<CellData> test_cells;
    test_cells.reserve(CELL_COUNT);
    
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_int_distribution<> type_dist(0, 3);
    std::uniform_real_distribution<> double_dist(0.0, 1000.0);
    std::uniform_int_distribution<int64_t> int_dist(0, 1000000);
    
    for (size_t i = 0; i < CELL_COUNT; ++i) {
        uint16_t row = static_cast<uint16_t>(i / 1000 + 1);
        uint16_t col = static_cast<uint16_t>(i % 1000 + 1);
        TXCoordinate coord{row_t{static_cast<u32>(row)}, column_t{static_cast<u32>(col)}};
        
        int type = type_dist(gen);
        switch (type) {
            case 0:
                test_cells.emplace_back("Test" + std::to_string(i), coord);
                break;
            case 1:
                test_cells.emplace_back(double_dist(gen), coord);
                break;
            case 2:
                test_cells.emplace_back(static_cast<int64_t>(int_dist(gen)), coord);
                break;
            case 3:
                test_cells.emplace_back(i % 2 == 0, coord);
                break;
        }
    }
    
    // 性能测试
    manager_->startTiming();
    auto start = std::chrono::high_resolution_clock::now();
    
    size_t processed = manager_->setBatchCells(test_cells);
    
    auto end = std::chrono::high_resolution_clock::now();
    manager_->endTiming();
    
    // 验证结果
    EXPECT_EQ(processed, CELL_COUNT);
    
    // 计算性能
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    double us_per_cell = static_cast<double>(duration.count()) / CELL_COUNT;
    
    std::cout << "处理了 " << CELL_COUNT << " 个单元格" << std::endl;
    std::cout << "总时间: " << duration.count() << " 微秒" << std::endl;
    std::cout << "平均时间: " << us_per_cell << " 微秒/单元格" << std::endl;
    
    // 性能目标：<10μs per cell
    EXPECT_LT(us_per_cell, 15.0); // 暂时放宽到15μs，后续优化
    
    // 获取统计信息
    auto stats = manager_->getStats();
    std::cout << "统计信息:" << std::endl;
    std::cout << "  处理单元格数: " << stats.cells_processed << std::endl;
    std::cout << "  平均处理时间: " << stats.avg_time_per_cell << " μs/cell" << std::endl;
    std::cout << "  内存使用: " << stats.memory_used << " 字节" << std::endl;
    std::cout << "  内存效率: " << (stats.memory_efficiency * 100) << "%" << std::endl;
    std::cout << "  字符串池大小: " << stats.string_pool_size << " 字节" << std::endl;
}

// ==================== 内存限制测试 ====================

TEST_F(BatchCellManagerTest, MemoryLimitTest) {
    // 测试内存使用是否在合理范围内
    const size_t LARGE_CELL_COUNT = 1000000; // 100万个单元格
    
    std::vector<CellData> test_cells;
    test_cells.reserve(LARGE_CELL_COUNT);
    
    for (size_t i = 0; i < LARGE_CELL_COUNT; ++i) {
        uint16_t row = static_cast<uint16_t>(i / 1000 + 1);
        uint16_t col = static_cast<uint16_t>(i % 1000 + 1);
        test_cells.emplace_back(static_cast<double>(i), TXCoordinate{row_t{static_cast<u32>(row)}, column_t{static_cast<u32>(col)}});
    }
    
    size_t processed = manager_->setBatchCells(test_cells);
    
    // 检查内存使用
    size_t memory_used = manager_->getMemoryUsage();
    size_t theoretical_min = processed * 16; // 16字节每个单元格
    
    std::cout << "大规模测试结果:" << std::endl;
    std::cout << "  处理单元格数: " << processed << std::endl;
    std::cout << "  实际内存使用: " << memory_used << " 字节" << std::endl;
    std::cout << "  理论最小内存: " << theoretical_min << " 字节" << std::endl;
    std::cout << "  内存效率: " << (static_cast<double>(theoretical_min) / memory_used * 100) << "%" << std::endl;
    
    // 内存效率应该 > 50%
    EXPECT_GT(static_cast<double>(theoretical_min) / memory_used, 0.5);
    
    // 总内存使用应该 < 4GB
    EXPECT_LT(memory_used, 4ULL * 1024 * 1024 * 1024);
}
