//
// @file test_xsimd_optimizations.cpp
// @brief 基于 xsimd 的SIMD优化测试
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXXSIMDOptimizations.hpp"
#include "TinaXlsx/TXUltraCompactCell.hpp"
#include <chrono>
#include <random>
#include <iostream>
#include <numeric>
#include <algorithm>

using namespace TinaXlsx;

class XSIMDOptimizationTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 初始化测试数据
        generateTestData();
    }
    
    void TearDown() override {
        // 清理
    }
    
    void generateTestData() {
        const size_t TEST_SIZE = 100000;
        
        // 生成随机数据
        std::random_device rd;
        std::mt19937 gen(rd());
        std::uniform_real_distribution<> double_dist(0.0, 1000.0);
        std::uniform_int_distribution<int64_t> int_dist(0, 1000000);
        std::uniform_int_distribution<int32_t> int32_dist(0, 100000);
        std::uniform_real_distribution<float> float_dist(0.0f, 1000.0f);
        
        test_doubles_.reserve(TEST_SIZE);
        test_int64s_.reserve(TEST_SIZE);
        test_int32s_.reserve(TEST_SIZE);
        test_floats_.reserve(TEST_SIZE);
        test_rows_.reserve(TEST_SIZE);
        test_cols_.reserve(TEST_SIZE);
        
        for (size_t i = 0; i < TEST_SIZE; ++i) {
            test_doubles_.push_back(double_dist(gen));
            test_int64s_.push_back(int_dist(gen));
            test_int32s_.push_back(int32_dist(gen));
            test_floats_.push_back(float_dist(gen));
            test_rows_.push_back(static_cast<uint16_t>(i / 1000 + 1));
            test_cols_.push_back(static_cast<uint16_t>(i % 1000 + 1));
        }
    }
    
    std::vector<double> test_doubles_;
    std::vector<int64_t> test_int64s_;
    std::vector<int32_t> test_int32s_;
    std::vector<float> test_floats_;
    std::vector<uint16_t> test_rows_;
    std::vector<uint16_t> test_cols_;
};

// ==================== xsimd能力检测测试 ====================

TEST_F(XSIMDOptimizationTest, XSIMDCapabilityDetection) {
    std::cout << "\n=== xsimd能力检测 ===" << std::endl;
    std::cout << XSIMDCapabilities::getSIMDArchInfo() << std::endl;
    std::cout << XSIMDCapabilities::getPerformanceInfo() << std::endl;
    
    // 基本验证
    EXPECT_GT(XSIMDCapabilities::getOptimalBatchSize(), 0);
    EXPECT_GT(XSIMDCapabilities::getSIMDRegisterSize(), 0);
    
    // 检查SIMD支持
    std::cout << "SIMD支持检测:" << std::endl;
    std::cout << "  - double: " << (XSIMDCapabilities::supportsSIMD<double>() ? "支持" : "不支持") << std::endl;
    std::cout << "  - float: " << (XSIMDCapabilities::supportsSIMD<float>() ? "支持" : "不支持") << std::endl;
    std::cout << "  - int64_t: " << (XSIMDCapabilities::supportsSIMD<int64_t>() ? "支持" : "不支持") << std::endl;
    std::cout << "  - int32_t: " << (XSIMDCapabilities::supportsSIMD<int32_t>() ? "支持" : "不支持") << std::endl;
}

// ==================== 数据类型转换测试 ====================

TEST_F(XSIMDOptimizationTest, DataTypeConversions) {
    const size_t TEST_SIZE = 10000;
    
    // 测试double转换
    std::vector<double> input_doubles(test_doubles_.begin(), test_doubles_.begin() + TEST_SIZE);
    std::vector<UltraCompactCell> output_cells;
    
    auto start = std::chrono::high_resolution_clock::now();
    TXXSIMDProcessor::convertDoublesToCells(input_doubles, output_cells);
    auto end = std::chrono::high_resolution_clock::now();
    
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "xsimd转换 " << TEST_SIZE << " 个double: " << duration.count() << " 微秒" << std::endl;
    
    // 验证结果
    EXPECT_EQ(output_cells.size(), input_doubles.size());
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        EXPECT_EQ(output_cells[i].getType(), UltraCompactCell::CellType::Number);
        EXPECT_DOUBLE_EQ(output_cells[i].getNumberValue(), input_doubles[i]);
    }
    
    // 测试int64转换
    std::vector<int64_t> input_int64s(test_int64s_.begin(), test_int64s_.begin() + TEST_SIZE);
    std::vector<UltraCompactCell> int64_cells;
    
    start = std::chrono::high_resolution_clock::now();
    TXXSIMDProcessor::convertInt64sToCells(input_int64s, int64_cells);
    end = std::chrono::high_resolution_clock::now();
    
    duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "xsimd转换 " << TEST_SIZE << " 个int64: " << duration.count() << " 微秒" << std::endl;
    
    // 验证结果
    EXPECT_EQ(int64_cells.size(), input_int64s.size());
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        EXPECT_EQ(int64_cells[i].getType(), UltraCompactCell::CellType::Integer);
        EXPECT_EQ(int64_cells[i].getIntegerValue(), input_int64s[i]);
    }
    
    // 测试float转换
    std::vector<float> input_floats(test_floats_.begin(), test_floats_.begin() + TEST_SIZE);
    std::vector<UltraCompactCell> float_cells;
    
    start = std::chrono::high_resolution_clock::now();
    TXXSIMDProcessor::convertFloatsToCells(input_floats, float_cells);
    end = std::chrono::high_resolution_clock::now();
    
    duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "xsimd转换 " << TEST_SIZE << " 个float: " << duration.count() << " 微秒" << std::endl;
    
    // 验证结果
    EXPECT_EQ(float_cells.size(), input_floats.size());
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        EXPECT_EQ(float_cells[i].getType(), UltraCompactCell::CellType::Number);
        EXPECT_NEAR(float_cells[i].getNumberValue(), static_cast<double>(input_floats[i]), 1e-6);
    }
}

// ==================== 内存操作测试 ====================

TEST_F(XSIMDOptimizationTest, MemoryOperations) {
    const size_t TEST_SIZE = 50000;
    
    std::vector<UltraCompactCell> cells;
    TXXSIMDProcessor::convertDoublesToCells(
        std::vector<double>(test_doubles_.begin(), test_doubles_.begin() + TEST_SIZE), 
        cells
    );
    
    // 测试复制
    std::vector<UltraCompactCell> copy_cells;
    auto start = std::chrono::high_resolution_clock::now();
    TXXSIMDProcessor::copyCells(cells, copy_cells);
    auto end = std::chrono::high_resolution_clock::now();
    
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "xsimd复制 " << TEST_SIZE << " 个单元格: " << duration.count() << " 微秒" << std::endl;
    
    // 验证复制结果
    EXPECT_TRUE(TXXSIMDProcessor::compareCells(cells, copy_cells));
    
    // 测试清零
    start = std::chrono::high_resolution_clock::now();
    TXXSIMDProcessor::clearCells(copy_cells);
    end = std::chrono::high_resolution_clock::now();
    
    duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "xsimd清零 " << TEST_SIZE << " 个单元格: " << duration.count() << " 微秒" << std::endl;
    
    // 验证清零结果
    for (const auto& cell : copy_cells) {
        EXPECT_EQ(cell.getType(), UltraCompactCell::CellType::Empty);
    }
    
    // 测试填充
    UltraCompactCell fill_value(42.0);
    start = std::chrono::high_resolution_clock::now();
    TXXSIMDProcessor::fillCells(copy_cells, fill_value);
    end = std::chrono::high_resolution_clock::now();
    
    duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "xsimd填充 " << TEST_SIZE << " 个单元格: " << duration.count() << " 微秒" << std::endl;
    
    // 验证填充结果
    for (const auto& cell : copy_cells) {
        EXPECT_EQ(cell.getType(), UltraCompactCell::CellType::Number);
        EXPECT_DOUBLE_EQ(cell.getNumberValue(), 42.0);
    }
}

// ==================== 坐标操作测试 ====================

TEST_F(XSIMDOptimizationTest, CoordinateOperations) {
    const size_t TEST_SIZE = 30000;
    
    std::vector<UltraCompactCell> cells;
    TXXSIMDProcessor::convertDoublesToCells(
        std::vector<double>(test_doubles_.begin(), test_doubles_.begin() + TEST_SIZE), 
        cells
    );
    
    std::vector<uint16_t> rows(test_rows_.begin(), test_rows_.begin() + TEST_SIZE);
    std::vector<uint16_t> cols(test_cols_.begin(), test_cols_.begin() + TEST_SIZE);
    
    // 测试设置坐标
    auto start = std::chrono::high_resolution_clock::now();
    TXXSIMDProcessor::setCoordinates(cells, rows, cols);
    auto end = std::chrono::high_resolution_clock::now();
    
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "xsimd设置坐标 " << TEST_SIZE << " 个单元格: " << duration.count() << " 微秒" << std::endl;
    
    // 验证坐标设置
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        EXPECT_EQ(cells[i].getRow(), rows[i]);
        EXPECT_EQ(cells[i].getCol(), cols[i]);
    }
    
    // 测试获取坐标
    std::vector<uint16_t> out_rows, out_cols;
    start = std::chrono::high_resolution_clock::now();
    TXXSIMDProcessor::getCoordinates(cells, out_rows, out_cols);
    end = std::chrono::high_resolution_clock::now();
    
    duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "xsimd获取坐标 " << TEST_SIZE << " 个单元格: " << duration.count() << " 微秒" << std::endl;
    
    // 验证坐标获取
    EXPECT_EQ(out_rows, rows);
    EXPECT_EQ(out_cols, cols);
    
    // 测试坐标变换
    start = std::chrono::high_resolution_clock::now();
    TXXSIMDProcessor::transformCoordinates(cells, 10, 5);
    end = std::chrono::high_resolution_clock::now();
    
    duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "xsimd变换坐标 " << TEST_SIZE << " 个单元格: " << duration.count() << " 微秒" << std::endl;
    
    // 验证坐标变换
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        EXPECT_EQ(cells[i].getRow(), rows[i] + 10);
        EXPECT_EQ(cells[i].getCol(), cols[i] + 5);
    }
}

// ==================== 数值计算测试 ====================

TEST_F(XSIMDOptimizationTest, NumericOperations) {
    const size_t TEST_SIZE = 100000;
    
    std::vector<UltraCompactCell> cells;
    TXXSIMDProcessor::convertDoublesToCells(
        std::vector<double>(test_doubles_.begin(), test_doubles_.begin() + TEST_SIZE), 
        cells
    );
    
    // 测试求和
    auto start = std::chrono::high_resolution_clock::now();
    double sum = TXXSIMDProcessor::sumNumbers(cells);
    auto end = std::chrono::high_resolution_clock::now();
    
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "xsimd求和 " << TEST_SIZE << " 个数值: " << duration.count() << " 微秒" << std::endl;
    
    // 验证求和结果
    double expected_sum = std::accumulate(test_doubles_.begin(), test_doubles_.begin() + TEST_SIZE, 0.0);
    EXPECT_NEAR(sum, expected_sum, 1e-6);
    
    std::cout << "求和结果: " << sum << ", 期望: " << expected_sum << std::endl;
    
    // 测试统计
    start = std::chrono::high_resolution_clock::now();
    auto stats = TXXSIMDProcessor::calculateStats(cells);
    end = std::chrono::high_resolution_clock::now();
    
    duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "xsimd统计 " << TEST_SIZE << " 个数值: " << duration.count() << " 微秒" << std::endl;
    
    std::cout << "统计结果:" << std::endl;
    std::cout << "  数量: " << stats.count << std::endl;
    std::cout << "  求和: " << stats.sum << std::endl;
    std::cout << "  均值: " << stats.mean << std::endl;
    std::cout << "  最小值: " << stats.min << std::endl;
    std::cout << "  最大值: " << stats.max << std::endl;
    std::cout << "  方差: " << stats.variance << std::endl;
    std::cout << "  标准差: " << stats.std_dev << std::endl;
    
    // 验证统计结果
    EXPECT_EQ(stats.count, TEST_SIZE);
    EXPECT_NEAR(stats.sum, expected_sum, 1e-6);
    EXPECT_GT(stats.mean, 0);
    EXPECT_GT(stats.max, stats.min);
    EXPECT_GT(stats.std_dev, 0);
    
    // 测试标量运算
    std::vector<UltraCompactCell> result_cells;
    start = std::chrono::high_resolution_clock::now();
    TXXSIMDProcessor::scalarOperation(cells, 2.0, result_cells, '*');
    end = std::chrono::high_resolution_clock::now();
    
    duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "xsimd标量乘法 " << TEST_SIZE << " 个数值: " << duration.count() << " 微秒" << std::endl;
    
    // 验证标量运算结果
    for (size_t i = 0; i < 1000; ++i) { // 验证前1000个元素
        EXPECT_NEAR(result_cells[i].getNumberValue(), test_doubles_[i] * 2.0, 1e-6);
    }
}

// ==================== 性能基准测试 ====================

TEST_F(XSIMDOptimizationTest, PerformanceBenchmark) {
    const size_t LARGE_SIZE = 1000000; // 100万个元素
    
    std::cout << "\n=== xsimd性能基准测试 (100万元素) ===" << std::endl;
    
    // 运行转换基准测试
    auto convert_result = TXXSIMDProcessor::benchmarkSIMD("convert", LARGE_SIZE);
    
    std::cout << "转换性能测试:" << std::endl;
    std::cout << "  xsimd时间: " << convert_result.xsimd_time_ms << " ms" << std::endl;
    std::cout << "  标量时间: " << convert_result.scalar_time_ms << " ms" << std::endl;
    std::cout << "  加速比: " << convert_result.speedup_ratio << "x" << std::endl;
    std::cout << "  操作数/秒: " << convert_result.operations_per_second << std::endl;
    
    // 运行求和基准测试
    auto sum_result = TXXSIMDProcessor::benchmarkSIMD("sum", LARGE_SIZE);
    
    std::cout << "\n求和性能测试:" << std::endl;
    std::cout << "  xsimd时间: " << sum_result.xsimd_time_ms << " ms" << std::endl;
    std::cout << "  标量时间: " << sum_result.scalar_time_ms << " ms" << std::endl;
    std::cout << "  加速比: " << sum_result.speedup_ratio << "x" << std::endl;
    std::cout << "  操作数/秒: " << sum_result.operations_per_second << std::endl;
    
    // 性能要求验证
    EXPECT_GT(convert_result.speedup_ratio, 1.0); // xsimd应该比标量快
    EXPECT_GT(sum_result.speedup_ratio, 1.0); // xsimd应该比标量快
    
    std::cout << "\n性能目标验证:" << std::endl;
    std::cout << "  转换加速比 > 1.0: " << (convert_result.speedup_ratio > 1.0 ? "✓" : "✗") << std::endl;
    std::cout << "  求和加速比 > 1.0: " << (sum_result.speedup_ratio > 1.0 ? "✓" : "✗") << std::endl;
    
    // 输出架构信息
    std::cout << "\n架构信息:" << std::endl;
    std::cout << convert_result.arch_info << std::endl;
}
