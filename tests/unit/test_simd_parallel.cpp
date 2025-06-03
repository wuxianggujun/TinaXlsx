//
// @file test_simd_parallel.cpp
// @brief SIMD + 并行处理测试
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXSIMDOptimizations.hpp"
#include "TinaXlsx/TXSIMDParallelProcessor.hpp"
#include "TinaXlsx/TXUltraCompactCell.hpp"
#include <chrono>
#include <random>
#include <iostream>
#include <numeric>
#include <algorithm>

using namespace TinaXlsx;

class SIMDParallelTest : public ::testing::Test {
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
        
        test_doubles_.reserve(TEST_SIZE);
        test_int64s_.reserve(TEST_SIZE);
        test_rows_.reserve(TEST_SIZE);
        test_cols_.reserve(TEST_SIZE);
        
        for (size_t i = 0; i < TEST_SIZE; ++i) {
            test_doubles_.push_back(double_dist(gen));
            test_int64s_.push_back(int_dist(gen));
            test_rows_.push_back(static_cast<uint16_t>(i / 1000 + 1));
            test_cols_.push_back(static_cast<uint16_t>(i % 1000 + 1));
        }
    }
    
    std::vector<double> test_doubles_;
    std::vector<int64_t> test_int64s_;
    std::vector<uint16_t> test_rows_;
    std::vector<uint16_t> test_cols_;
};

// ==================== SIMD能力检测测试 ====================

TEST_F(SIMDParallelTest, SIMDCapabilityDetection) {
    std::cout << "SIMD能力检测:" << std::endl;
    std::cout << "  AVX2支持: " << (SIMDCapabilities::hasAVX2() ? "是" : "否") << std::endl;
    std::cout << "  SSE4.1支持: " << (SIMDCapabilities::hasSSE41() ? "是" : "否") << std::endl;
    std::cout << "  SSE2支持: " << (SIMDCapabilities::hasSSE2() ? "是" : "否") << std::endl;
    std::cout << "  最优批处理大小: " << SIMDCapabilities::getOptimalBatchSize() << std::endl;
    std::cout << "  SIMD类型: " << SIMDCapabilities::getSIMDInfo() << std::endl;
    
    // 基本验证
    EXPECT_TRUE(SIMDCapabilities::hasSSE2()); // 现代CPU都应该支持SSE2
    EXPECT_GT(SIMDCapabilities::getOptimalBatchSize(), 0);
}

// ==================== SIMD基础功能测试 ====================

TEST_F(SIMDParallelTest, SIMDBasicOperations) {
    const size_t TEST_SIZE = 1000;
    
    std::vector<double> input_doubles(TEST_SIZE);
    std::vector<UltraCompactCell> output_cells(TEST_SIZE);
    
    // 初始化测试数据
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        input_doubles[i] = static_cast<double>(i) * 3.14159;
    }
    
    // 测试SIMD转换
    auto start = std::chrono::high_resolution_clock::now();
    TXSIMDProcessor::convertDoublesToCells(input_doubles.data(), output_cells.data(), TEST_SIZE);
    auto end = std::chrono::high_resolution_clock::now();
    
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "SIMD转换 " << TEST_SIZE << " 个double: " << duration.count() << " 微秒" << std::endl;
    
    // 验证结果
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        EXPECT_EQ(output_cells[i].getType(), UltraCompactCell::CellType::Number);
        EXPECT_DOUBLE_EQ(output_cells[i].getNumberValue(), input_doubles[i]);
    }
}

TEST_F(SIMDParallelTest, SIMDMemoryOperations) {
    const size_t TEST_SIZE = 10000;
    
    std::vector<UltraCompactCell> cells(TEST_SIZE);
    std::vector<UltraCompactCell> copy_cells(TEST_SIZE);
    
    // 初始化数据
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        cells[i] = UltraCompactCell(static_cast<double>(i));
    }
    
    // 测试SIMD复制
    auto start = std::chrono::high_resolution_clock::now();
    TXSIMDProcessor::copyCells(cells.data(), copy_cells.data(), TEST_SIZE);
    auto end = std::chrono::high_resolution_clock::now();
    
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "SIMD复制 " << TEST_SIZE << " 个单元格: " << duration.count() << " 微秒" << std::endl;
    
    // 验证结果
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        EXPECT_EQ(copy_cells[i].getType(), cells[i].getType());
        EXPECT_DOUBLE_EQ(copy_cells[i].getNumberValue(), cells[i].getNumberValue());
    }
    
    // 测试SIMD清零
    start = std::chrono::high_resolution_clock::now();
    TXSIMDProcessor::clearCells(copy_cells.data(), TEST_SIZE);
    end = std::chrono::high_resolution_clock::now();
    
    duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "SIMD清零 " << TEST_SIZE << " 个单元格: " << duration.count() << " 微秒" << std::endl;
    
    // 验证结果
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        EXPECT_EQ(copy_cells[i].getType(), UltraCompactCell::CellType::Empty);
    }
}

TEST_F(SIMDParallelTest, SIMDNumericOperations) {
    const size_t TEST_SIZE = 50000;
    
    std::vector<UltraCompactCell> cells(TEST_SIZE);
    
    // 初始化数值数据
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        cells[i] = UltraCompactCell(static_cast<double>(i + 1));
    }
    
    // 测试SIMD求和
    auto start = std::chrono::high_resolution_clock::now();
    double sum = TXSIMDProcessor::sumNumbers(cells.data(), TEST_SIZE);
    auto end = std::chrono::high_resolution_clock::now();
    
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "SIMD求和 " << TEST_SIZE << " 个数值: " << duration.count() << " 微秒" << std::endl;
    
    // 验证结果（等差数列求和公式）
    double expected_sum = static_cast<double>(TEST_SIZE) * (TEST_SIZE + 1) / 2;
    EXPECT_NEAR(sum, expected_sum, 1e-6);
    
    std::cout << "求和结果: " << sum << ", 期望: " << expected_sum << std::endl;
}

// ==================== 并行处理器测试 ====================

TEST_F(SIMDParallelTest, ParallelProcessorBasic) {
    SIMDParallelConfig config;
    config.thread_count = 4;
    config.enable_simd = true;
    config.enable_parallel = true;
    
    TXSIMDParallelProcessor processor(config);
    
    // 测试并行转换
    std::vector<UltraCompactCell> output_cells;
    
    auto start = std::chrono::high_resolution_clock::now();
    processor.ultraFastConvertDoublesToCells(test_doubles_, output_cells);
    auto end = std::chrono::high_resolution_clock::now();
    
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "并行转换 " << test_doubles_.size() << " 个double: " << duration.count() << " 微秒" << std::endl;
    
    // 验证结果
    EXPECT_EQ(output_cells.size(), test_doubles_.size());
    for (size_t i = 0; i < test_doubles_.size(); ++i) {
        EXPECT_EQ(output_cells[i].getType(), UltraCompactCell::CellType::Number);
        EXPECT_DOUBLE_EQ(output_cells[i].getNumberValue(), test_doubles_[i]);
    }
    
    // 获取性能指标
    auto metrics = processor.getPerformanceMetrics();
    std::cout << "性能指标:" << std::endl;
    std::cout << "  总操作数: " << metrics.total_operations << std::endl;
    std::cout << "  平均时间: " << metrics.avg_time_per_operation_ns << " ns/操作" << std::endl;
    std::cout << "  操作数/秒: " << metrics.operations_per_second << std::endl;
    std::cout << "  SIMD类型: " << metrics.simd_type << std::endl;
    std::cout << "  线程数: " << metrics.thread_count << std::endl;
}

TEST_F(SIMDParallelTest, ParallelNumericOperations) {
    SIMDParallelConfig config;
    config.thread_count = std::thread::hardware_concurrency();
    config.enable_simd = true;
    config.enable_parallel = true;
    
    TXSIMDParallelProcessor processor(config);
    
    // 创建数值单元格
    std::vector<UltraCompactCell> cells;
    processor.ultraFastConvertDoublesToCells(test_doubles_, cells);
    
    // 测试并行求和
    auto start = std::chrono::high_resolution_clock::now();
    double sum = processor.ultraFastSumNumbers(cells);
    auto end = std::chrono::high_resolution_clock::now();
    
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "并行求和 " << cells.size() << " 个数值: " << duration.count() << " 微秒" << std::endl;
    
    // 验证结果
    double expected_sum = std::accumulate(test_doubles_.begin(), test_doubles_.end(), 0.0);
    EXPECT_NEAR(sum, expected_sum, 1e-6);
    
    // 测试并行统计
    start = std::chrono::high_resolution_clock::now();
    auto stats = processor.ultraFastCalculateStats(cells);
    end = std::chrono::high_resolution_clock::now();
    
    duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    std::cout << "并行统计 " << cells.size() << " 个数值: " << duration.count() << " 微秒" << std::endl;
    
    std::cout << "统计结果:" << std::endl;
    std::cout << "  数量: " << stats.count << std::endl;
    std::cout << "  求和: " << stats.sum << std::endl;
    std::cout << "  均值: " << stats.mean << std::endl;
    std::cout << "  最小值: " << stats.min << std::endl;
    std::cout << "  最大值: " << stats.max << std::endl;
    std::cout << "  标准差: " << stats.std_dev << std::endl;
    
    // 验证统计结果
    EXPECT_EQ(stats.count, test_doubles_.size());
    EXPECT_NEAR(stats.sum, expected_sum, 1e-6);
    EXPECT_GT(stats.mean, 0);
    EXPECT_GT(stats.max, stats.min);
}

// ==================== 性能对比测试 ====================

TEST_F(SIMDParallelTest, PerformanceComparison) {
    const size_t LARGE_SIZE = 1000000; // 100万个元素
    
    // 生成大量测试数据
    std::vector<double> large_doubles(LARGE_SIZE);
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_real_distribution<> dist(0.0, 1000.0);
    
    for (size_t i = 0; i < LARGE_SIZE; ++i) {
        large_doubles[i] = dist(gen);
    }
    
    std::cout << "\n=== 性能对比测试 (100万元素) ===" << std::endl;
    
    // 1. 标量处理
    std::vector<UltraCompactCell> scalar_output(LARGE_SIZE);
    auto start = std::chrono::high_resolution_clock::now();
    for (size_t i = 0; i < LARGE_SIZE; ++i) {
        scalar_output[i] = UltraCompactCell(large_doubles[i]);
    }
    auto end = std::chrono::high_resolution_clock::now();
    auto scalar_time = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "标量处理: " << scalar_time.count() << " 微秒" << std::endl;
    
    // 2. SIMD处理
    std::vector<UltraCompactCell> simd_output(LARGE_SIZE);
    start = std::chrono::high_resolution_clock::now();
    TXSIMDProcessor::convertDoublesToCells(large_doubles.data(), simd_output.data(), LARGE_SIZE);
    end = std::chrono::high_resolution_clock::now();
    auto simd_time = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "SIMD处理: " << simd_time.count() << " 微秒" << std::endl;
    
    // 3. 并行+SIMD处理
    SIMDParallelConfig config;
    config.enable_simd = true;
    config.enable_parallel = true;
    TXSIMDParallelProcessor processor(config);
    
    std::vector<UltraCompactCell> parallel_output;
    start = std::chrono::high_resolution_clock::now();
    processor.ultraFastConvertDoublesToCells(large_doubles, parallel_output);
    end = std::chrono::high_resolution_clock::now();
    auto parallel_time = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "并行+SIMD处理: " << parallel_time.count() << " 微秒" << std::endl;
    
    // 计算加速比
    double simd_speedup = static_cast<double>(scalar_time.count()) / simd_time.count();
    double parallel_speedup = static_cast<double>(scalar_time.count()) / parallel_time.count();
    
    std::cout << "\n加速比:" << std::endl;
    std::cout << "  SIMD vs 标量: " << simd_speedup << "x" << std::endl;
    std::cout << "  并行+SIMD vs 标量: " << parallel_speedup << "x" << std::endl;
    
    // 验证结果正确性
    for (size_t i = 0; i < 1000; ++i) { // 验证前1000个元素
        EXPECT_DOUBLE_EQ(scalar_output[i].getNumberValue(), large_doubles[i]);
        EXPECT_DOUBLE_EQ(simd_output[i].getNumberValue(), large_doubles[i]);
        EXPECT_DOUBLE_EQ(parallel_output[i].getNumberValue(), large_doubles[i]);
    }
    
    // 性能要求验证
    EXPECT_GT(simd_speedup, 1.0); // SIMD应该比标量快
    EXPECT_GT(parallel_speedup, simd_speedup); // 并行+SIMD应该最快
    
    std::cout << "\n性能目标验证:" << std::endl;
    std::cout << "  SIMD加速比 > 1.0: " << (simd_speedup > 1.0 ? "✓" : "✗") << std::endl;
    std::cout << "  并行加速比 > SIMD: " << (parallel_speedup > simd_speedup ? "✓" : "✗") << std::endl;
}
