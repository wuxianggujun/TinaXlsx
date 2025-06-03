//
// @file test_simd_comparison.cpp
// @brief SIMD实现对比测试 - 标量 vs xsimd vs 优化版本
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXOptimizedSIMD.hpp"
#include "TinaXlsx/TXXSIMDOptimizations.hpp"
#include "TinaXlsx/TXUltraCompactCell.hpp"
#include <chrono>
#include <random>
#include <iostream>
#include <numeric>

using namespace TinaXlsx;

class SIMDComparisonTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 生成测试数据
        const size_t TEST_SIZE = 100000;
        
        std::random_device rd;
        std::mt19937 gen(rd());
        std::uniform_real_distribution<> dist(0.0, 1000.0);
        
        test_doubles_.reserve(TEST_SIZE);
        for (size_t i = 0; i < TEST_SIZE; ++i) {
            test_doubles_.push_back(dist(gen));
        }
    }
    
    std::vector<double> test_doubles_;
};

// ==================== 转换性能对比测试 ====================

TEST_F(SIMDComparisonTest, ConversionPerformanceComparison) {
    const size_t TEST_SIZE = test_doubles_.size();
    
    std::cout << "\n=== 转换性能对比测试 (" << TEST_SIZE << " 元素) ===" << std::endl;
    
    // 准备输出容器
    std::vector<UltraCompactCell> scalar_output(TEST_SIZE);
    std::vector<UltraCompactCell> xsimd_output;
    std::vector<UltraCompactCell> optimized_output(TEST_SIZE);
    
    // 1. 标量版本测试
    auto start = std::chrono::high_resolution_clock::now();
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        scalar_output[i] = UltraCompactCell(test_doubles_[i]);
    }
    auto end = std::chrono::high_resolution_clock::now();
    auto scalar_time = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    // 2. xsimd版本测试
    start = std::chrono::high_resolution_clock::now();
    TXXSIMDProcessor::convertDoublesToCells(test_doubles_, xsimd_output);
    end = std::chrono::high_resolution_clock::now();
    auto xsimd_time = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    // 3. 优化版本测试
    start = std::chrono::high_resolution_clock::now();
    TXOptimizedSIMDProcessor::ultraFastConvertDoublesToCells(
        test_doubles_.data(), optimized_output.data(), TEST_SIZE);
    end = std::chrono::high_resolution_clock::now();
    auto optimized_time = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    // 输出结果
    std::cout << "标量版本:   " << scalar_time.count() << " 微秒" << std::endl;
    std::cout << "xsimd版本:  " << xsimd_time.count() << " 微秒" << std::endl;
    std::cout << "优化版本:   " << optimized_time.count() << " 微秒" << std::endl;
    
    // 计算加速比
    double xsimd_speedup = static_cast<double>(scalar_time.count()) / xsimd_time.count();
    double optimized_speedup = static_cast<double>(scalar_time.count()) / optimized_time.count();
    
    std::cout << "xsimd加速比:    " << xsimd_speedup << "x" << std::endl;
    std::cout << "优化版本加速比: " << optimized_speedup << "x" << std::endl;
    
    // 验证结果正确性
    for (size_t i = 0; i < 1000; ++i) { // 验证前1000个元素
        EXPECT_DOUBLE_EQ(scalar_output[i].getNumberValue(), test_doubles_[i]);
        EXPECT_DOUBLE_EQ(xsimd_output[i].getNumberValue(), test_doubles_[i]);
        EXPECT_DOUBLE_EQ(optimized_output[i].getNumberValue(), test_doubles_[i]);
    }
    
    // 性能断言（优化版本应该最快）
    EXPECT_LE(optimized_time.count(), scalar_time.count()) 
        << "优化版本应该不慢于标量版本";
}

// ==================== 求和性能对比测试 ====================

TEST_F(SIMDComparisonTest, SumPerformanceComparison) {
    const size_t TEST_SIZE = test_doubles_.size();
    
    std::cout << "\n=== 求和性能对比测试 (" << TEST_SIZE << " 元素) ===" << std::endl;
    
    // 准备测试数据
    std::vector<UltraCompactCell> cells(TEST_SIZE);
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        cells[i] = UltraCompactCell(test_doubles_[i]);
    }
    
    // 1. 标量版本测试
    auto start = std::chrono::high_resolution_clock::now();
    double scalar_sum = 0.0;
    for (const auto& cell : cells) {
        if (cell.getType() == UltraCompactCell::CellType::Number) {
            scalar_sum += cell.getNumberValue();
        }
    }
    auto end = std::chrono::high_resolution_clock::now();
    auto scalar_time = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    // 2. xsimd版本测试
    start = std::chrono::high_resolution_clock::now();
    double xsimd_sum = TXXSIMDProcessor::sumNumbers(cells);
    end = std::chrono::high_resolution_clock::now();
    auto xsimd_time = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    // 3. 优化版本测试
    start = std::chrono::high_resolution_clock::now();
    double optimized_sum = TXOptimizedSIMDProcessor::ultraFastSumNumbers(cells.data(), TEST_SIZE);
    end = std::chrono::high_resolution_clock::now();
    auto optimized_time = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    // 输出结果
    std::cout << "标量版本:   " << scalar_time.count() << " 微秒 (结果: " << scalar_sum << ")" << std::endl;
    std::cout << "xsimd版本:  " << xsimd_time.count() << " 微秒 (结果: " << xsimd_sum << ")" << std::endl;
    std::cout << "优化版本:   " << optimized_time.count() << " 微秒 (结果: " << optimized_sum << ")" << std::endl;
    
    // 计算加速比
    double xsimd_speedup = static_cast<double>(scalar_time.count()) / xsimd_time.count();
    double optimized_speedup = static_cast<double>(scalar_time.count()) / optimized_time.count();
    
    std::cout << "xsimd加速比:    " << xsimd_speedup << "x" << std::endl;
    std::cout << "优化版本加速比: " << optimized_speedup << "x" << std::endl;
    
    // 验证结果正确性
    double expected_sum = std::accumulate(test_doubles_.begin(), test_doubles_.end(), 0.0);
    EXPECT_NEAR(scalar_sum, expected_sum, 1e-6);
    EXPECT_NEAR(xsimd_sum, expected_sum, 1e-6);
    EXPECT_NEAR(optimized_sum, expected_sum, 1e-6);
    
    // 性能断言
    EXPECT_LE(optimized_time.count(), scalar_time.count()) 
        << "优化版本应该不慢于标量版本";
}

// ==================== 内存操作对比测试 ====================

TEST_F(SIMDComparisonTest, MemoryOperationsComparison) {
    const size_t TEST_SIZE = test_doubles_.size();
    
    std::cout << "\n=== 内存操作对比测试 (" << TEST_SIZE << " 元素) ===" << std::endl;
    
    // 准备测试数据
    std::vector<UltraCompactCell> source_cells(TEST_SIZE);
    for (size_t i = 0; i < TEST_SIZE; ++i) {
        source_cells[i] = UltraCompactCell(test_doubles_[i]);
    }
    
    // 测试清零操作
    {
        std::vector<UltraCompactCell> test_cells = source_cells;
        
        auto start = std::chrono::high_resolution_clock::now();
        TXOptimizedSIMDProcessor::ultraFastClearCells(test_cells.data(), TEST_SIZE);
        auto end = std::chrono::high_resolution_clock::now();
        auto clear_time = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        
        std::cout << "优化清零: " << clear_time.count() << " 微秒" << std::endl;
        
        // 验证清零结果
        for (size_t i = 0; i < 100; ++i) { // 验证前100个元素
            EXPECT_EQ(test_cells[i].getType(), UltraCompactCell::CellType::Empty);
        }
    }
    
    // 测试复制操作
    {
        std::vector<UltraCompactCell> dest_cells(TEST_SIZE);
        
        auto start = std::chrono::high_resolution_clock::now();
        TXOptimizedSIMDProcessor::ultraFastCopyCells(
            source_cells.data(), dest_cells.data(), TEST_SIZE);
        auto end = std::chrono::high_resolution_clock::now();
        auto copy_time = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        
        std::cout << "优化复制: " << copy_time.count() << " 微秒" << std::endl;
        
        // 验证复制结果
        for (size_t i = 0; i < 100; ++i) { // 验证前100个元素
            EXPECT_EQ(dest_cells[i], source_cells[i]);
        }
    }
}

// ==================== 完整性能测试套件 ====================

TEST_F(SIMDComparisonTest, ComprehensivePerformanceTest) {
    std::cout << "\n=== 运行完整性能测试套件 ===" << std::endl;
    
    SIMDPerformanceTester::runFullPerformanceTest();
    
    std::string report = SIMDPerformanceTester::generatePerformanceReport();
    std::cout << report << std::endl;
}

// ==================== 可扩展性测试 ====================

TEST_F(SIMDComparisonTest, ScalabilityTest) {
    std::cout << "\n=== 可扩展性测试 ===" << std::endl;
    
    const std::vector<size_t> test_sizes = {1000, 10000, 100000, 1000000};
    
    for (size_t size : test_sizes) {
        std::cout << "\n测试大小: " << size << " 元素" << std::endl;
        
        // 生成测试数据
        std::vector<double> data(size);
        for (size_t i = 0; i < size; ++i) {
            data[i] = static_cast<double>(i) * 3.14159;
        }
        
        // 测试转换性能
        std::vector<UltraCompactCell> output(size);
        
        auto start = std::chrono::high_resolution_clock::now();
        TXOptimizedSIMDProcessor::ultraFastConvertDoublesToCells(
            data.data(), output.data(), size);
        auto end = std::chrono::high_resolution_clock::now();
        auto time = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        
        double throughput = static_cast<double>(size) / (time.count() / 1000000.0);
        
        std::cout << "  转换时间: " << time.count() << " 微秒" << std::endl;
        std::cout << "  吞吐量: " << static_cast<size_t>(throughput) << " 元素/秒" << std::endl;
        std::cout << "  平均时间: " << static_cast<double>(time.count()) / size << " 微秒/元素" << std::endl;
    }
}
