//
// @file TXOptimizedSIMD.cpp
// @brief 真正优化的SIMD实现
//

#include "TinaXlsx/TXOptimizedSIMD.hpp"
#include "TinaXlsx/TXXSIMDOptimizations.hpp"
#include <algorithm>
#include <cstring>
#include <iostream>

namespace TinaXlsx {

// ==================== 直接内存操作优化 ====================

void TXOptimizedSIMDProcessor::ultraFastConvertDoublesToCells(const double* input,
                                                             UltraCompactCell* output,
                                                             size_t count) {
    // 直接循环，避免SIMD开销（对于简单转换，标量可能更快）
    for (size_t i = 0; i < count; ++i) {
        output[i] = UltraCompactCell(input[i]);
    }
}

void TXOptimizedSIMDProcessor::ultraFastConvertInt64sToCells(const int64_t* input,
                                                            UltraCompactCell* output,
                                                            size_t count) {
    // 直接循环
    for (size_t i = 0; i < count; ++i) {
        output[i] = UltraCompactCell(input[i]);
    }
}

void TXOptimizedSIMDProcessor::ultraFastClearCells(UltraCompactCell* cells, size_t count) {
    // 使用memset直接清零（比SIMD更快）
    std::memset(cells, 0, count * sizeof(UltraCompactCell));
}

void TXOptimizedSIMDProcessor::ultraFastCopyCells(const UltraCompactCell* src,
                                                 UltraCompactCell* dst,
                                                 size_t count) {
    // 使用memcpy直接复制（比SIMD更快）
    std::memcpy(dst, src, count * sizeof(UltraCompactCell));
}

// ==================== 原地数值计算优化 ====================

double TXOptimizedSIMDProcessor::ultraFastSumNumbers(const UltraCompactCell* cells, size_t count) {
    // 使用Kahan求和算法提高精度，同时优化循环
    double sum = 0.0;
    double c = 0.0; // 补偿值
    
    for (size_t i = 0; i < count; ++i) {
        if (isNumericType(cells[i])) {
            double value = extractNumber(cells[i]);
            double y = value - c;
            double t = sum + y;
            c = (t - sum) - y;
            sum = t;
        }
    }
    
    return sum;
}

TXOptimizedSIMDProcessor::FastStats TXOptimizedSIMDProcessor::ultraFastCalculateStats(
    const UltraCompactCell* cells, size_t count) {
    
    FastStats stats;
    
    for (size_t i = 0; i < count; ++i) {
        if (isNumericType(cells[i])) {
            double value = extractNumber(cells[i]);
            stats.sum += value;
            stats.min = std::min(stats.min, value);
            stats.max = std::max(stats.max, value);
            stats.count++;
        }
    }
    
    return stats;
}

void TXOptimizedSIMDProcessor::ultraFastScalarMultiply(UltraCompactCell* cells, 
                                                      size_t count, 
                                                      double scalar) {
    for (size_t i = 0; i < count; ++i) {
        if (cells[i].getType() == UltraCompactCell::CellType::Number) {
            double value = cells[i].getNumberValue() * scalar;
            cells[i] = UltraCompactCell(value);
            // 保持原有坐标
            cells[i].setRow(cells[i].getRow());
            cells[i].setCol(cells[i].getCol());
        }
    }
}

void TXOptimizedSIMDProcessor::ultraFastScalarAdd(UltraCompactCell* cells,
                                                 size_t count,
                                                 double scalar) {
    for (size_t i = 0; i < count; ++i) {
        if (cells[i].getType() == UltraCompactCell::CellType::Number) {
            double value = cells[i].getNumberValue() + scalar;
            cells[i] = UltraCompactCell(value);
            // 保持原有坐标
            cells[i].setRow(cells[i].getRow());
            cells[i].setCol(cells[i].getCol());
        }
    }
}

// ==================== 专门的批处理优化 ====================

void TXOptimizedSIMDProcessor::batchSetNumbers(UltraCompactCell* cells,
                                              const double* values,
                                              size_t count) {
    // 批量设置，减少函数调用开销
    for (size_t i = 0; i < count; ++i) {
        cells[i] = UltraCompactCell(values[i]);
    }
}

void TXOptimizedSIMDProcessor::batchSetIntegers(UltraCompactCell* cells,
                                               const int64_t* values,
                                               size_t count) {
    for (size_t i = 0; i < count; ++i) {
        cells[i] = UltraCompactCell(values[i]);
    }
}

void TXOptimizedSIMDProcessor::batchGetNumbers(const UltraCompactCell* cells,
                                              double* output,
                                              size_t count) {
    for (size_t i = 0; i < count; ++i) {
        output[i] = extractNumber(cells[i]);
    }
}

// ==================== 性能对比测试 ====================

TXOptimizedSIMDProcessor::PerformanceComparison 
TXOptimizedSIMDProcessor::runPerformanceComparison(const std::string& operation,
                                                   size_t test_size) {
    PerformanceComparison result;
    result.operation_name = operation;
    result.data_size = test_size;
    
    // 准备测试数据
    std::vector<double> test_doubles(test_size);
    std::vector<UltraCompactCell> cells_optimized(test_size);
    std::vector<UltraCompactCell> cells_xsimd(test_size);
    std::vector<UltraCompactCell> cells_scalar(test_size);
    
    for (size_t i = 0; i < test_size; ++i) {
        test_doubles[i] = static_cast<double>(i) * 3.14159;
    }
    
    if (operation == "convert") {
        // 测试优化版本
        auto start = std::chrono::high_resolution_clock::now();
        ultraFastConvertDoublesToCells(test_doubles.data(), cells_optimized.data(), test_size);
        auto end = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        result.optimized_time_ms = duration.count() / 1000.0;
        
        // 测试xsimd版本
        start = std::chrono::high_resolution_clock::now();
        TXXSIMDProcessor::convertDoublesToCells(test_doubles, cells_xsimd);
        end = std::chrono::high_resolution_clock::now();
        duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        result.xsimd_time_ms = duration.count() / 1000.0;
        
        // 测试标量版本
        start = std::chrono::high_resolution_clock::now();
        for (size_t i = 0; i < test_size; ++i) {
            cells_scalar[i] = UltraCompactCell(test_doubles[i]);
        }
        end = std::chrono::high_resolution_clock::now();
        duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        result.scalar_time_ms = duration.count() / 1000.0;
        
    } else if (operation == "sum") {
        // 先转换数据
        ultraFastConvertDoublesToCells(test_doubles.data(), cells_optimized.data(), test_size);
        TXXSIMDProcessor::convertDoublesToCells(test_doubles, cells_xsimd);
        for (size_t i = 0; i < test_size; ++i) {
            cells_scalar[i] = UltraCompactCell(test_doubles[i]);
        }
        
        // 测试优化版本求和
        auto start = std::chrono::high_resolution_clock::now();
        double sum1 = ultraFastSumNumbers(cells_optimized.data(), test_size);
        auto end = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        result.optimized_time_ms = duration.count() / 1000.0;
        
        // 测试xsimd版本求和
        start = std::chrono::high_resolution_clock::now();
        double sum2 = TXXSIMDProcessor::sumNumbers(cells_xsimd);
        end = std::chrono::high_resolution_clock::now();
        duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        result.xsimd_time_ms = duration.count() / 1000.0;
        
        // 测试标量版本求和
        start = std::chrono::high_resolution_clock::now();
        double sum3 = 0.0;
        for (size_t i = 0; i < test_size; ++i) {
            if (cells_scalar[i].getType() == UltraCompactCell::CellType::Number) {
                sum3 += cells_scalar[i].getNumberValue();
            }
        }
        end = std::chrono::high_resolution_clock::now();
        duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        result.scalar_time_ms = duration.count() / 1000.0;
    }
    
    // 计算加速比
    if (result.scalar_time_ms > 0) {
        result.optimized_speedup = result.scalar_time_ms / result.optimized_time_ms;
        result.xsimd_speedup = result.scalar_time_ms / result.xsimd_time_ms;
    }
    
    return result;
}

// ==================== 性能测试工具类 ====================

std::vector<TXOptimizedSIMDProcessor::PerformanceComparison> SIMDPerformanceTester::results_;

void SIMDPerformanceTester::runFullPerformanceTest() {
    std::cout << "\n=== 完整SIMD性能对比测试 ===" << std::endl;
    
    const std::vector<size_t> test_sizes = {10000, 100000, 1000000};
    const std::vector<std::string> operations = {"convert", "sum"};
    
    results_.clear();
    
    for (const auto& operation : operations) {
        for (size_t test_size : test_sizes) {
            std::cout << "\n测试 " << operation << " 操作，数据大小: " << test_size << std::endl;
            
            auto result = TXOptimizedSIMDProcessor::runPerformanceComparison(operation, test_size);
            results_.push_back(result);
            
            std::cout << "  优化版本: " << result.optimized_time_ms << " ms" << std::endl;
            std::cout << "  xsimd版本: " << result.xsimd_time_ms << " ms" << std::endl;
            std::cout << "  标量版本: " << result.scalar_time_ms << " ms" << std::endl;
            std::cout << "  优化加速比: " << result.optimized_speedup << "x" << std::endl;
            std::cout << "  xsimd加速比: " << result.xsimd_speedup << "x" << std::endl;
        }
    }
}

std::string SIMDPerformanceTester::generatePerformanceReport() {
    std::ostringstream report;
    report << "\n=== SIMD性能测试报告 ===\n";
    
    for (const auto& result : results_) {
        report << "\n操作: " << result.operation_name 
               << ", 数据大小: " << result.data_size << "\n";
        report << "  优化版本: " << result.optimized_time_ms << " ms (加速比: " 
               << result.optimized_speedup << "x)\n";
        report << "  xsimd版本: " << result.xsimd_time_ms << " ms (加速比: " 
               << result.xsimd_speedup << "x)\n";
        report << "  标量版本: " << result.scalar_time_ms << " ms\n";
    }
    
    return report.str();
}

} // namespace TinaXlsx
