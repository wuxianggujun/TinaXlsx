//
// @file TXOptimizedSIMD.hpp
// @brief 真正优化的SIMD实现 - 针对UltraCompactCell的专门优化
//

#pragma once

#include "TXUltraCompactCell.hpp"
#include <xsimd/xsimd.hpp>
#include <vector>
#include <chrono>

namespace TinaXlsx {

/**
 * @brief 真正优化的SIMD处理器
 * 
 * 这个版本专门针对UltraCompactCell的内存布局进行优化，
 * 避免不必要的数据提取和临时向量分配
 */
class TXOptimizedSIMDProcessor {
public:
    // ==================== 直接内存操作优化 ====================
    
    /**
     * @brief 超高速批量转换double到UltraCompactCell
     * 直接操作内存，避免临时向量
     */
    static void ultraFastConvertDoublesToCells(const double* input,
                                              UltraCompactCell* output,
                                              size_t count);
    
    /**
     * @brief 超高速批量转换int64到UltraCompactCell
     */
    static void ultraFastConvertInt64sToCells(const int64_t* input,
                                             UltraCompactCell* output,
                                             size_t count);
    
    /**
     * @brief 超高速内存清零
     * 使用SIMD直接清零16字节块
     */
    static void ultraFastClearCells(UltraCompactCell* cells, size_t count);
    
    /**
     * @brief 超高速内存复制
     * 使用SIMD批量复制16字节块
     */
    static void ultraFastCopyCells(const UltraCompactCell* src,
                                  UltraCompactCell* dst,
                                  size_t count);
    
    // ==================== 原地数值计算优化 ====================
    
    /**
     * @brief 超高速求和 - 直接从UltraCompactCell内存读取
     */
    static double ultraFastSumNumbers(const UltraCompactCell* cells, size_t count);
    
    /**
     * @brief 超高速统计计算
     */
    struct FastStats {
        double sum = 0.0;
        double min = std::numeric_limits<double>::infinity();
        double max = -std::numeric_limits<double>::infinity();
        size_t count = 0;
    };
    
    static FastStats ultraFastCalculateStats(const UltraCompactCell* cells, size_t count);
    
    /**
     * @brief 超高速标量运算 - 原地修改
     */
    static void ultraFastScalarMultiply(UltraCompactCell* cells, size_t count, double scalar);
    static void ultraFastScalarAdd(UltraCompactCell* cells, size_t count, double scalar);
    
    // ==================== 专门的批处理优化 ====================
    
    /**
     * @brief 批量设置数值 - 避免单个构造
     */
    static void batchSetNumbers(UltraCompactCell* cells, 
                               const double* values, 
                               size_t count);
    
    /**
     * @brief 批量设置整数
     */
    static void batchSetIntegers(UltraCompactCell* cells,
                                const int64_t* values,
                                size_t count);
    
    /**
     * @brief 批量提取数值 - 直接内存访问
     */
    static void batchGetNumbers(const UltraCompactCell* cells,
                               double* output,
                               size_t count);
    
    // ==================== 性能对比测试 ====================
    
    struct PerformanceComparison {
        double optimized_time_ms = 0.0;
        double xsimd_time_ms = 0.0;
        double scalar_time_ms = 0.0;
        double optimized_speedup = 0.0;
        double xsimd_speedup = 0.0;
        std::string operation_name;
        size_t data_size = 0;
    };
    
    /**
     * @brief 运行性能对比测试
     */
    static PerformanceComparison runPerformanceComparison(const std::string& operation,
                                                          size_t test_size = 1000000);

private:
    // ==================== 内部优化实现 ====================
    
    /**
     * @brief 检查UltraCompactCell是否为数值类型（内联优化）
     */
    static inline bool isNumericType(const UltraCompactCell& cell) {
        auto type = cell.getType();
        return type == UltraCompactCell::CellType::Number || 
               type == UltraCompactCell::CellType::Integer;
    }
    
    /**
     * @brief 快速提取数值（避免虚函数调用）
     */
    static inline double extractNumber(const UltraCompactCell& cell) {
        if (cell.getType() == UltraCompactCell::CellType::Number) {
            return cell.getNumberValue();
        } else if (cell.getType() == UltraCompactCell::CellType::Integer) {
            return static_cast<double>(cell.getIntegerValue());
        }
        return 0.0;
    }
    
    /**
     * @brief SIMD优化的数值提取
     */
    static void simdExtractNumbers(const UltraCompactCell* cells,
                                  double* output,
                                  size_t count);
    
    /**
     * @brief SIMD优化的数值设置
     */
    static void simdSetNumbers(UltraCompactCell* cells,
                              const double* input,
                              size_t count);
};

/**
 * @brief 性能测试工具类
 */
class SIMDPerformanceTester {
public:
    /**
     * @brief 运行完整的性能测试套件
     */
    static void runFullPerformanceTest();
    
    /**
     * @brief 测试不同数据大小的性能
     */
    static void testScalability();
    
    /**
     * @brief 测试不同数据类型的性能
     */
    static void testDataTypes();
    
    /**
     * @brief 生成性能报告
     */
    static std::string generatePerformanceReport();

private:
    static std::vector<TXOptimizedSIMDProcessor::PerformanceComparison> results_;
};

} // namespace TinaXlsx
