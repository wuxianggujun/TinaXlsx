//
// @file TXSIMDOptimizations.hpp
// @brief SIMD优化模块 - 高性能向量化操作
//

#pragma once

#include "TXUltraCompactCell.hpp"
#include "TXTypes.hpp"
#include <vector>
#include <cstdint>

// 检测SIMD支持
#ifdef _MSC_VER
    #include <intrin.h>
    #define TX_SIMD_AVAILABLE
#elif defined(__GNUC__) || defined(__clang__)
    #include <immintrin.h>
    #define TX_SIMD_AVAILABLE
#endif

namespace TinaXlsx {

/**
 * @brief SIMD优化配置
 */
struct SIMDConfig {
    static constexpr size_t AVX2_BATCH_SIZE = 32;      // AVX2批处理大小
    static constexpr size_t SSE_BATCH_SIZE = 16;       // SSE批处理大小
    static constexpr size_t SCALAR_BATCH_SIZE = 8;     // 标量批处理大小
    static constexpr size_t ALIGNMENT = 32;            // 内存对齐要求
};

/**
 * @brief SIMD能力检测
 */
class SIMDCapabilities {
public:
    static bool hasAVX2();
    static bool hasSSE41();
    static bool hasSSE2();
    static size_t getOptimalBatchSize();
    static const char* getSIMDInfo();
    
private:
    static bool avx2_checked_;
    static bool avx2_available_;
    static bool sse41_checked_;
    static bool sse41_available_;
    static bool sse2_checked_;
    static bool sse2_available_;
};

/**
 * @brief SIMD优化的批处理操作
 */
class TXSIMDProcessor {
public:
    TXSIMDProcessor();
    ~TXSIMDProcessor() = default;
    
    // ==================== 数据类型转换优化 ====================
    
    /**
     * @brief 批量转换double到UltraCompactCell
     */
    static void convertDoublesToCells(const double* input,
                                     UltraCompactCell* output,
                                     size_t count);
    
    /**
     * @brief 批量转换int64到UltraCompactCell
     */
    static void convertInt64sToCells(const int64_t* input,
                                    UltraCompactCell* output,
                                    size_t count);
    
    /**
     * @brief 批量转换UltraCompactCell到double
     */
    static void convertCellsToDoubles(const UltraCompactCell* input,
                                     double* output,
                                     size_t count);
    
    /**
     * @brief 批量转换UltraCompactCell到int64
     */
    static void convertCellsToInt64s(const UltraCompactCell* input,
                                    int64_t* output,
                                    size_t count);
    
    // ==================== 内存操作优化 ====================
    
    /**
     * @brief 批量清零UltraCompactCell
     */
    static void clearCells(UltraCompactCell* cells, size_t count);
    
    /**
     * @brief 批量复制UltraCompactCell
     */
    static void copyCells(const UltraCompactCell* src,
                         UltraCompactCell* dst,
                         size_t count);
    
    /**
     * @brief 批量比较UltraCompactCell
     */
    static bool compareCells(const UltraCompactCell* a,
                            const UltraCompactCell* b,
                            size_t count);
    
    // ==================== 坐标操作优化 ====================
    
    /**
     * @brief 批量设置坐标
     */
    static void setCoordinates(UltraCompactCell* cells,
                              const uint16_t* rows,
                              const uint16_t* cols,
                              size_t count);
    
    /**
     * @brief 批量获取坐标
     */
    static void getCoordinates(const UltraCompactCell* cells,
                              uint16_t* rows,
                              uint16_t* cols,
                              size_t count);
    
    // ==================== 类型操作优化 ====================
    
    /**
     * @brief 批量设置单元格类型
     */
    static void setCellTypes(UltraCompactCell* cells,
                            const uint8_t* types,
                            size_t count);
    
    /**
     * @brief 批量获取单元格类型
     */
    static void getCellTypes(const UltraCompactCell* cells,
                            uint8_t* types,
                            size_t count);
    
    /**
     * @brief 批量筛选特定类型的单元格
     */
    static size_t filterCellsByType(const UltraCompactCell* input,
                                   UltraCompactCell* output,
                                   UltraCompactCell::CellType type,
                                   size_t count);
    
    // ==================== 样式操作优化 ====================
    
    /**
     * @brief 批量设置样式索引
     */
    static void setStyleIndices(UltraCompactCell* cells,
                               const uint8_t* styles,
                               size_t count);
    
    /**
     * @brief 批量获取样式索引
     */
    static void getStyleIndices(const UltraCompactCell* cells,
                               uint8_t* styles,
                               size_t count);
    
    // ==================== 数值计算优化 ====================
    
    /**
     * @brief 批量数值运算（加法）
     */
    static void addNumbers(const UltraCompactCell* a,
                          const UltraCompactCell* b,
                          UltraCompactCell* result,
                          size_t count);
    
    /**
     * @brief 批量数值运算（乘法）
     */
    static void multiplyNumbers(const UltraCompactCell* a,
                               const UltraCompactCell* b,
                               UltraCompactCell* result,
                               size_t count);
    
    /**
     * @brief 批量数值统计（求和）
     */
    static double sumNumbers(const UltraCompactCell* cells, size_t count);
    
    /**
     * @brief 批量数值统计（最大值）
     */
    static double maxNumbers(const UltraCompactCell* cells, size_t count);
    
    /**
     * @brief 批量数值统计（最小值）
     */
    static double minNumbers(const UltraCompactCell* cells, size_t count);
    
    // ==================== 性能测试 ====================
    
    /**
     * @brief SIMD性能基准测试
     */
    struct SIMDPerformanceResult {
        double simd_time_ms = 0.0;
        double scalar_time_ms = 0.0;
        double speedup_ratio = 0.0;
        size_t operations_per_second = 0;
        std::string simd_type;
    };
    
    static SIMDPerformanceResult benchmarkSIMD(size_t test_size = 1000000);
    
private:
    // ==================== 内部SIMD实现 ====================
    
#ifdef TX_SIMD_AVAILABLE
    // AVX2优化实现
    static void convertDoublesToCells_AVX2(const double* input,
                                          UltraCompactCell* output,
                                          size_t count);
    
    static void convertInt64sToCells_AVX2(const int64_t* input,
                                         UltraCompactCell* output,
                                         size_t count);
    
    static void clearCells_AVX2(UltraCompactCell* cells, size_t count);
    
    static void copyCells_AVX2(const UltraCompactCell* src,
                              UltraCompactCell* dst,
                              size_t count);
    
    static void setCoordinates_AVX2(UltraCompactCell* cells,
                                   const uint16_t* rows,
                                   const uint16_t* cols,
                                   size_t count);
    
    static double sumNumbers_AVX2(const UltraCompactCell* cells, size_t count);
    
    // SSE优化实现
    static void convertDoublesToCells_SSE(const double* input,
                                         UltraCompactCell* output,
                                         size_t count);
    
    static void clearCells_SSE(UltraCompactCell* cells, size_t count);
    
    static void copyCells_SSE(const UltraCompactCell* src,
                             UltraCompactCell* dst,
                             size_t count);
#endif
    
    // 标量回退实现
    static void convertDoublesToCells_Scalar(const double* input,
                                            UltraCompactCell* output,
                                            size_t count);
    
    static void convertInt64sToCells_Scalar(const int64_t* input,
                                           UltraCompactCell* output,
                                           size_t count);
    
    static void clearCells_Scalar(UltraCompactCell* cells, size_t count);
    
    static void copyCells_Scalar(const UltraCompactCell* src,
                                UltraCompactCell* dst,
                                size_t count);
    
    static double sumNumbers_Scalar(const UltraCompactCell* cells, size_t count);
    
    // 内存对齐辅助函数
    static bool isAligned(const void* ptr, size_t alignment);
    static void* alignedAlloc(size_t size, size_t alignment);
    static void alignedFree(void* ptr);
};

} // namespace TinaXlsx
