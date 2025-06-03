//
// @file TXXSIMDOptimizations.hpp
// @brief 基于 xsimd 的高性能SIMD优化模块
//

#pragma once

#include "TXUltraCompactCell.hpp"
#include "TXTypes.hpp"
#include <xsimd/xsimd.hpp>
#include <vector>
#include <cstdint>
#include <chrono>
#include <string>
#include <functional>
#include <limits>

namespace TinaXlsx {

/**
 * @brief xsimd SIMD优化配置
 */
struct XSIMDConfig {
    static constexpr size_t DEFAULT_BATCH_SIZE = 64;       // 默认批处理大小
    static constexpr size_t MIN_BATCH_SIZE = 8;            // 最小批处理大小
    static constexpr size_t MAX_BATCH_SIZE = 1024;         // 最大批处理大小
    static constexpr size_t ALIGNMENT = 32;                // 内存对齐要求
};

/**
 * @brief xsimd SIMD能力检测和信息
 */
class XSIMDCapabilities {
public:
    /**
     * @brief 获取当前架构支持的SIMD指令集信息
     */
    static std::string getSIMDArchInfo();
    
    /**
     * @brief 获取最优批处理大小
     */
    static size_t getOptimalBatchSize();
    
    /**
     * @brief 获取SIMD寄存器大小（字节）
     */
    static size_t getSIMDRegisterSize();
    
    /**
     * @brief 检查是否支持特定数据类型的SIMD操作
     */
    template<typename T>
    static bool supportsSIMD();
    
    /**
     * @brief 获取性能基准信息
     */
    static std::string getPerformanceInfo();
};

/**
 * @brief 基于 xsimd 的高性能SIMD处理器
 */
class TXXSIMDProcessor {
public:
    TXXSIMDProcessor() = default;
    ~TXXSIMDProcessor() = default;
    
    // ==================== 数据类型转换优化 ====================
    
    /**
     * @brief 高性能批量转换double到UltraCompactCell
     */
    static void convertDoublesToCells(const std::vector<double>& input,
                                     std::vector<UltraCompactCell>& output);
    
    /**
     * @brief 高性能批量转换int64到UltraCompactCell
     */
    static void convertInt64sToCells(const std::vector<int64_t>& input,
                                    std::vector<UltraCompactCell>& output);
    
    /**
     * @brief 高性能批量转换float到UltraCompactCell
     */
    static void convertFloatsToCells(const std::vector<float>& input,
                                    std::vector<UltraCompactCell>& output);
    
    /**
     * @brief 高性能批量转换int32到UltraCompactCell
     */
    static void convertInt32sToCells(const std::vector<int32_t>& input,
                                    std::vector<UltraCompactCell>& output);
    
    /**
     * @brief 高性能批量转换UltraCompactCell到double
     */
    static void convertCellsToDoubles(const std::vector<UltraCompactCell>& input,
                                     std::vector<double>& output);
    
    /**
     * @brief 高性能批量转换UltraCompactCell到int64
     */
    static void convertCellsToInt64s(const std::vector<UltraCompactCell>& input,
                                    std::vector<int64_t>& output);
    
    // ==================== 内存操作优化 ====================
    
    /**
     * @brief 高性能批量清零UltraCompactCell
     */
    static void clearCells(std::vector<UltraCompactCell>& cells);
    
    /**
     * @brief 高性能批量复制UltraCompactCell
     */
    static void copyCells(const std::vector<UltraCompactCell>& src,
                         std::vector<UltraCompactCell>& dst);
    
    /**
     * @brief 高性能批量比较UltraCompactCell
     */
    static bool compareCells(const std::vector<UltraCompactCell>& a,
                            const std::vector<UltraCompactCell>& b);
    
    /**
     * @brief 高性能批量填充UltraCompactCell
     */
    static void fillCells(std::vector<UltraCompactCell>& cells,
                         const UltraCompactCell& value);
    
    // ==================== 坐标操作优化 ====================
    
    /**
     * @brief 高性能批量设置坐标
     */
    static void setCoordinates(std::vector<UltraCompactCell>& cells,
                              const std::vector<uint16_t>& rows,
                              const std::vector<uint16_t>& cols);
    
    /**
     * @brief 高性能批量获取坐标
     */
    static void getCoordinates(const std::vector<UltraCompactCell>& cells,
                              std::vector<uint16_t>& rows,
                              std::vector<uint16_t>& cols);
    
    /**
     * @brief 高性能批量坐标变换
     */
    static void transformCoordinates(std::vector<UltraCompactCell>& cells,
                                    int16_t row_offset,
                                    int16_t col_offset);
    
    // ==================== 数值计算优化 ====================
    
    /**
     * @brief 高性能数值求和
     */
    static double sumNumbers(const std::vector<UltraCompactCell>& cells);
    
    /**
     * @brief 高性能数值统计
     */
    struct NumericStats {
        double sum = 0.0;
        double mean = 0.0;
        double min = std::numeric_limits<double>::infinity();
        double max = -std::numeric_limits<double>::infinity();
        double variance = 0.0;
        double std_dev = 0.0;
        size_t count = 0;
    };
    
    static NumericStats calculateStats(const std::vector<UltraCompactCell>& cells);
    
    /**
     * @brief 高性能数值运算（加法）
     */
    static void addNumbers(const std::vector<UltraCompactCell>& a,
                          const std::vector<UltraCompactCell>& b,
                          std::vector<UltraCompactCell>& result);
    
    /**
     * @brief 高性能数值运算（减法）
     */
    static void subtractNumbers(const std::vector<UltraCompactCell>& a,
                               const std::vector<UltraCompactCell>& b,
                               std::vector<UltraCompactCell>& result);
    
    /**
     * @brief 高性能数值运算（乘法）
     */
    static void multiplyNumbers(const std::vector<UltraCompactCell>& a,
                               const std::vector<UltraCompactCell>& b,
                               std::vector<UltraCompactCell>& result);
    
    /**
     * @brief 高性能数值运算（除法）
     */
    static void divideNumbers(const std::vector<UltraCompactCell>& a,
                             const std::vector<UltraCompactCell>& b,
                             std::vector<UltraCompactCell>& result);
    
    /**
     * @brief 高性能标量运算
     */
    static void scalarOperation(const std::vector<UltraCompactCell>& input,
                               double scalar,
                               std::vector<UltraCompactCell>& result,
                               char operation); // '+', '-', '*', '/'
    
    // ==================== 筛选和排序优化 ====================
    
    /**
     * @brief 高性能筛选单元格
     */
    static std::vector<UltraCompactCell> filterCells(
        const std::vector<UltraCompactCell>& cells,
        const std::function<bool(const UltraCompactCell&)>& predicate);
    
    /**
     * @brief 高性能查找操作
     */
    static std::vector<size_t> findCells(
        const std::vector<UltraCompactCell>& cells,
        const UltraCompactCell& target);
    
    /**
     * @brief 高性能计数操作
     */
    static size_t countCells(
        const std::vector<UltraCompactCell>& cells,
        const std::function<bool(const UltraCompactCell&)>& predicate);
    
    // ==================== 性能测试和基准 ====================
    
    /**
     * @brief SIMD性能测试结果
     */
    struct SIMDPerformanceResult {
        double xsimd_time_ms = 0.0;
        double scalar_time_ms = 0.0;
        double speedup_ratio = 0.0;
        size_t operations_per_second = 0;
        std::string arch_info;
        std::string operation_name;
        size_t data_size = 0;
    };
    
    /**
     * @brief 运行SIMD性能基准测试
     */
    static SIMDPerformanceResult benchmarkSIMD(const std::string& operation = "all",
                                              size_t test_size = 1000000);
    
    /**
     * @brief 运行特定操作的性能测试
     */
    static SIMDPerformanceResult benchmarkOperation(
        const std::string& operation_name,
        const std::function<void()>& xsimd_operation,
        const std::function<void()>& scalar_operation,
        size_t test_size);
    
    // ==================== 内存对齐和优化 ====================
    
    /**
     * @brief 检查内存对齐
     */
    static bool isAligned(const void* ptr, size_t alignment = XSIMDConfig::ALIGNMENT);
    
    /**
     * @brief 分配对齐内存
     */
    static void* alignedAlloc(size_t size, size_t alignment = XSIMDConfig::ALIGNMENT);
    
    /**
     * @brief 释放对齐内存
     */
    static void alignedFree(void* ptr);
    
    /**
     * @brief 创建对齐的vector
     */
    template<typename T>
    static std::vector<T> createAlignedVector(size_t size);

private:
    // ==================== 内部SIMD实现 ====================
    
    /**
     * @brief 内部批处理转换实现
     */
    template<typename InputT, typename OutputT>
    static void batchConvert(const std::vector<InputT>& input,
                            std::vector<OutputT>& output,
                            const std::function<OutputT(InputT)>& converter);
    
    /**
     * @brief 内部SIMD数值运算实现
     */
    template<typename T>
    static void simdNumericOperation(const std::vector<T>& a,
                                    const std::vector<T>& b,
                                    std::vector<T>& result,
                                    const std::function<T(T, T)>& operation);
    
    /**
     * @brief 内部SIMD归约操作实现
     */
    template<typename T>
    static T simdReduce(const std::vector<T>& data,
                       T init,
                       const std::function<T(T, T)>& operation);
    
    /**
     * @brief 获取最优批处理大小（内部）
     */
    static size_t getOptimalBatchSizeInternal(size_t data_size, size_t element_size);
};

// ==================== 模板实现 ====================

template<typename T>
bool XSIMDCapabilities::supportsSIMD() {
    return xsimd::simd_type<T>::size > 1;
}

template<typename T>
std::vector<T> TXXSIMDProcessor::createAlignedVector(size_t size) {
    std::vector<T> vec;
    vec.reserve(size);
    vec.resize(size);
    return vec;
}

} // namespace TinaXlsx
