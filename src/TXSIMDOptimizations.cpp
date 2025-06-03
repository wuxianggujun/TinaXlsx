//
// @file TXSIMDOptimizations.cpp
// @brief SIMD优化实现
//

#include "TinaXlsx/TXSIMDOptimizations.hpp"
#include <chrono>
#include <algorithm>
#include <cstring>
#include <iostream>

#ifdef TX_SIMD_AVAILABLE
    #ifdef _MSC_VER
        #include <intrin.h>
    #else
        #include <immintrin.h>
        #include <cpuid.h>
    #endif
#endif

namespace TinaXlsx {

// ==================== SIMD能力检测实现 ====================

bool SIMDCapabilities::avx2_checked_ = false;
bool SIMDCapabilities::avx2_available_ = false;
bool SIMDCapabilities::sse41_checked_ = false;
bool SIMDCapabilities::sse41_available_ = false;
bool SIMDCapabilities::sse2_checked_ = false;
bool SIMDCapabilities::sse2_available_ = false;

bool SIMDCapabilities::hasAVX2() {
    if (!avx2_checked_) {
#ifdef TX_SIMD_AVAILABLE
        #ifdef _MSC_VER
            int cpuInfo[4];
            __cpuid(cpuInfo, 7);
            avx2_available_ = (cpuInfo[1] & (1 << 5)) != 0;
        #else
            unsigned int eax, ebx, ecx, edx;
            if (__get_cpuid_max(0, nullptr) >= 7) {
                __cpuid_count(7, 0, eax, ebx, ecx, edx);
                avx2_available_ = (ebx & (1 << 5)) != 0;
            }
        #endif
#endif
        avx2_checked_ = true;
    }
    return avx2_available_;
}

bool SIMDCapabilities::hasSSE41() {
    if (!sse41_checked_) {
#ifdef TX_SIMD_AVAILABLE
        #ifdef _MSC_VER
            int cpuInfo[4];
            __cpuid(cpuInfo, 1);
            sse41_available_ = (cpuInfo[2] & (1 << 19)) != 0;
        #else
            unsigned int eax, ebx, ecx, edx;
            __cpuid(1, eax, ebx, ecx, edx);
            sse41_available_ = (ecx & (1 << 19)) != 0;
        #endif
#endif
        sse41_checked_ = true;
    }
    return sse41_available_;
}

bool SIMDCapabilities::hasSSE2() {
    if (!sse2_checked_) {
#ifdef TX_SIMD_AVAILABLE
        #ifdef _MSC_VER
            int cpuInfo[4];
            __cpuid(cpuInfo, 1);
            sse2_available_ = (cpuInfo[3] & (1 << 26)) != 0;
        #else
            unsigned int eax, ebx, ecx, edx;
            __cpuid(1, eax, ebx, ecx, edx);
            sse2_available_ = (edx & (1 << 26)) != 0;
        #endif
#endif
        sse2_checked_ = true;
    }
    return sse2_available_;
}

size_t SIMDCapabilities::getOptimalBatchSize() {
    if (hasAVX2()) {
        return SIMDConfig::AVX2_BATCH_SIZE;
    } else if (hasSSE41() || hasSSE2()) {
        return SIMDConfig::SSE_BATCH_SIZE;
    } else {
        return SIMDConfig::SCALAR_BATCH_SIZE;
    }
}

const char* SIMDCapabilities::getSIMDInfo() {
    if (hasAVX2()) {
        return "AVX2";
    } else if (hasSSE41()) {
        return "SSE4.1";
    } else if (hasSSE2()) {
        return "SSE2";
    } else {
        return "Scalar";
    }
}

// ==================== TXSIMDProcessor实现 ====================

TXSIMDProcessor::TXSIMDProcessor() {
    // 初始化时检测SIMD能力
    SIMDCapabilities::hasAVX2();
    SIMDCapabilities::hasSSE41();
    SIMDCapabilities::hasSSE2();
}

// ==================== 数据类型转换优化 ====================

void TXSIMDProcessor::convertDoublesToCells(const double* input,
                                           UltraCompactCell* output,
                                           size_t count) {
    if (count == 0) return;
    
#ifdef TX_SIMD_AVAILABLE
    if (SIMDCapabilities::hasAVX2() && count >= SIMDConfig::AVX2_BATCH_SIZE) {
        convertDoublesToCells_AVX2(input, output, count);
    } else if (SIMDCapabilities::hasSSE2() && count >= SIMDConfig::SSE_BATCH_SIZE) {
        convertDoublesToCells_SSE(input, output, count);
    } else {
        convertDoublesToCells_Scalar(input, output, count);
    }
#else
    convertDoublesToCells_Scalar(input, output, count);
#endif
}

void TXSIMDProcessor::convertInt64sToCells(const int64_t* input,
                                          UltraCompactCell* output,
                                          size_t count) {
    if (count == 0) return;
    
#ifdef TX_SIMD_AVAILABLE
    if (SIMDCapabilities::hasAVX2() && count >= SIMDConfig::AVX2_BATCH_SIZE) {
        convertInt64sToCells_AVX2(input, output, count);
    } else {
        convertInt64sToCells_Scalar(input, output, count);
    }
#else
    convertInt64sToCells_Scalar(input, output, count);
#endif
}

void TXSIMDProcessor::convertCellsToDoubles(const UltraCompactCell* input,
                                           double* output,
                                           size_t count) {
    // 标量实现（SIMD优化可以后续添加）
    for (size_t i = 0; i < count; ++i) {
        if (input[i].getType() == UltraCompactCell::CellType::Number) {
            output[i] = input[i].getNumberValue();
        } else if (input[i].getType() == UltraCompactCell::CellType::Integer) {
            output[i] = static_cast<double>(input[i].getIntegerValue());
        } else {
            output[i] = 0.0;
        }
    }
}

void TXSIMDProcessor::convertCellsToInt64s(const UltraCompactCell* input,
                                          int64_t* output,
                                          size_t count) {
    // 标量实现
    for (size_t i = 0; i < count; ++i) {
        if (input[i].getType() == UltraCompactCell::CellType::Integer) {
            output[i] = input[i].getIntegerValue();
        } else if (input[i].getType() == UltraCompactCell::CellType::Number) {
            output[i] = static_cast<int64_t>(input[i].getNumberValue());
        } else {
            output[i] = 0;
        }
    }
}

// ==================== 内存操作优化 ====================

void TXSIMDProcessor::clearCells(UltraCompactCell* cells, size_t count) {
    if (count == 0) return;
    
#ifdef TX_SIMD_AVAILABLE
    if (SIMDCapabilities::hasAVX2() && count >= SIMDConfig::AVX2_BATCH_SIZE) {
        clearCells_AVX2(cells, count);
    } else if (SIMDCapabilities::hasSSE2() && count >= SIMDConfig::SSE_BATCH_SIZE) {
        clearCells_SSE(cells, count);
    } else {
        clearCells_Scalar(cells, count);
    }
#else
    clearCells_Scalar(cells, count);
#endif
}

void TXSIMDProcessor::copyCells(const UltraCompactCell* src,
                               UltraCompactCell* dst,
                               size_t count) {
    if (count == 0) return;
    
#ifdef TX_SIMD_AVAILABLE
    if (SIMDCapabilities::hasAVX2() && count >= SIMDConfig::AVX2_BATCH_SIZE) {
        copyCells_AVX2(src, dst, count);
    } else if (SIMDCapabilities::hasSSE2() && count >= SIMDConfig::SSE_BATCH_SIZE) {
        copyCells_SSE(src, dst, count);
    } else {
        copyCells_Scalar(src, dst, count);
    }
#else
    copyCells_Scalar(src, dst, count);
#endif
}

bool TXSIMDProcessor::compareCells(const UltraCompactCell* a,
                                  const UltraCompactCell* b,
                                  size_t count) {
    // 简单的内存比较
    return std::memcmp(a, b, count * sizeof(UltraCompactCell)) == 0;
}

// ==================== 坐标操作优化 ====================

void TXSIMDProcessor::setCoordinates(UltraCompactCell* cells,
                                    const uint16_t* rows,
                                    const uint16_t* cols,
                                    size_t count) {
#ifdef TX_SIMD_AVAILABLE
    if (SIMDCapabilities::hasAVX2() && count >= SIMDConfig::AVX2_BATCH_SIZE) {
        setCoordinates_AVX2(cells, rows, cols, count);
        return;
    }
#endif
    
    // 标量实现
    for (size_t i = 0; i < count; ++i) {
        cells[i].setRow(rows[i]);
        cells[i].setCol(cols[i]);
    }
}

void TXSIMDProcessor::getCoordinates(const UltraCompactCell* cells,
                                    uint16_t* rows,
                                    uint16_t* cols,
                                    size_t count) {
    // 标量实现
    for (size_t i = 0; i < count; ++i) {
        rows[i] = cells[i].getRow();
        cols[i] = cells[i].getCol();
    }
}

// ==================== 类型操作优化 ====================

void TXSIMDProcessor::setCellTypes(UltraCompactCell* cells,
                                  const uint8_t* types,
                                  size_t count) {
    // 标量实现
    for (size_t i = 0; i < count; ++i) {
        cells[i].setType(static_cast<UltraCompactCell::CellType>(types[i]));
    }
}

void TXSIMDProcessor::getCellTypes(const UltraCompactCell* cells,
                                  uint8_t* types,
                                  size_t count) {
    // 标量实现
    for (size_t i = 0; i < count; ++i) {
        types[i] = static_cast<uint8_t>(cells[i].getType());
    }
}

size_t TXSIMDProcessor::filterCellsByType(const UltraCompactCell* input,
                                         UltraCompactCell* output,
                                         UltraCompactCell::CellType type,
                                         size_t count) {
    size_t output_count = 0;
    for (size_t i = 0; i < count; ++i) {
        if (input[i].getType() == type) {
            output[output_count++] = input[i];
        }
    }
    return output_count;
}

// ==================== 样式操作优化 ====================

void TXSIMDProcessor::setStyleIndices(UltraCompactCell* cells,
                                     const uint8_t* styles,
                                     size_t count) {
    // 标量实现
    for (size_t i = 0; i < count; ++i) {
        cells[i].setStyleIndex(styles[i]);
    }
}

void TXSIMDProcessor::getStyleIndices(const UltraCompactCell* cells,
                                     uint8_t* styles,
                                     size_t count) {
    // 标量实现
    for (size_t i = 0; i < count; ++i) {
        styles[i] = cells[i].getStyleIndex();
    }
}

// ==================== 数值计算优化 ====================

void TXSIMDProcessor::addNumbers(const UltraCompactCell* a,
                                const UltraCompactCell* b,
                                UltraCompactCell* result,
                                size_t count) {
    // 标量实现
    for (size_t i = 0; i < count; ++i) {
        if (a[i].getType() == UltraCompactCell::CellType::Number &&
            b[i].getType() == UltraCompactCell::CellType::Number) {
            double sum = a[i].getNumberValue() + b[i].getNumberValue();
            result[i] = UltraCompactCell(sum);
            result[i].setRow(a[i].getRow());
            result[i].setCol(a[i].getCol());
        } else {
            result[i] = a[i]; // 复制第一个操作数
        }
    }
}

void TXSIMDProcessor::multiplyNumbers(const UltraCompactCell* a,
                                     const UltraCompactCell* b,
                                     UltraCompactCell* result,
                                     size_t count) {
    // 标量实现
    for (size_t i = 0; i < count; ++i) {
        if (a[i].getType() == UltraCompactCell::CellType::Number &&
            b[i].getType() == UltraCompactCell::CellType::Number) {
            double product = a[i].getNumberValue() * b[i].getNumberValue();
            result[i] = UltraCompactCell(product);
            result[i].setRow(a[i].getRow());
            result[i].setCol(a[i].getCol());
        } else {
            result[i] = a[i]; // 复制第一个操作数
        }
    }
}

double TXSIMDProcessor::sumNumbers(const UltraCompactCell* cells, size_t count) {
#ifdef TX_SIMD_AVAILABLE
    if (SIMDCapabilities::hasAVX2() && count >= SIMDConfig::AVX2_BATCH_SIZE) {
        return sumNumbers_AVX2(cells, count);
    }
#endif
    return sumNumbers_Scalar(cells, count);
}

double TXSIMDProcessor::maxNumbers(const UltraCompactCell* cells, size_t count) {
    double max_val = -std::numeric_limits<double>::infinity();
    for (size_t i = 0; i < count; ++i) {
        if (cells[i].getType() == UltraCompactCell::CellType::Number) {
            max_val = std::max(max_val, cells[i].getNumberValue());
        }
    }
    return max_val;
}

double TXSIMDProcessor::minNumbers(const UltraCompactCell* cells, size_t count) {
    double min_val = std::numeric_limits<double>::infinity();
    for (size_t i = 0; i < count; ++i) {
        if (cells[i].getType() == UltraCompactCell::CellType::Number) {
            min_val = std::min(min_val, cells[i].getNumberValue());
        }
    }
    return min_val;
}

// ==================== SIMD具体实现 ====================

#ifdef TX_SIMD_AVAILABLE

// AVX2优化实现
void TXSIMDProcessor::convertDoublesToCells_AVX2(const double* input,
                                                UltraCompactCell* output,
                                                size_t count) {
    const size_t simd_count = (count / 4) * 4; // AVX2处理4个double

    for (size_t i = 0; i < simd_count; i += 4) {
        // 加载4个double值
        __m256d values = _mm256_load_pd(&input[i]);

        // 转换为UltraCompactCell（标量回退）
        for (size_t j = 0; j < 4; ++j) {
            double val = input[i + j];
            output[i + j] = UltraCompactCell(val);
        }
    }

    // 处理剩余元素
    for (size_t i = simd_count; i < count; ++i) {
        output[i] = UltraCompactCell(input[i]);
    }
}

void TXSIMDProcessor::convertInt64sToCells_AVX2(const int64_t* input,
                                               UltraCompactCell* output,
                                               size_t count) {
    const size_t simd_count = (count / 4) * 4; // AVX2处理4个int64

    for (size_t i = 0; i < simd_count; i += 4) {
        // 加载4个int64值
        __m256i values = _mm256_load_si256(reinterpret_cast<const __m256i*>(&input[i]));

        // 转换为UltraCompactCell（标量回退）
        for (size_t j = 0; j < 4; ++j) {
            int64_t val = input[i + j];
            output[i + j] = UltraCompactCell(val);
        }
    }

    // 处理剩余元素
    for (size_t i = simd_count; i < count; ++i) {
        output[i] = UltraCompactCell(input[i]);
    }
}

void TXSIMDProcessor::clearCells_AVX2(UltraCompactCell* cells, size_t count) {
    const size_t simd_count = (count / 2) * 2; // AVX2处理2个UltraCompactCell (32字节)
    __m256i zero = _mm256_setzero_si256();

    for (size_t i = 0; i < simd_count; i += 2) {
        _mm256_store_si256(reinterpret_cast<__m256i*>(&cells[i]), zero);
    }

    // 处理剩余元素
    for (size_t i = simd_count; i < count; ++i) {
        cells[i].clear();
    }
}

void TXSIMDProcessor::copyCells_AVX2(const UltraCompactCell* src,
                                    UltraCompactCell* dst,
                                    size_t count) {
    const size_t simd_count = (count / 2) * 2; // AVX2处理2个UltraCompactCell

    for (size_t i = 0; i < simd_count; i += 2) {
        __m256i data = _mm256_load_si256(reinterpret_cast<const __m256i*>(&src[i]));
        _mm256_store_si256(reinterpret_cast<__m256i*>(&dst[i]), data);
    }

    // 处理剩余元素
    for (size_t i = simd_count; i < count; ++i) {
        dst[i] = src[i];
    }
}

void TXSIMDProcessor::setCoordinates_AVX2(UltraCompactCell* cells,
                                         const uint16_t* rows,
                                         const uint16_t* cols,
                                         size_t count) {
    // AVX2优化的坐标设置（简化实现）
    const size_t simd_count = (count / 8) * 8; // 处理8个坐标对

    for (size_t i = 0; i < simd_count; i += 8) {
        // 加载8个row和8个col
        __m128i rows_vec = _mm_load_si128(reinterpret_cast<const __m128i*>(&rows[i]));
        __m128i cols_vec = _mm_load_si128(reinterpret_cast<const __m128i*>(&cols[i]));

        // 标量回退设置坐标
        for (size_t j = 0; j < 8; ++j) {
            cells[i + j].setRow(rows[i + j]);
            cells[i + j].setCol(cols[i + j]);
        }
    }

    // 处理剩余元素
    for (size_t i = simd_count; i < count; ++i) {
        cells[i].setRow(rows[i]);
        cells[i].setCol(cols[i]);
    }
}

double TXSIMDProcessor::sumNumbers_AVX2(const UltraCompactCell* cells, size_t count) {
    __m256d sum_vec = _mm256_setzero_pd();
    const size_t simd_count = (count / 4) * 4;

    for (size_t i = 0; i < simd_count; i += 4) {
        // 检查类型并提取数值（标量方式）
        double values[4] = {0.0, 0.0, 0.0, 0.0};
        for (size_t j = 0; j < 4; ++j) {
            if (cells[i + j].getType() == UltraCompactCell::CellType::Number) {
                values[j] = cells[i + j].getNumberValue();
            }
        }

        // 加载到AVX2寄存器并累加
        __m256d vals = _mm256_load_pd(values);
        sum_vec = _mm256_add_pd(sum_vec, vals);
    }

    // 水平求和
    double result[4];
    _mm256_store_pd(result, sum_vec);
    double total = result[0] + result[1] + result[2] + result[3];

    // 处理剩余元素
    for (size_t i = simd_count; i < count; ++i) {
        if (cells[i].getType() == UltraCompactCell::CellType::Number) {
            total += cells[i].getNumberValue();
        }
    }

    return total;
}

// SSE优化实现
void TXSIMDProcessor::convertDoublesToCells_SSE(const double* input,
                                               UltraCompactCell* output,
                                               size_t count) {
    const size_t simd_count = (count / 2) * 2; // SSE处理2个double

    for (size_t i = 0; i < simd_count; i += 2) {
        // 标量回退（SSE实现可以更复杂）
        output[i] = UltraCompactCell(input[i]);
        output[i + 1] = UltraCompactCell(input[i + 1]);
    }

    // 处理剩余元素
    for (size_t i = simd_count; i < count; ++i) {
        output[i] = UltraCompactCell(input[i]);
    }
}

void TXSIMDProcessor::clearCells_SSE(UltraCompactCell* cells, size_t count) {
    const size_t simd_count = (count / 1) * 1; // SSE处理1个UltraCompactCell (16字节)
    __m128i zero = _mm_setzero_si128();

    for (size_t i = 0; i < simd_count; ++i) {
        _mm_store_si128(reinterpret_cast<__m128i*>(&cells[i]), zero);
    }
}

void TXSIMDProcessor::copyCells_SSE(const UltraCompactCell* src,
                                   UltraCompactCell* dst,
                                   size_t count) {
    for (size_t i = 0; i < count; ++i) {
        __m128i data = _mm_load_si128(reinterpret_cast<const __m128i*>(&src[i]));
        _mm_store_si128(reinterpret_cast<__m128i*>(&dst[i]), data);
    }
}

#endif // TX_SIMD_AVAILABLE

// ==================== 标量回退实现 ====================

void TXSIMDProcessor::convertDoublesToCells_Scalar(const double* input,
                                                  UltraCompactCell* output,
                                                  size_t count) {
    for (size_t i = 0; i < count; ++i) {
        output[i] = UltraCompactCell(input[i]);
    }
}

void TXSIMDProcessor::convertInt64sToCells_Scalar(const int64_t* input,
                                                UltraCompactCell* output,
                                                size_t count) {
    for (size_t i = 0; i < count; ++i) {
        output[i] = UltraCompactCell(input[i]);
    }
}

void TXSIMDProcessor::clearCells_Scalar(UltraCompactCell* cells, size_t count) {
    for (size_t i = 0; i < count; ++i) {
        cells[i].clear();
    }
}

void TXSIMDProcessor::copyCells_Scalar(const UltraCompactCell* src,
                                      UltraCompactCell* dst,
                                      size_t count) {
    std::memcpy(dst, src, count * sizeof(UltraCompactCell));
}

double TXSIMDProcessor::sumNumbers_Scalar(const UltraCompactCell* cells, size_t count) {
    double sum = 0.0;
    for (size_t i = 0; i < count; ++i) {
        if (cells[i].getType() == UltraCompactCell::CellType::Number) {
            sum += cells[i].getNumberValue();
        }
    }
    return sum;
}

// ==================== 性能测试实现 ====================

TXSIMDProcessor::SIMDPerformanceResult TXSIMDProcessor::benchmarkSIMD(size_t test_size) {
    SIMDPerformanceResult result;
    result.simd_type = SIMDCapabilities::getSIMDInfo();

    // 准备测试数据
    std::vector<double> input_doubles(test_size);
    std::vector<UltraCompactCell> output_simd(test_size);
    std::vector<UltraCompactCell> output_scalar(test_size);

    // 初始化测试数据
    for (size_t i = 0; i < test_size; ++i) {
        input_doubles[i] = static_cast<double>(i) * 3.14159;
    }

    // SIMD性能测试
    auto start = std::chrono::high_resolution_clock::now();
    convertDoublesToCells(input_doubles.data(), output_simd.data(), test_size);
    auto end = std::chrono::high_resolution_clock::now();

    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    result.simd_time_ms = duration.count() / 1000.0;

    // 标量性能测试
    start = std::chrono::high_resolution_clock::now();
    convertDoublesToCells_Scalar(input_doubles.data(), output_scalar.data(), test_size);
    end = std::chrono::high_resolution_clock::now();

    duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    result.scalar_time_ms = duration.count() / 1000.0;

    // 计算加速比
    if (result.scalar_time_ms > 0) {
        result.speedup_ratio = result.scalar_time_ms / result.simd_time_ms;
    }

    // 计算操作数/秒
    if (result.simd_time_ms > 0) {
        result.operations_per_second = static_cast<size_t>(test_size / (result.simd_time_ms / 1000.0));
    }

    // 验证结果正确性
    bool results_match = compareCells(output_simd.data(), output_scalar.data(), test_size);
    if (!results_match) {
        std::cerr << "Warning: SIMD and scalar results do not match!" << std::endl;
    }

    return result;
}

// ==================== 内存对齐辅助函数 ====================

bool TXSIMDProcessor::isAligned(const void* ptr, size_t alignment) {
    return (reinterpret_cast<uintptr_t>(ptr) % alignment) == 0;
}

void* TXSIMDProcessor::alignedAlloc(size_t size, size_t alignment) {
#ifdef _MSC_VER
    return _aligned_malloc(size, alignment);
#else
    void* ptr = nullptr;
    if (posix_memalign(&ptr, alignment, size) != 0) {
        return nullptr;
    }
    return ptr;
#endif
}

void TXSIMDProcessor::alignedFree(void* ptr) {
#ifdef _MSC_VER
    _aligned_free(ptr);
#else
    free(ptr);
#endif
}

} // namespace TinaXlsx
