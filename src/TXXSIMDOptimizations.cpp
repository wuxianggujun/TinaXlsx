//
// @file TXXSIMDOptimizations.cpp
// @brief 基于 xsimd 的高性能SIMD优化实现
//

#include "TinaXlsx/TXXSIMDOptimizations.hpp"
#include <algorithm>
#include <numeric>
#include <cmath>
#include <iostream>
#include <sstream>

namespace TinaXlsx {

// ==================== XSIMDCapabilities 实现 ====================

std::string XSIMDCapabilities::getSIMDArchInfo() {
    std::ostringstream oss;
    oss << "xsimd架构信息:\n";
    oss << "  - 支持的指令集: " << xsimd::default_arch::name() << "\n";
    oss << "  - SIMD寄存器大小: " << getSIMDRegisterSize() << " 字节\n";
    oss << "  - double向量大小: " << xsimd::simd_type<double>::size << "\n";
    oss << "  - float向量大小: " << xsimd::simd_type<float>::size << "\n";
    oss << "  - int64向量大小: " << xsimd::simd_type<int64_t>::size << "\n";
    oss << "  - int32向量大小: " << xsimd::simd_type<int32_t>::size << "\n";
    return oss.str();
}

size_t XSIMDCapabilities::getOptimalBatchSize() {
    // 基于SIMD寄存器大小计算最优批处理大小
    size_t register_size = getSIMDRegisterSize();
    size_t double_batch = register_size / sizeof(double);
    return std::max(double_batch * 4, XSIMDConfig::MIN_BATCH_SIZE);
}

size_t XSIMDCapabilities::getSIMDRegisterSize() {
    // 基于double类型的SIMD向量大小估算寄存器大小
    return xsimd::simd_type<double>::size * sizeof(double);
}

std::string XSIMDCapabilities::getPerformanceInfo() {
    std::ostringstream oss;
    oss << "xsimd性能信息:\n";
    oss << "  - 最优批处理大小: " << getOptimalBatchSize() << "\n";
    oss << "  - 理论加速比 (double): " << xsimd::simd_type<double>::size << "x\n";
    oss << "  - 理论加速比 (float): " << xsimd::simd_type<float>::size << "x\n";
    oss << "  - 理论加速比 (int64): " << xsimd::simd_type<int64_t>::size << "x\n";
    oss << "  - 理论加速比 (int32): " << xsimd::simd_type<int32_t>::size << "x\n";
    return oss.str();
}

// ==================== TXXSIMDProcessor 实现 ====================

// ==================== 数据类型转换优化 ====================

void TXXSIMDProcessor::convertDoublesToCells(const std::vector<double>& input,
                                            std::vector<UltraCompactCell>& output) {
    if (input.empty()) return;
    
    output.resize(input.size());
    
    // 使用xsimd进行批量转换
    const size_t simd_size = xsimd::simd_type<double>::size;
    const size_t simd_end = (input.size() / simd_size) * simd_size;
    
    // SIMD批处理
    for (size_t i = 0; i < simd_end; i += simd_size) {
        // 加载SIMD向量
        auto simd_values = xsimd::load_unaligned(&input[i]);
        
        // 逐个转换（这里可以进一步优化）
        for (size_t j = 0; j < simd_size; ++j) {
            output[i + j] = UltraCompactCell(input[i + j]);
        }
    }
    
    // 处理剩余元素
    for (size_t i = simd_end; i < input.size(); ++i) {
        output[i] = UltraCompactCell(input[i]);
    }
}

void TXXSIMDProcessor::convertInt64sToCells(const std::vector<int64_t>& input,
                                           std::vector<UltraCompactCell>& output) {
    if (input.empty()) return;
    
    output.resize(input.size());
    
    // 使用xsimd进行批量转换
    const size_t simd_size = xsimd::simd_type<int64_t>::size;
    const size_t simd_end = (input.size() / simd_size) * simd_size;
    
    // SIMD批处理
    for (size_t i = 0; i < simd_end; i += simd_size) {
        // 加载SIMD向量
        auto simd_values = xsimd::load_unaligned(&input[i]);
        
        // 逐个转换
        for (size_t j = 0; j < simd_size; ++j) {
            output[i + j] = UltraCompactCell(input[i + j]);
        }
    }
    
    // 处理剩余元素
    for (size_t i = simd_end; i < input.size(); ++i) {
        output[i] = UltraCompactCell(input[i]);
    }
}

void TXXSIMDProcessor::convertFloatsToCells(const std::vector<float>& input,
                                           std::vector<UltraCompactCell>& output) {
    if (input.empty()) return;
    
    output.resize(input.size());
    
    // 使用xsimd进行批量转换
    const size_t simd_size = xsimd::simd_type<float>::size;
    const size_t simd_end = (input.size() / simd_size) * simd_size;
    
    // SIMD批处理
    for (size_t i = 0; i < simd_end; i += simd_size) {
        // 加载SIMD向量
        auto simd_values = xsimd::load_unaligned(&input[i]);
        
        // 转换为double并创建单元格
        for (size_t j = 0; j < simd_size; ++j) {
            output[i + j] = UltraCompactCell(static_cast<double>(input[i + j]));
        }
    }
    
    // 处理剩余元素
    for (size_t i = simd_end; i < input.size(); ++i) {
        output[i] = UltraCompactCell(static_cast<double>(input[i]));
    }
}

void TXXSIMDProcessor::convertInt32sToCells(const std::vector<int32_t>& input,
                                           std::vector<UltraCompactCell>& output) {
    if (input.empty()) return;
    
    output.resize(input.size());
    
    // 使用xsimd进行批量转换
    const size_t simd_size = xsimd::simd_type<int32_t>::size;
    const size_t simd_end = (input.size() / simd_size) * simd_size;
    
    // SIMD批处理
    for (size_t i = 0; i < simd_end; i += simd_size) {
        // 加载SIMD向量
        auto simd_values = xsimd::load_unaligned(&input[i]);
        
        // 转换为int64并创建单元格
        for (size_t j = 0; j < simd_size; ++j) {
            output[i + j] = UltraCompactCell(static_cast<int64_t>(input[i + j]));
        }
    }
    
    // 处理剩余元素
    for (size_t i = simd_end; i < input.size(); ++i) {
        output[i] = UltraCompactCell(static_cast<int64_t>(input[i]));
    }
}

void TXXSIMDProcessor::convertCellsToDoubles(const std::vector<UltraCompactCell>& input,
                                            std::vector<double>& output) {
    if (input.empty()) return;
    
    output.resize(input.size());
    
    // 标量实现（可以进一步优化）
    for (size_t i = 0; i < input.size(); ++i) {
        if (input[i].getType() == UltraCompactCell::CellType::Number) {
            output[i] = input[i].getNumberValue();
        } else if (input[i].getType() == UltraCompactCell::CellType::Integer) {
            output[i] = static_cast<double>(input[i].getIntegerValue());
        } else {
            output[i] = 0.0;
        }
    }
}

void TXXSIMDProcessor::convertCellsToInt64s(const std::vector<UltraCompactCell>& input,
                                           std::vector<int64_t>& output) {
    if (input.empty()) return;
    
    output.resize(input.size());
    
    // 标量实现
    for (size_t i = 0; i < input.size(); ++i) {
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

void TXXSIMDProcessor::clearCells(std::vector<UltraCompactCell>& cells) {
    if (cells.empty()) return;
    
    // 使用标准库的高效实现
    std::fill(cells.begin(), cells.end(), UltraCompactCell());
}

void TXXSIMDProcessor::copyCells(const std::vector<UltraCompactCell>& src,
                                std::vector<UltraCompactCell>& dst) {
    if (src.empty()) return;
    
    dst.resize(src.size());
    
    // 使用标准库的高效实现
    std::copy(src.begin(), src.end(), dst.begin());
}

bool TXXSIMDProcessor::compareCells(const std::vector<UltraCompactCell>& a,
                                   const std::vector<UltraCompactCell>& b) {
    if (a.size() != b.size()) return false;
    
    // 使用标准库的高效实现
    return std::equal(a.begin(), a.end(), b.begin());
}

void TXXSIMDProcessor::fillCells(std::vector<UltraCompactCell>& cells,
                                const UltraCompactCell& value) {
    if (cells.empty()) return;
    
    // 使用标准库的高效实现
    std::fill(cells.begin(), cells.end(), value);
}

// ==================== 坐标操作优化 ====================

void TXXSIMDProcessor::setCoordinates(std::vector<UltraCompactCell>& cells,
                                     const std::vector<uint16_t>& rows,
                                     const std::vector<uint16_t>& cols) {
    if (cells.empty() || rows.size() != cells.size() || cols.size() != cells.size()) {
        return;
    }
    
    // 使用xsimd优化坐标设置
    const size_t simd_size = xsimd::simd_type<uint16_t>::size;
    const size_t simd_end = (cells.size() / simd_size) * simd_size;
    
    // SIMD批处理
    for (size_t i = 0; i < simd_end; i += simd_size) {
        // 加载坐标向量
        auto simd_rows = xsimd::load_unaligned(&rows[i]);
        auto simd_cols = xsimd::load_unaligned(&cols[i]);
        
        // 逐个设置坐标
        for (size_t j = 0; j < simd_size; ++j) {
            cells[i + j].setRow(rows[i + j]);
            cells[i + j].setCol(cols[i + j]);
        }
    }
    
    // 处理剩余元素
    for (size_t i = simd_end; i < cells.size(); ++i) {
        cells[i].setRow(rows[i]);
        cells[i].setCol(cols[i]);
    }
}

void TXXSIMDProcessor::getCoordinates(const std::vector<UltraCompactCell>& cells,
                                     std::vector<uint16_t>& rows,
                                     std::vector<uint16_t>& cols) {
    if (cells.empty()) return;
    
    rows.resize(cells.size());
    cols.resize(cells.size());
    
    // 标量实现
    for (size_t i = 0; i < cells.size(); ++i) {
        rows[i] = cells[i].getRow();
        cols[i] = cells[i].getCol();
    }
}

void TXXSIMDProcessor::transformCoordinates(std::vector<UltraCompactCell>& cells,
                                           int16_t row_offset,
                                           int16_t col_offset) {
    if (cells.empty()) return;
    
    // 标量实现（可以进一步优化）
    for (auto& cell : cells) {
        uint16_t new_row = static_cast<uint16_t>(
            std::max(0, static_cast<int32_t>(cell.getRow()) + row_offset));
        uint16_t new_col = static_cast<uint16_t>(
            std::max(0, static_cast<int32_t>(cell.getCol()) + col_offset));
        
        cell.setRow(new_row);
        cell.setCol(new_col);
    }
}

// ==================== 数值计算优化 ====================

double TXXSIMDProcessor::sumNumbers(const std::vector<UltraCompactCell>& cells) {
    if (cells.empty()) return 0.0;

    // 提取数值到临时向量
    std::vector<double> values;
    values.reserve(cells.size());

    for (const auto& cell : cells) {
        if (cell.getType() == UltraCompactCell::CellType::Number) {
            values.push_back(cell.getNumberValue());
        } else if (cell.getType() == UltraCompactCell::CellType::Integer) {
            values.push_back(static_cast<double>(cell.getIntegerValue()));
        }
    }

    if (values.empty()) return 0.0;

    // 使用xsimd进行高性能求和
    const size_t simd_size = xsimd::simd_type<double>::size;
    const size_t simd_end = (values.size() / simd_size) * simd_size;

    auto sum_vec = xsimd::broadcast<double>(0.0);

    // SIMD批处理求和
    for (size_t i = 0; i < simd_end; i += simd_size) {
        auto values_vec = xsimd::load_unaligned(&values[i]);
        sum_vec = sum_vec + values_vec;
    }

    // 水平求和
    double total = xsimd::reduce_add(sum_vec);

    // 处理剩余元素
    for (size_t i = simd_end; i < values.size(); ++i) {
        total += values[i];
    }

    return total;
}

TXXSIMDProcessor::NumericStats TXXSIMDProcessor::calculateStats(const std::vector<UltraCompactCell>& cells) {
    NumericStats stats;
    if (cells.empty()) return stats;

    // 提取数值到临时向量
    std::vector<double> values;
    values.reserve(cells.size());

    for (const auto& cell : cells) {
        if (cell.getType() == UltraCompactCell::CellType::Number) {
            values.push_back(cell.getNumberValue());
        } else if (cell.getType() == UltraCompactCell::CellType::Integer) {
            values.push_back(static_cast<double>(cell.getIntegerValue()));
        }
    }

    if (values.empty()) return stats;

    stats.count = values.size();

    // 使用xsimd计算统计信息
    const size_t simd_size = xsimd::simd_type<double>::size;
    const size_t simd_end = (values.size() / simd_size) * simd_size;

    auto sum_vec = xsimd::broadcast<double>(0.0);
    auto min_vec = xsimd::broadcast<double>(std::numeric_limits<double>::infinity());
    auto max_vec = xsimd::broadcast<double>(-std::numeric_limits<double>::infinity());

    // SIMD批处理统计
    for (size_t i = 0; i < simd_end; i += simd_size) {
        auto values_vec = xsimd::load_unaligned(&values[i]);
        sum_vec = sum_vec + values_vec;
        min_vec = xsimd::min(min_vec, values_vec);
        max_vec = xsimd::max(max_vec, values_vec);
    }

    // 归约操作
    stats.sum = xsimd::reduce_add(sum_vec);
    stats.min = xsimd::reduce_min(min_vec);
    stats.max = xsimd::reduce_max(max_vec);

    // 处理剩余元素
    for (size_t i = simd_end; i < values.size(); ++i) {
        stats.sum += values[i];
        stats.min = std::min(stats.min, values[i]);
        stats.max = std::max(stats.max, values[i]);
    }

    // 计算均值
    stats.mean = stats.sum / stats.count;

    // 计算方差（第二遍）
    auto variance_vec = xsimd::broadcast<double>(0.0);
    auto mean_vec = xsimd::broadcast<double>(stats.mean);

    for (size_t i = 0; i < simd_end; i += simd_size) {
        auto values_vec = xsimd::load_unaligned(&values[i]);
        auto diff_vec = values_vec - mean_vec;
        variance_vec = variance_vec + (diff_vec * diff_vec);
    }

    stats.variance = xsimd::reduce_add(variance_vec);

    // 处理剩余元素的方差
    for (size_t i = simd_end; i < values.size(); ++i) {
        double diff = values[i] - stats.mean;
        stats.variance += diff * diff;
    }

    stats.variance /= stats.count;
    stats.std_dev = std::sqrt(stats.variance);

    return stats;
}

void TXXSIMDProcessor::addNumbers(const std::vector<UltraCompactCell>& a,
                                 const std::vector<UltraCompactCell>& b,
                                 std::vector<UltraCompactCell>& result) {
    if (a.size() != b.size()) return;

    result.resize(a.size());

    // 标量实现（可以进一步优化）
    for (size_t i = 0; i < a.size(); ++i) {
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

void TXXSIMDProcessor::scalarOperation(const std::vector<UltraCompactCell>& input,
                                      double scalar,
                                      std::vector<UltraCompactCell>& result,
                                      char operation) {
    if (input.empty()) return;

    result.resize(input.size());

    // 提取数值
    std::vector<double> values;
    values.reserve(input.size());

    for (const auto& cell : input) {
        if (cell.getType() == UltraCompactCell::CellType::Number) {
            values.push_back(cell.getNumberValue());
        } else if (cell.getType() == UltraCompactCell::CellType::Integer) {
            values.push_back(static_cast<double>(cell.getIntegerValue()));
        } else {
            values.push_back(0.0);
        }
    }

    // 使用xsimd进行标量运算
    const size_t simd_size = xsimd::simd_type<double>::size;
    const size_t simd_end = (values.size() / simd_size) * simd_size;

    auto scalar_vec = xsimd::broadcast<double>(scalar);

    // SIMD批处理
    for (size_t i = 0; i < simd_end; i += simd_size) {
        auto values_vec = xsimd::load_unaligned(&values[i]);
        auto result_vec = values_vec;

        switch (operation) {
            case '+':
                result_vec = values_vec + scalar_vec;
                break;
            case '-':
                result_vec = values_vec - scalar_vec;
                break;
            case '*':
                result_vec = values_vec * scalar_vec;
                break;
            case '/':
                result_vec = values_vec / scalar_vec;
                break;
            default:
                result_vec = values_vec;
                break;
        }

        // 存储结果
        xsimd::store_unaligned(&values[i], result_vec);
    }

    // 处理剩余元素
    for (size_t i = simd_end; i < values.size(); ++i) {
        switch (operation) {
            case '+':
                values[i] += scalar;
                break;
            case '-':
                values[i] -= scalar;
                break;
            case '*':
                values[i] *= scalar;
                break;
            case '/':
                values[i] /= scalar;
                break;
        }
    }

    // 转换回UltraCompactCell
    for (size_t i = 0; i < input.size(); ++i) {
        result[i] = UltraCompactCell(values[i]);
        result[i].setRow(input[i].getRow());
        result[i].setCol(input[i].getCol());
    }
}

// ==================== 性能测试和基准 ====================

TXXSIMDProcessor::SIMDPerformanceResult TXXSIMDProcessor::benchmarkSIMD(const std::string& operation,
                                                                        size_t test_size) {
    SIMDPerformanceResult result;
    result.arch_info = XSIMDCapabilities::getSIMDArchInfo();
    result.operation_name = operation;
    result.data_size = test_size;

    // 生成测试数据
    std::vector<double> test_doubles(test_size);
    std::vector<UltraCompactCell> test_cells;

    for (size_t i = 0; i < test_size; ++i) {
        test_doubles[i] = static_cast<double>(i) * 3.14159;
    }

    if (operation == "convert" || operation == "all") {
        // 测试转换性能
        auto start = std::chrono::high_resolution_clock::now();
        convertDoublesToCells(test_doubles, test_cells);
        auto end = std::chrono::high_resolution_clock::now();

        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        result.xsimd_time_ms = duration.count() / 1000.0;

        // 标量版本
        std::vector<UltraCompactCell> scalar_cells(test_size);
        start = std::chrono::high_resolution_clock::now();
        for (size_t i = 0; i < test_size; ++i) {
            scalar_cells[i] = UltraCompactCell(test_doubles[i]);
        }
        end = std::chrono::high_resolution_clock::now();

        duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        result.scalar_time_ms = duration.count() / 1000.0;
    }

    if (operation == "sum" || operation == "all") {
        // 测试求和性能
        convertDoublesToCells(test_doubles, test_cells);

        auto start = std::chrono::high_resolution_clock::now();
        double xsimd_sum = sumNumbers(test_cells);
        auto end = std::chrono::high_resolution_clock::now();

        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        result.xsimd_time_ms = duration.count() / 1000.0;

        // 标量版本
        start = std::chrono::high_resolution_clock::now();
        double scalar_sum = 0.0;
        for (const auto& cell : test_cells) {
            if (cell.getType() == UltraCompactCell::CellType::Number) {
                scalar_sum += cell.getNumberValue();
            }
        }
        end = std::chrono::high_resolution_clock::now();

        duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        result.scalar_time_ms = duration.count() / 1000.0;
    }

    // 计算加速比
    if (result.scalar_time_ms > 0) {
        result.speedup_ratio = result.scalar_time_ms / result.xsimd_time_ms;
    }

    // 计算操作数/秒
    if (result.xsimd_time_ms > 0) {
        result.operations_per_second = static_cast<size_t>(test_size / (result.xsimd_time_ms / 1000.0));
    }

    return result;
}

// ==================== 内存对齐和优化 ====================

bool TXXSIMDProcessor::isAligned(const void* ptr, size_t alignment) {
    return (reinterpret_cast<uintptr_t>(ptr) % alignment) == 0;
}

void* TXXSIMDProcessor::alignedAlloc(size_t size, size_t alignment) {
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

void TXXSIMDProcessor::alignedFree(void* ptr) {
#ifdef _MSC_VER
    _aligned_free(ptr);
#else
    free(ptr);
#endif
}

} // namespace TinaXlsx
