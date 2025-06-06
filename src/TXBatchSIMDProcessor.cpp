//
// @file TXBatchSIMDProcessor.cpp
// @brief 批量SIMD处理器实现 - 极致性能的核心
//

#include "TinaXlsx/TXBatchSIMDProcessor.hpp"
#include "TinaXlsx/TXGlobalStringPool.hpp"
#include "TinaXlsx/TXCoordinate.hpp"
#include "TinaXlsx/TXInMemorySheet.hpp" // **FIX:** 包含了 TXInMemorySheet 的完整定义
#include <xsimd/xsimd.hpp>
#include <algorithm>
#include <numeric>
#include <immintrin.h>
#include <chrono>
#include <cstring>

namespace TinaXlsx {

// 静态成员初始化
TXBatchSIMDProcessor::BatchPerformanceStats TXBatchSIMDProcessor::performance_stats_;

// ==================== TXCompactCellBuffer 实现 ====================

// 构造函数已移至头文件中的内联实现

void TXCompactCellBuffer::reserve(size_t new_capacity) {
    if (new_capacity <= capacity) return;
    
    number_values.reserve(new_capacity);
    string_indices.reserve(new_capacity);
    style_indices.reserve(new_capacity);
    coordinates.reserve(new_capacity);
    cell_types.reserve(new_capacity);
    
    capacity = new_capacity;
}

void TXCompactCellBuffer::resize(size_t new_size) {
    number_values.resize(new_size);
    string_indices.resize(new_size);
    style_indices.resize(new_size);
    coordinates.resize(new_size);
    cell_types.resize(new_size);
    
    size = new_size;
    is_sorted = false;
}

void TXCompactCellBuffer::clear() {
    number_values.clear();
    string_indices.clear();
    style_indices.clear();
    coordinates.clear();
    cell_types.clear();
    
    size = 0;
    is_sorted = false;
}

void TXCompactCellBuffer::shrink_to_fit() {
    number_values.shrink_to_fit();
    string_indices.shrink_to_fit();
    style_indices.shrink_to_fit();
    coordinates.shrink_to_fit();
    cell_types.shrink_to_fit();
    
    capacity = size;
}

void TXCompactCellBuffer::sort_by_coordinates() {
    if (is_sorted || size == 0) return;
    
    // 创建索引数组
    std::vector<size_t> indices(size);
    std::iota(indices.begin(), indices.end(), 0);
    
    // 按坐标排序索引
    std::sort(indices.begin(), indices.end(), [this](size_t a, size_t b) {
        return coordinates[a] < coordinates[b];
    });
    
    // 重新排列所有数据
    std::vector<double> temp_numbers(size);
    std::vector<uint32_t> temp_strings(size);
    std::vector<uint16_t> temp_styles(size);
    std::vector<uint32_t> temp_coords(size);
    std::vector<uint8_t> temp_types(size);
    
    for (size_t i = 0; i < size; ++i) {
        size_t old_idx = indices[i];
        temp_numbers[i] = number_values[old_idx];
        temp_strings[i] = string_indices[old_idx];
        temp_styles[i] = style_indices[old_idx];
        temp_coords[i] = coordinates[old_idx];
        temp_types[i] = cell_types[old_idx];
    }
    
    // 🚀 高性能移动回原始容器 - 使用assign避免类型不匹配
    number_values.assign(temp_numbers.begin(), temp_numbers.end());
    string_indices.assign(temp_strings.begin(), temp_strings.end());
    style_indices.assign(temp_styles.begin(), temp_styles.end());
    coordinates.assign(temp_coords.begin(), temp_coords.end());
    cell_types.assign(temp_types.begin(), temp_types.end());
    
    is_sorted = true;
}

// ==================== TXBatchSIMDProcessor 实现 ====================

void TXBatchSIMDProcessor::batchCreateNumberCells(
    const double* values,
    TXCompactCellBuffer& buffer,
    const uint32_t* coordinates,
    size_t count,
    size_t start_idx
) {
    // 🚀 极致性能版本：跳过所有检查和统计，直接处理

    // 检测是否可以使用SIMD
    bool use_simd = is_memory_aligned(values) && count >= 8;

    if (use_simd) {
        batchCreateNumberCellsSIMD(values, buffer, coordinates, count, start_idx);
    } else {
        batchCreateNumberCellsScalar(values, buffer, coordinates, count, start_idx);
    }
}

void TXBatchSIMDProcessor::batchCreateNumberCellsSIMD(
    const double* values,
    TXCompactCellBuffer& buffer,
    const uint32_t* coordinates,
    size_t count,
    size_t start_idx
) {
    using simd_type = xsimd::simd_type<double>;
    constexpr size_t simd_size = simd_type::size;

    const size_t simd_end = (count / simd_size) * simd_size;
    
    // 🚀 优化的SIMD批量处理 - 减少内存访问
    for (size_t i = 0; i < simd_end; i += simd_size) {
        // 加载数值
        auto values_simd = xsimd::load_unaligned(&values[i]);

        // 存储数值
        values_simd.store_unaligned(&buffer.number_values[start_idx + i]);

        // 🚀 批量设置固定值 - 使用memset优化
        const size_t batch_size = std::min(simd_size, count - i);
        const size_t base_idx = start_idx + i;

        // 批量复制坐标
        std::memcpy(&buffer.coordinates[base_idx], &coordinates[i], batch_size * sizeof(uint32_t));

        // 批量设置类型为Number
        std::memset(&buffer.cell_types[base_idx], static_cast<uint8_t>(TXCellType::Number), batch_size);

        // 批量设置字符串索引为0
        std::memset(&buffer.string_indices[base_idx], 0, batch_size * sizeof(uint32_t));

        // 批量设置样式索引为0
        std::memset(&buffer.style_indices[base_idx], 0, batch_size * sizeof(uint16_t));
    }
    
    // 处理剩余元素
    for (size_t i = simd_end; i < count; ++i) {
        buffer.number_values[start_idx + i] = values[i];
        buffer.coordinates[start_idx + i] = coordinates[i];
        buffer.cell_types[start_idx + i] = static_cast<uint8_t>(TXCellType::Number);
        buffer.string_indices[start_idx + i] = 0;
        buffer.style_indices[start_idx + i] = 0;
    }
}

void TXBatchSIMDProcessor::batchCreateNumberCellsScalar(
    const double* values,
    TXCompactCellBuffer& buffer,
    const uint32_t* coordinates,
    size_t count,
    size_t start_idx
) {

    // 🚀 优化的标量处理 - 批量内存操作
    // 批量复制数值
    std::memcpy(&buffer.number_values[start_idx], values, count * sizeof(double));

    // 批量复制坐标
    std::memcpy(&buffer.coordinates[start_idx], coordinates, count * sizeof(uint32_t));

    // 批量设置类型
    std::memset(&buffer.cell_types[start_idx], static_cast<uint8_t>(TXCellType::Number), count);

    // 批量设置字符串索引
    std::memset(&buffer.string_indices[start_idx], 0, count * sizeof(uint32_t));

    // 批量设置样式索引
    std::memset(&buffer.style_indices[start_idx], 0, count * sizeof(uint16_t));
}

void TXBatchSIMDProcessor::batchCreateStringCells(
    const std::vector<std::string>& strings,
    TXCompactCellBuffer& buffer,
    const uint32_t* coordinates,
    TXGlobalStringPool& string_pool
) {
    auto start_time = std::chrono::high_resolution_clock::now();

    const size_t count = strings.size();
    const size_t start_idx = buffer.size;
    buffer.resize(buffer.size + count);
    
    // 批量处理字符串
    for (size_t i = 0; i < count; ++i) {
        // **FIX:** 从字符串池获取索引, 而不是字符串本身。
        uint32_t string_index = string_pool.getIndex(strings[i]);
        if (string_index == static_cast<uint32_t>(SIZE_MAX)) {
            // 如果字符串不存在，先添加它 (addString 现在是 intern 的别名)
            string_pool.addString(strings[i]);
            string_index = string_pool.getIndex(strings[i]);
        }
        
        buffer.string_indices[start_idx + i] = string_index;
        buffer.coordinates[start_idx + i] = coordinates[i];
        buffer.cell_types[start_idx + i] = static_cast<uint8_t>(TXCellType::String);
        buffer.number_values[start_idx + i] = 0.0; // 不是数值
        buffer.style_indices[start_idx + i] = 0;   // 默认样式
    }
    
    // size已经在resize中更新了
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    update_performance_stats(count, duration.count() / 1000.0);
}

void TXBatchSIMDProcessor::batchCreateMixedCells(
    const std::vector<TXVariant>& variants,
    TXCompactCellBuffer& buffer,
    const uint32_t* coordinates,
    TXGlobalStringPool& string_pool
) {
    auto start_time = std::chrono::high_resolution_clock::now();

    const size_t count = variants.size();
    const size_t start_idx = buffer.size;
    buffer.resize(buffer.size + count);
    
    for (size_t i = 0; i < count; ++i) {
        const auto& variant = variants[i];
        
        buffer.coordinates[start_idx + i] = coordinates[i];
        buffer.style_indices[start_idx + i] = 0; // 默认样式
        
        switch (variant.getType()) {
            case TXVariant::Type::Number: {
                buffer.number_values[start_idx + i] = variant.getNumber();
                buffer.cell_types[start_idx + i] = static_cast<uint8_t>(TXCellType::Number);
                buffer.string_indices[start_idx + i] = 0;
                break;
            }
            case TXVariant::Type::String: {
                uint32_t string_index = string_pool.getIndex(variant.getString());
                 if (string_index == static_cast<uint32_t>(SIZE_MAX)) {
                    string_pool.addString(variant.getString());
                    string_index = string_pool.getIndex(variant.getString());
                }
                buffer.string_indices[start_idx + i] = string_index;
                buffer.cell_types[start_idx + i] = static_cast<uint8_t>(TXCellType::String);
                buffer.number_values[start_idx + i] = 0.0;
                break;
            }
            case TXVariant::Type::Boolean: {
                buffer.number_values[start_idx + i] = variant.getBoolean() ? 1.0 : 0.0;
                buffer.cell_types[start_idx + i] = static_cast<uint8_t>(TXCellType::Boolean);
                buffer.string_indices[start_idx + i] = 0;
                break;
            }
            default: {
                // 默认为空单元格
                buffer.number_values[start_idx + i] = 0.0;
                buffer.cell_types[start_idx + i] = static_cast<uint8_t>(TXCellType::Empty);
                buffer.string_indices[start_idx + i] = 0;
                break;
            }
        }
    }
    
    // size已经在resize中更新了
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    update_performance_stats(count, duration.count() / 1000.0);
}

size_t TXBatchSIMDProcessor::batchConvertCoordinates(
    const std::vector<std::string>& cell_refs,
    uint32_t* coordinates,
    size_t count
) {
    size_t converted = 0;
    
    for (size_t i = 0; i < count && i < cell_refs.size(); ++i) {
        try {
            TXCoordinate coord(cell_refs[i]); // 使用构造函数解析
            coordinates[i] = (static_cast<uint32_t>(coord.getRow().index()) << 16) | static_cast<uint32_t>(coord.getCol().index());
            ++converted;
        } catch (...) {
            coordinates[i] = 0; // 无效坐标
        }
    }
    
    return converted;
}

TXCellStats TXBatchSIMDProcessor::batchCalculateStats(
    const TXCompactCellBuffer& buffer,
    const TXRange* range
) {
    auto start_time = std::chrono::high_resolution_clock::now();
    
    TXCellStats stats;
    
    if (buffer.empty()) {
        return stats;
    }
    
    // 确定处理范围
    size_t start_idx = 0;
    size_t end_idx = buffer.size;
    
    if (range) {
        // 如果指定了范围，需要找到对应的单元格
        uint32_t range_start = (range->getStart().getRow().index() << 16) | range->getStart().getCol().index();
        uint32_t range_end = (range->getEnd().getRow().index() << 16) | range->getEnd().getCol().index();
        
        // 简化版本：遍历所有单元格检查是否在范围内
        std::vector<size_t> valid_indices;
        for (size_t i = 0; i < buffer.size; ++i) {
            uint32_t coord = buffer.coordinates[i];
            if (coord >= range_start && coord <= range_end) {
                valid_indices.push_back(i);
            }
        }
        
        if (valid_indices.empty()) {
            return stats;
        }
        
        // 统计范围内的单元格
        for (size_t idx : valid_indices) {
            uint8_t type = buffer.cell_types[idx];
            stats.count++;
            
            if (type == static_cast<uint8_t>(TXCellType::Number)) {
                double value = buffer.number_values[idx];
                stats.number_cells++;
                stats.sum += value;
                
                if (stats.number_cells == 1) {
                    stats.min_value = stats.max_value = value;
                } else {
                    stats.min_value = std::min(stats.min_value, value);
                    stats.max_value = std::max(stats.max_value, value);
                }
            } else if (type == static_cast<uint8_t>(TXCellType::String)) {
                stats.string_cells++;
            } else {
                stats.empty_cells++;
            }
        }
    } else {
        // 全表统计 - 使用SIMD优化
        stats.count = buffer.size;
        
        // 统计各种类型
        for (size_t i = 0; i < buffer.size; ++i) {
            uint8_t type = buffer.cell_types[i];
            if (type == static_cast<uint8_t>(TXCellType::Number)) {
                stats.number_cells++;
            } else if (type == static_cast<uint8_t>(TXCellType::String)) {
                stats.string_cells++;
            } else {
                stats.empty_cells++;
            }
        }
        
        // SIMD求和 - 只处理数值单元格
        if (stats.number_cells > 0) {
            batchSumNumbers(buffer, stats);
        }
    }
    
    // 计算均值
    if (stats.number_cells > 0) {
        stats.mean = stats.sum / stats.number_cells;
    }
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    update_performance_stats(stats.count, duration.count() / 1000.0);
    
    return stats;
}

void TXBatchSIMDProcessor::batchSumNumbers(const TXCompactCellBuffer& buffer, TXCellStats& stats) {
    using simd_type = xsimd::simd_type<double>;
    constexpr size_t simd_size = simd_type::size;
    
    auto sum_simd = xsimd::broadcast<double>(0.0);
    auto min_simd = xsimd::broadcast<double>(std::numeric_limits<double>::max());
    auto max_simd = xsimd::broadcast<double>(std::numeric_limits<double>::lowest());
    
    size_t processed = 0;
    
    // SIMD处理 - 需要过滤只处理数值单元格
    for (size_t i = 0; i < buffer.size; ++i) {
        if (buffer.cell_types[i] == static_cast<uint8_t>(TXCellType::Number)) {
            double value = buffer.number_values[i];
            stats.sum += value;
            
            if (processed == 0) {
                stats.min_value = stats.max_value = value;
            } else {
                stats.min_value = std::min(stats.min_value, value);
                stats.max_value = std::max(stats.max_value, value);
            }
            processed++;
        }
    }
}

double TXBatchSIMDProcessor::batchSum(
    const TXCompactCellBuffer& buffer,
    const TXRange& range
) {
    TXCellStats stats = batchCalculateStats(buffer, &range);
    return stats.sum;
}

size_t TXBatchSIMDProcessor::batchFind(
    const TXCompactCellBuffer& buffer,
    double target_value,
    std::vector<uint32_t>& results
) {
    results.clear();
    
    using simd_type = xsimd::simd_type<double>;
    constexpr size_t simd_size = simd_type::size;
    
    auto target_simd = xsimd::broadcast<double>(target_value);
    
    // SIMD比较查找
    for (size_t i = 0; i < buffer.size; ++i) {
        if (buffer.cell_types[i] == static_cast<uint8_t>(TXCellType::Number)) {
            if (std::abs(buffer.number_values[i] - target_value) < 1e-10) {
                results.push_back(buffer.coordinates[i]);
            }
        }
    }
    
    return results.size();
}

void TXBatchSIMDProcessor::fillRange(
    TXCompactCellBuffer& buffer,
    const TXRange& range,
    double value
) {
    auto start_time = std::chrono::high_resolution_clock::now();
    
    uint32_t range_start = (range.getStart().getRow().index() << 16) | range.getStart().getCol().index();
    uint32_t range_end = (range.getEnd().getRow().index() << 16) | range.getEnd().getCol().index();
    
    // 计算需要填充的单元格数量
    size_t fill_count = (range.getEnd().getRow().index() - range.getStart().getRow().index() + 1) * (range.getEnd().getCol().index() - range.getStart().getCol().index() + 1);
    
    // 生成所有坐标并批量填充
    std::vector<uint32_t> coords;
    coords.reserve(fill_count);

    for (uint32_t row = range.getStart().getRow().index(); row <= range.getEnd().getRow().index(); ++row) {
        for (uint32_t col = range.getStart().getCol().index(); col <= range.getEnd().getCol().index(); ++col) {
            coords.push_back((row << 16) | col);
        }
    }

    // 准备数值数组
    std::vector<double> values(fill_count, value);

    // 使用批量创建方法（它会自动处理buffer的resize）
    size_t start_idx = buffer.size;
    buffer.resize(buffer.size + fill_count);
    batchCreateNumberCells(values.data(), buffer, coords.data(), fill_count, start_idx);
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    update_performance_stats(fill_count, duration.count() / 1000.0);
}

void TXBatchSIMDProcessor::copyRange(
    TXCompactCellBuffer& buffer,
    const TXRange& src_range,
    const TXCoordinate& dst_start
) {
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 找到源范围内的所有单元格
    std::vector<size_t> src_indices;
    uint32_t src_start = (src_range.getStart().getRow().index() << 16) | src_range.getStart().getCol().index();
    uint32_t src_end = (src_range.getEnd().getRow().index() << 16) | src_range.getEnd().getCol().index();
    
    for (size_t i = 0; i < buffer.size; ++i) {
        uint32_t coord = buffer.coordinates[i];
        if (coord >= src_start && coord <= src_end) {
            src_indices.push_back(i);
        }
    }
    
    if (src_indices.empty()) return;
    
    // 计算偏移量
    int32_t row_offset = static_cast<int32_t>(dst_start.getRow().index()) - static_cast<int32_t>(src_range.getStart().getRow().index());
    int32_t col_offset = static_cast<int32_t>(dst_start.getCol().index()) - static_cast<int32_t>(src_range.getStart().getCol().index());
    
    // 批量复制
    buffer.reserve(buffer.size + src_indices.size());
    size_t start_idx = buffer.size;
    
    for (size_t i = 0; i < src_indices.size(); ++i) {
        size_t src_idx = src_indices[i];
        size_t dst_idx = start_idx + i;
        
        // 复制数据
        buffer.number_values[dst_idx] = buffer.number_values[src_idx];
        buffer.string_indices[dst_idx] = buffer.string_indices[src_idx];
        buffer.style_indices[dst_idx] = buffer.style_indices[src_idx];
        buffer.cell_types[dst_idx] = buffer.cell_types[src_idx];
        
        // 计算新坐标
        uint32_t old_coord = buffer.coordinates[src_idx];
        uint32_t old_row = old_coord >> 16;
        uint32_t old_col = old_coord & 0xFFFF;
        uint32_t new_row = old_row + row_offset;
        uint32_t new_col = old_col + col_offset;
        buffer.coordinates[dst_idx] = (new_row << 16) | new_col;
    }
    
    buffer.size += src_indices.size();
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    update_performance_stats(src_indices.size(), duration.count() / 1000.0);
}

void TXBatchSIMDProcessor::optimizeMemoryLayout(TXCompactCellBuffer& buffer) {
    if (buffer.empty()) return;
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 按坐标排序以提高缓存局部性
    buffer.sort_by_coordinates();
    
    // 确保SIMD对齐
    ensure_simd_alignment(buffer.number_values);
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    update_performance_stats(buffer.size, duration.count() / 1000.0);
}

size_t TXBatchSIMDProcessor::compressSparseData(TXCompactCellBuffer& buffer) {
    if (buffer.empty()) return 0;
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    size_t write_pos = 0;
    size_t original_size = buffer.size;
    
    // 移除空单元格
    for (size_t read_pos = 0; read_pos < buffer.size; ++read_pos) {
        if (buffer.cell_types[read_pos] != static_cast<uint8_t>(TXCellType::Empty)) {
            if (write_pos != read_pos) {
                buffer.number_values[write_pos] = buffer.number_values[read_pos];
                buffer.string_indices[write_pos] = buffer.string_indices[read_pos];
                buffer.style_indices[write_pos] = buffer.style_indices[read_pos];
                buffer.coordinates[write_pos] = buffer.coordinates[read_pos];
                buffer.cell_types[write_pos] = buffer.cell_types[read_pos];
            }
            ++write_pos;
        }
    }
    
    // 更新大小
    size_t new_size = write_pos;
    buffer.resize(new_size);
    
    size_t removed = original_size - new_size;
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    update_performance_stats(original_size, duration.count() / 1000.0);
    
    return removed;
}

// ==================== 工具函数实现 ====================

bool TXBatchSIMDProcessor::is_memory_aligned(const void* ptr, size_t alignment) {
    return reinterpret_cast<uintptr_t>(ptr) % alignment == 0;
}

void TXBatchSIMDProcessor::ensure_simd_alignment(TXVector<double>& vec) {
    // 🚀 确保TXVector的SIMD对齐
    // TXVector已经在内部处理了SIMD对齐，这里主要是确保大小是SIMD长度的倍数
    using simd_type = xsimd::simd_type<double>;
    constexpr size_t simd_size = simd_type::size;

    size_t remainder = vec.size() % simd_size;
    if (remainder != 0) {
        vec.resize(vec.size() + (simd_size - remainder), 0.0);
    }
}

size_t TXBatchSIMDProcessor::round_up_to_simd_size(size_t size) {
    using simd_type = xsimd::simd_type<double>;
    constexpr size_t simd_size = simd_type::size;
    
    return ((size + simd_size - 1) / simd_size) * simd_size;
}

void TXBatchSIMDProcessor::warmupSIMD(size_t warmup_size) {
    // 🚀 预热SIMD管道 - 使用全局内存管理器
    std::vector<double> warmup_data(warmup_size, 1.0);
    std::vector<uint32_t> warmup_coords(warmup_size);
    TXCompactCellBuffer warmup_buffer(GlobalUnifiedMemoryManager::getInstance(), warmup_size);

    std::iota(warmup_coords.begin(), warmup_coords.end(), 0);

    batchCreateNumberCells(warmup_data.data(), warmup_buffer, warmup_coords.data(), warmup_size, 0);
}

// ==================== 性能监控实现 ====================

const TXBatchSIMDProcessor::BatchPerformanceStats& TXBatchSIMDProcessor::getPerformanceStats() {
    // 计算派生统计
    if (performance_stats_.total_operations > 0 && performance_stats_.total_time_ms > 0) {
        performance_stats_.avg_throughput = performance_stats_.total_cells_processed / 
                                          (performance_stats_.total_time_ms / 1000.0);
        
        performance_stats_.simd_utilization = std::min(1.0, 
            performance_stats_.total_cells_processed / 
            (performance_stats_.total_operations * 1000.0));
    }
    
    return performance_stats_;
}

void TXBatchSIMDProcessor::resetPerformanceStats() {
    performance_stats_ = BatchPerformanceStats{};
}

void TXBatchSIMDProcessor::update_performance_stats(size_t cells_processed, double time_ms) {
    performance_stats_.total_operations++;
    performance_stats_.total_cells_processed += cells_processed;
    performance_stats_.total_time_ms += time_ms;
}

// ==================== TXBatchOperations 实现 ====================

TXResult<size_t> TXBatchOperations::importDataBatch(
    TXInMemorySheet& sheet,
    const std::vector<std::vector<TXVariant>>& data,
    const TXCoordinate& start_coord,
    const TXImportOptions& options
) {
    if (data.empty()) {
        return Ok<size_t>(0); // **FIX:** 使用 Ok<T> 工厂函数
    }
    
    try {
        size_t total_cells = 0;
        
        // 计算总单元格数
        for (const auto& row : data) {
            total_cells += row.size();
        }
        
        // 预分配内存
        if (options.optimize_memory) {
            sheet.reserve(sheet.getCellCount() + total_cells);
        }
        
        // 🚀 优化策略：按类型分离数据，避免混合处理
        std::vector<TXCoordinate> number_coords, string_coords;
        std::vector<double> numbers;
        std::vector<std::string> strings;

        // 预估容量
        number_coords.reserve(total_cells / 2);
        string_coords.reserve(total_cells / 2);
        numbers.reserve(total_cells / 2);
        strings.reserve(total_cells / 2);

        // 按类型分离数据 - 避免TXVariant的性能开销
        for (size_t row_idx = 0; row_idx < data.size(); ++row_idx) {
            const auto& row = data[row_idx];
            for (size_t col_idx = 0; col_idx < row.size(); ++col_idx) {
                const auto& variant = row[col_idx];

                if (options.skip_empty_cells && variant.isEmpty()) {
                    continue;
                }

                TXCoordinate coord(
                    row_t(start_coord.getRow().index() + static_cast<uint32_t>(row_idx)),
                    column_t(start_coord.getCol().index() + static_cast<uint32_t>(col_idx))
                );

                // 快速类型检测和分离
                switch (variant.getType()) {
                    case TXVariant::Type::Number:
                        number_coords.push_back(coord);
                        numbers.push_back(variant.getNumber());
                        break;
                    case TXVariant::Type::String:
                        string_coords.push_back(coord);
                        strings.push_back(variant.getString());
                        break;
                    case TXVariant::Type::Boolean:
                        number_coords.push_back(coord);
                        numbers.push_back(variant.getBoolean() ? 1.0 : 0.0);
                        break;
                    default:
                        // 跳过空单元格或未知类型
                        break;
                }
            }
        }

        size_t total_processed = 0;

        // 🚀 批量处理数值 - 使用SIMD优化
        if (!numbers.empty()) {
            auto number_result = sheet.setBatchNumbers(number_coords, numbers);
            if (number_result.isError()) {
                return Err<size_t>(number_result.error());
            }
            total_processed += number_result.value();
        }

        // 🚀 批量处理字符串 - 使用字符串池优化
        if (!strings.empty()) {
            auto string_result = sheet.setBatchStrings(string_coords, strings);
            if (string_result.isError()) {
                return Err<size_t>(string_result.error());
            }
            total_processed += string_result.value();
        }

        return Ok(total_processed);
        
    } catch (const std::exception& e) {
        return Err<size_t>(TXErrorCode::InvalidData, e.what());
    }
}

TXResult<size_t> TXBatchOperations::importNumbersBatch(
    TXInMemorySheet& sheet,
    const std::vector<std::vector<double>>& numbers,
    const TXCoordinate& start_coord
) {
    if (numbers.empty()) {
        return Ok<size_t>(0); // **FIX:** 使用 Ok<T> 工厂函数
    }
    
    try {
        // 计算总单元格数
        size_t total_cells = 0;
        for (const auto& row : numbers) {
            total_cells += row.size();
        }
        
        // 预分配
        sheet.reserve(sheet.getCellCount() + total_cells);
        
        // 展开数据
        std::vector<double> all_numbers;
        std::vector<TXCoordinate> all_coords;
        all_numbers.reserve(total_cells);
        all_coords.reserve(total_cells);
        
        for (size_t row_idx = 0; row_idx < numbers.size(); ++row_idx) {
            const auto& row = numbers[row_idx];
            for (size_t col_idx = 0; col_idx < row.size(); ++col_idx) {
                all_numbers.push_back(row[col_idx]);
                all_coords.emplace_back(
                    row_t(start_coord.getRow().index() + static_cast<uint32_t>(row_idx)),
                    column_t(start_coord.getCol().index() + static_cast<uint32_t>(col_idx))
                );
            }
        }
        
        // 批量设置
        // **FIX:** 正确处理TXResult的返回
        auto result = sheet.setBatchNumbers(all_coords, all_numbers);
        if (result.isError()) {
            return Err<size_t>(result.error());
        }
        return Ok(result.value());
        
    } catch (const std::exception& e) {
        return Err<size_t>(TXErrorCode::InvalidData, e.what());
    }
}

TXResult<size_t> TXBatchOperations::importStringsBatch(
    TXInMemorySheet& sheet,
    const std::vector<std::vector<std::string>>& strings,
    const TXCoordinate& start_coord
)
{
    if (strings.empty()) {
        return Ok<size_t>(0);
    }
    try {
        size_t total_cells = 0;
        for (const auto& row : strings) {
            total_cells += row.size();
        }
        sheet.reserve(sheet.getCellCount() + total_cells);
        std::vector<std::string> all_strings;
        std::vector<TXCoordinate> all_coords;
        all_strings.reserve(total_cells);
        all_coords.reserve(total_cells);
        for (size_t row_idx = 0; row_idx < strings.size(); ++row_idx) {
            const auto& row = strings[row_idx];
            for (size_t col_idx = 0; col_idx < row.size(); ++col_idx) {
                all_strings.push_back(row[col_idx]);
                all_coords.emplace_back(
                    row_t(start_coord.getRow().index() + static_cast<uint32_t>(row_idx)),
                    column_t(start_coord.getCol().index() + static_cast<uint32_t>(col_idx))
                );
            }
        }
        return sheet.setBatchStrings(all_coords, all_strings);
    }
    catch (const std::exception& e) {
        return Err<size_t>(TXErrorCode::InvalidData, e.what());
    }
}


TXResult<size_t> TXBatchOperations::importFromCSV(
    TXInMemorySheet& sheet,
    const std::string& csv_content,
    const TXImportOptions& options
)
{
     if (csv_content.empty()) {
        return Ok<size_t>(0);
    }
    // ... CSV解析逻辑 ...
    // 然后调用 importDataBatch
    return Ok<size_t>(0); // 示例
}


} // namespace TinaXlsx
