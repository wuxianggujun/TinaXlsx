//
// @file TXInMemorySheet.cpp
// @brief 内存优先工作表实现 - 完全内存中操作，极致性能
//

#include "TinaXlsx/TXInMemorySheet.hpp"
#include "TinaXlsx/TXZeroCopySerializer.hpp"
#include "TinaXlsx/TXBatchSIMDProcessor.hpp"
#include "TinaXlsx/TXGlobalStringPool.hpp"
#include "TinaXlsx/TXUnifiedMemoryManager.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include "TinaXlsx/TXXMLTemplates.hpp"
#include <algorithm>
#include <chrono>
#include <sstream>
#include <fstream>
#include <map>
#include <cstdio>
#include <cstdlib>

namespace TinaXlsx {

// ==================== TXMemoryLayoutOptimizer 实现 ====================

void TXMemoryLayoutOptimizer::optimizeForSequentialAccess(TXCompactCellBuffer& buffer) {
    if (buffer.empty()) return;
    
    // 按坐标排序以提高顺序访问性能
    buffer.sort_by_coordinates();
}

void TXMemoryLayoutOptimizer::optimizeForExcelAccess(TXCompactCellBuffer& buffer) {
    if (buffer.empty()) return;
    
    // Excel按行访问，优化行的局部性
    optimizeForSequentialAccess(buffer);
}

void TXMemoryLayoutOptimizer::optimizeForSIMD(TXCompactCellBuffer& buffer) {
    if (buffer.empty()) return;
    
    // 确保SIMD对齐
    TXBatchSIMDProcessor::optimizeMemoryLayout(buffer);
}

std::vector<TXRowGroup> TXMemoryLayoutOptimizer::generateRowGroups(const TXCompactCellBuffer& buffer) {
    std::vector<TXRowGroup> row_groups;
    
    if (buffer.empty()) return row_groups;
    
    // 按行分组
    std::map<uint32_t, std::vector<size_t>> rows_map;
    
    for (size_t i = 0; i < buffer.size; ++i) {
        uint32_t coord = buffer.coordinates[i];
        uint32_t row = coord >> 16;
        rows_map[row].push_back(i);
    }
    
    // 生成行分组
    for (const auto& [row_index, cell_indices] : rows_map) {
        if (!cell_indices.empty()) {
            TXRowGroup group;
            group.row_index = row_index;
            group.start_cell_index = cell_indices[0];
            group.cell_count = cell_indices.size();
            row_groups.push_back(group);
        }
    }
    
    return row_groups;
}

// ==================== TXInMemorySheet 实现 ====================

TXInMemorySheet::TXInMemorySheet(
    const std::string& name,
    TXUnifiedMemoryManager& memory_manager,
    TXGlobalStringPool& string_pool
) : cell_buffer_(memory_manager)  // 🚀 使用内存管理器初始化cell_buffer_
  , memory_manager_(memory_manager)
  , string_pool_(string_pool)
  , name_(name)
  , optimizer_(std::make_unique<TXMemoryLayoutOptimizer>()) {

    // 初始化性能统计
    stats_ = {};
}

TXInMemorySheet::~TXInMemorySheet() = default;

TXInMemorySheet::TXInMemorySheet(TXInMemorySheet&& other) noexcept
    : cell_buffer_(std::move(other.cell_buffer_))
    , memory_manager_(other.memory_manager_)
    , string_pool_(other.string_pool_)
    , optimizer_(std::move(other.optimizer_))
    , coord_to_index_(std::move(other.coord_to_index_))
    , name_(std::move(other.name_))
    , max_row_(other.max_row_)
    , max_col_(other.max_col_)
    , dirty_(other.dirty_)
    , auto_optimize_(other.auto_optimize_)
    , stats_(other.stats_) {
}

TXInMemorySheet& TXInMemorySheet::operator=(TXInMemorySheet&& other) noexcept {
    if (this != &other) {
        cell_buffer_ = std::move(other.cell_buffer_);
        optimizer_ = std::move(other.optimizer_);
        coord_to_index_ = std::move(other.coord_to_index_);
        name_ = std::move(other.name_);
        max_row_ = other.max_row_;
        max_col_ = other.max_col_;
        dirty_ = other.dirty_;
        auto_optimize_ = other.auto_optimize_;
        stats_ = other.stats_;
    }
    return *this;
}

// ==================== 批量操作接口实现 ====================

TXResult<size_t> TXInMemorySheet::setBatchNumbers(
    const std::vector<TXCoordinate>& coords, 
    const std::vector<double>& values
) {
    auto start_time = std::chrono::high_resolution_clock::now();
    
    if (coords.size() != values.size()) {
        return TXResult<size_t>(TXError(TXErrorCode::InvalidArgument, "坐标和数值数组大小不匹配"));
    }
    
    if (coords.empty()) {
        return TXResult<size_t>(static_cast<size_t>(0));
    }
    
    try {
        // 🚀 极致性能优化：使用您的内存管理器进行高性能分配
        const size_t count = coords.size();
        const size_t old_size = cell_buffer_.size;

        // 🚀 为了极致性能，跳过监控系统
        // memory_manager_.startMonitoring(); // 开始监控性能

        // 🚀 使用内存管理器智能预分配所有内存
        size_t new_size = old_size + count;

        // 预分配额外空间以减少后续分配
        size_t growth_factor = std::max(count, new_size / 4); // 25%增长或当前批次大小
        size_t target_capacity = new_size + growth_factor;

        if (cell_buffer_.capacity < target_capacity) {
            cell_buffer_.reserve(target_capacity);
        }

        cell_buffer_.resize(new_size);

        // 🚀 零开销批量转换：使用您的内存管理器分配高性能临时缓冲区
        size_t bytes_needed = count * sizeof(uint32_t);
        uint32_t* packed_coords = static_cast<uint32_t*>(memory_manager_.allocate(bytes_needed));
        if (!packed_coords) {
            return TXResult<size_t>(TXError(TXErrorCode::MemoryError, "内存分配失败"));
        }

        // 🚀 批量转换坐标并更新边界（合并循环减少开销）
        uint32_t current_max_row = max_row_;
        uint32_t current_max_col = max_col_;

        for (size_t i = 0; i < count; ++i) {
            const auto& coord = coords[i];
            const uint32_t row = coord.getRow().index();
            const uint32_t col = coord.getCol().index();

            // 批量更新边界
            if (row > current_max_row) current_max_row = row;
            if (col > current_max_col) current_max_col = col;

            // 转换坐标
            packed_coords[i] = (row << 16) | col;
        }

        // 一次性更新边界
        max_row_ = current_max_row;
        max_col_ = current_max_col;

        // 🚀 超高性能SIMD处理
        TXBatchSIMDProcessor::batchCreateNumberCells(
            values.data(), cell_buffer_, packed_coords, count, old_size);

        // 🚀 批量更新索引：使用reserve避免重复分配
        coord_to_index_.reserve(coord_to_index_.size() + count);
        for (size_t i = 0; i < count; ++i) {
            coord_to_index_[packed_coords[i]] = old_size + i;
        }

        // 🚀 释放临时缓冲区
        memory_manager_.deallocate(packed_coords);

        dirty_ = true;

        // 🚀 跳过智能清理和监控以获得极致性能
        // auto cleanup_bytes = memory_manager_.smartCleanup();
        // memory_manager_.stopMonitoring();

        // 跳过maybeOptimize()以获得极致性能

        return TXResult<size_t>(count);
        
    } catch (const std::exception& e) {
        return TXResult<size_t>(TXError(TXErrorCode::MemoryError, 
                                      std::string("批量数值设置失败: ") + e.what()));
    }
}

TXResult<size_t> TXInMemorySheet::setBatchStrings(
    const std::vector<TXCoordinate>& coords, 
    const std::vector<std::string>& values
) {
    auto start_time = std::chrono::high_resolution_clock::now();
    
    if (coords.size() != values.size()) {
        return TXResult<size_t>(TXError(TXErrorCode::InvalidParameter, "坐标和字符串数组大小不匹配"));
    }
    
    if (coords.empty()) {
        return TXResult<size_t>(static_cast<size_t>(0));
    }
    
    try {
        // 转换坐标格式
        std::vector<uint32_t> packed_coords(coords.size());
        for (size_t i = 0; i < coords.size(); ++i) {
            packed_coords[i] = coordToKey(coords[i]);
            updateBounds(coords[i]);
        }
        
        // 使用SIMD批量处理
        size_t old_size = cell_buffer_.size;
        TXBatchSIMDProcessor::batchCreateStringCells(
            values, cell_buffer_, packed_coords.data(), string_pool_);
        
        // 更新索引
        for (size_t i = 0; i < coords.size(); ++i) {
            updateIndex(coords[i], old_size + i);
        }
        
        dirty_ = true;
        maybeOptimize();
        
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
        updateStats(coords.size(), duration.count() / 1000.0);
        
        return TXResult<size_t>(coords.size());
        
    } catch (const std::exception& e) {
        return TXResult<size_t>(TXError(TXErrorCode::MemoryError, 
                                      std::string("批量字符串设置失败: ") + e.what()));
    }
}

TXResult<size_t> TXInMemorySheet::setBatchStyles(
    const std::vector<TXCoordinate>& coords, 
    const std::vector<uint16_t>& style_indices
) {
    if (coords.size() != style_indices.size()) {
        return TXResult<size_t>(TXError(TXErrorCode::InvalidParameter, "坐标和样式数组大小不匹配"));
    }
    
    try {
        size_t applied = 0;
        
        for (size_t i = 0; i < coords.size(); ++i) {
            uint32_t key = coordToKey(coords[i]);
            auto it = coord_to_index_.find(key);
            
            if (it != coord_to_index_.end()) {
                cell_buffer_.style_indices[it->second] = style_indices[i];
                ++applied;
            }
        }
        
        if (applied > 0) {
            dirty_ = true;
        }
        
        return TXResult<size_t>(applied);
        
    } catch (const std::exception& e) {
        return TXResult<size_t>(TXError(TXErrorCode::InvalidData, 
                                      std::string("批量样式设置失败: ") + e.what()));
    }
}

TXResult<size_t> TXInMemorySheet::setBatchMixed(
    const std::vector<TXCoordinate>& coords,
    const std::vector<TXVariant>& variants
) {
    auto start_time = std::chrono::high_resolution_clock::now();
    
    if (coords.size() != variants.size()) {
        return TXResult<size_t>(TXError(TXErrorCode::InvalidParameter, "坐标和数据数组大小不匹配"));
    }
    
    if (coords.empty()) {
        return TXResult<size_t>(static_cast<size_t>(0));
    }
    
    try {
        // 转换坐标格式
        std::vector<uint32_t> packed_coords(coords.size());
        for (size_t i = 0; i < coords.size(); ++i) {
            packed_coords[i] = coordToKey(coords[i]);
            updateBounds(coords[i]);
        }
        
        // 使用SIMD批量处理
        size_t old_size = cell_buffer_.size;
        TXBatchSIMDProcessor::batchCreateMixedCells(
            variants, cell_buffer_, packed_coords.data(), string_pool_);
        
        // 更新索引
        for (size_t i = 0; i < coords.size(); ++i) {
            updateIndex(coords[i], old_size + i);
        }
        
        dirty_ = true;
        maybeOptimize();
        
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
        updateStats(coords.size(), duration.count() / 1000.0);
        
        return TXResult<size_t>(coords.size());
        
    } catch (const std::exception& e) {
        return TXResult<size_t>(TXError(TXErrorCode::MemoryError, 
                                      std::string("批量混合数据设置失败: ") + e.what()));
    }
}

// ==================== SIMD优化的范围操作 ====================

TXResult<size_t> TXInMemorySheet::fillRange(const TXRange& range, double value) {
    if (!range.isValid()) {
        return TXResult<size_t>(TXError(TXErrorCode::InvalidRange, "无效的范围"));
    }
    
    auto start_coord = range.getStart();
    auto end_coord = range.getEnd();
    
    try {
        std::vector<TXCoordinate> coords;
        coords.reserve(static_cast<size_t>(range.getCellCount()));
        
        for (uint32_t row = start_coord.getRow().index(); row <= end_coord.getRow().index(); ++row) {
            for (uint32_t col = start_coord.getCol().index(); col <= end_coord.getCol().index(); ++col) {
                coords.emplace_back(row_t(row), column_t(col));
                updateBounds(coords.back());
            }
        }
        
        std::vector<double> values(coords.size(), value);
        return setBatchNumbers(coords, values);
        
    } catch (const std::exception& e) {
        return TXResult<size_t>(TXError(TXErrorCode::OperationFailed, 
                                      std::string("范围填充失败: ") + e.what()));
    }
}

TXResult<size_t> TXInMemorySheet::fillRange(const TXRange& range, const std::string& value) {
    // 转换为setBatchStrings调用
    std::vector<TXCoordinate> coords;
    std::vector<std::string> values;
    
    for (uint32_t row = range.getStart().getRow().index(); row <= range.getEnd().getRow().index(); ++row) {
        for (uint32_t col = range.getStart().getCol().index(); col <= range.getEnd().getCol().index(); ++col) {
            coords.emplace_back(row_t(row), column_t(col));
            values.push_back(value);
        }
    }
    
    return setBatchStrings(coords, values);
}

TXResult<size_t> TXInMemorySheet::copyRange(const TXRange& src_range, const TXCoordinate& dst_start) {
    auto start_time = std::chrono::high_resolution_clock::now();
    
    try {
        size_t old_size = cell_buffer_.size;
        TXBatchSIMDProcessor::copyRange(cell_buffer_, src_range, dst_start);
        size_t copied = cell_buffer_.size - old_size;
        
        // 更新边界和索引（简化版本）
        int32_t row_offset = dst_start.getRow().index() - src_range.getStart().getRow().index();
        int32_t col_offset = dst_start.getCol().index() - src_range.getStart().getCol().index();
        
        size_t index = old_size;
        for (uint32_t row = src_range.getStart().getRow().index(); row <= src_range.getEnd().getRow().index(); ++row) {
            for (uint32_t col = src_range.getStart().getCol().index(); col <= src_range.getEnd().getCol().index(); ++col) {
                TXCoordinate dst_coord(row_t(row + row_offset), column_t(col + col_offset));
                updateBounds(dst_coord);
                updateIndex(dst_coord, index++);
            }
        }
        
        dirty_ = true;
        
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
        updateStats(copied, duration.count() / 1000.0);
        
        return TXResult<size_t>(copied);
        
    } catch (const std::exception& e) {
        return TXResult<size_t>(TXError(TXErrorCode::InvalidData, 
                                      std::string("范围复制失败: ") + e.what()));
    }
}

TXResult<size_t> TXInMemorySheet::clearRange(const TXRange& range) {
    try {
        size_t cleared = 0;
        
        // 找到范围内的所有单元格并标记为删除
        std::vector<size_t> to_remove;
        uint32_t range_start = (range.getStart().getRow().index() << 16) | range.getStart().getCol().index();
        uint32_t range_end = (range.getEnd().getRow().index() << 16) | range.getEnd().getCol().index();
        
        for (size_t i = 0; i < cell_buffer_.size; ++i) {
            uint32_t coord = cell_buffer_.coordinates[i];
            if (coord >= range_start && coord <= range_end) {
                to_remove.push_back(i);
            }
        }
        
        // 从后往前删除，避免索引变化
        std::reverse(to_remove.begin(), to_remove.end());
        for (size_t idx : to_remove) {
            // 移除单元格（简化版本：标记为空）
            cell_buffer_.cell_types[idx] = static_cast<uint8_t>(TXCellType::Empty);
            cleared++;
        }
        
        if (cleared > 0) {
            dirty_ = true;
            // 压缩空白单元格
            TXBatchSIMDProcessor::compressSparseData(cell_buffer_);
            // 重建索引
            coord_to_index_.clear();
            for (size_t i = 0; i < cell_buffer_.size; ++i) {
                coord_to_index_[cell_buffer_.coordinates[i]] = i;
            }
        }
        
        return TXResult<size_t>(cleared);
        
    } catch (const std::exception& e) {
        return TXResult<size_t>(TXError(TXErrorCode::InvalidData, 
                                      std::string("范围清除失败: ") + e.what()));
    }
}

// ==================== 单个单元格操作 ====================

TXResult<void> TXInMemorySheet::setNumber(const TXCoordinate& coord, double value) {
    try {
        uint32_t key = coordToKey(coord);
        auto it = coord_to_index_.find(key);
        
        if (it != coord_to_index_.end()) {
            // 更新现有单元格
            size_t index = it->second;
            cell_buffer_.number_values[index] = value;
            cell_buffer_.cell_types[index] = static_cast<uint8_t>(TXCellType::Number);
        } else {
            // 创建新单元格
            if (cell_buffer_.size >= cell_buffer_.capacity) {
                return TXResult<void>(TXError(TXErrorCode::MemoryError, "单元格缓冲区已满"));
            }
            
            size_t index = cell_buffer_.size++;
            cell_buffer_.coordinates[index] = key;
            cell_buffer_.number_values[index] = value;
            cell_buffer_.cell_types[index] = static_cast<uint8_t>(TXCellType::Number);
            cell_buffer_.style_indices[index] = 0;
            
            coord_to_index_[key] = index;
            updateBounds(coord);
        }
        
        dirty_ = true;
        return TXResult<void>();
        
    } catch (const std::exception& e) {
        return TXResult<void>(TXError(TXErrorCode::MemoryError, 
                                     std::string("设置数值失败: ") + e.what()));
    }
}

TXResult<void> TXInMemorySheet::setString(const TXCoordinate& coord, const std::string& value) {
    try {
        const std::string& interned_str = string_pool_.intern(value);
        
        uint32_t key = coordToKey(coord);
        auto it = coord_to_index_.find(key);
        
        if (it != coord_to_index_.end()) {
            // 更新现有单元格
            size_t index = it->second;
            string_pool_.intern(interned_str); // 确保字符串在池中
            cell_buffer_.string_indices[index] = static_cast<uint32_t>(string_pool_.getIndex(interned_str));
            cell_buffer_.cell_types[index] = static_cast<uint8_t>(TXCellType::String);
        } else {
            // 创建新单元格
            if (cell_buffer_.size >= cell_buffer_.capacity) {
                return TXResult<void>(TXError(TXErrorCode::MemoryError, "单元格缓冲区已满"));
            }
            
            size_t index = cell_buffer_.size++;
            cell_buffer_.coordinates[index] = key;
            string_pool_.intern(interned_str); // 确保字符串在池中
            cell_buffer_.string_indices[index] = static_cast<uint32_t>(string_pool_.getIndex(interned_str));
            cell_buffer_.cell_types[index] = static_cast<uint8_t>(TXCellType::String);
            cell_buffer_.style_indices[index] = 0;
            
            coord_to_index_[key] = index;
            updateBounds(coord);
        }
        
        dirty_ = true;
        return TXResult<void>();
        
    } catch (const std::exception& e) {
        return TXResult<void>(TXError(TXErrorCode::MemoryError, 
                                     std::string("设置字符串失败: ") + e.what()));
    }
}

TXResult<TXVariant> TXInMemorySheet::getValue(const TXCoordinate& coord) const {
    uint32_t key = coordToKey(coord);
    auto it = coord_to_index_.find(key);
    
    if (it == coord_to_index_.end()) {
        return TXResult<TXVariant>(TXVariant()); // 空单元格
    }
    
    size_t index = it->second;
    uint8_t type = cell_buffer_.cell_types[index];
    
    switch (static_cast<TXCellType>(type)) {
        case TXCellType::Number:
            return TXResult<TXVariant>(TXVariant(cell_buffer_.number_values[index]));
        case TXCellType::String: {
            // 这里应该从字符串池获取实际字符串
            uint32_t string_index = cell_buffer_.string_indices[index];
            std::string str = string_pool_.getString(string_index);
            return TXResult<TXVariant>(TXVariant(str));
        }
        case TXCellType::Boolean:
            return TXResult<TXVariant>(TXVariant(cell_buffer_.number_values[index] != 0.0));
        default:
            return TXResult<TXVariant>(TXVariant()); // 空单元格
    }
}

bool TXInMemorySheet::hasCell(const TXCoordinate& coord) const {
    uint32_t key = coordToKey(coord);
    return coord_to_index_.find(key) != coord_to_index_.end();
}

// ==================== 高级批量数据导入 ====================

TXResult<size_t> TXInMemorySheet::importData(
    const std::vector<std::vector<TXVariant>>& data,
    const TXCoordinate& start_coord,
    const TXImportOptions& options
) {
    return TXBatchOperations::importDataBatch(*this, data, start_coord, options);
}

TXResult<size_t> TXInMemorySheet::importNumbers(
    const std::vector<std::vector<double>>& numbers,
    const TXCoordinate& start_coord
) {
    return TXBatchOperations::importNumbersBatch(*this, numbers, start_coord);
}

TXResult<size_t> TXInMemorySheet::importFromCSV(
    const std::string& csv_content,
    const TXImportOptions& options
) {
    try {
        // 简单的CSV解析
        std::vector<std::vector<TXVariant>> data;
        std::istringstream stream(csv_content);
        std::string line;
        
        while (std::getline(stream, line)) {
            std::vector<TXVariant> row;
            std::istringstream line_stream(line);
            std::string cell;
            
            while (std::getline(line_stream, cell, ',')) {
                // 简单的类型检测
                if (options.auto_detect_types) {
                    try {
                        double num = std::stod(cell);
                        row.emplace_back(num);
                    } catch (...) {
                        row.emplace_back(cell);
                    }
                } else {
                    row.emplace_back(cell);
                }
            }
            
            if (!row.empty()) {
                data.push_back(std::move(row));
            }
        }
        
        return importData(data, TXCoordinate(row_t(1), column_t(1)), options);
        
    } catch (const std::exception& e) {
        return TXResult<size_t>(TXError(TXErrorCode::InvalidData, 
                                      std::string("CSV导入失败: ") + e.what()));
    }
}

// ==================== 统计和查询 ====================

TXCellStats TXInMemorySheet::getStats(const TXRange* range) const {
    return TXBatchSIMDProcessor::batchCalculateStats(cell_buffer_, range);
}

TXResult<double> TXInMemorySheet::sum(const TXRange& range) const {
    try {
        double result = TXBatchSIMDProcessor::batchSum(cell_buffer_, range);
        return TXResult<double>(result);
    } catch (const std::exception& e) {
        return TXResult<double>(TXError(TXErrorCode::OperationFailed, 
                                      std::string("求和失败: ") + e.what()));
    }
}

std::vector<TXCoordinate> TXInMemorySheet::findValue(
    double target_value,
    const TXRange* range
) const {
    std::vector<uint32_t> found_coords;
    TXBatchSIMDProcessor::batchFind(cell_buffer_, target_value, found_coords);
    
    std::vector<TXCoordinate> result;
    for (uint32_t coord : found_coords) {
        result.push_back(keyToCoord(coord));
    }
    
    return result;
}

// ==================== 内存和性能优化 ====================

void TXInMemorySheet::optimizeMemoryLayout() {
    if (cell_buffer_.empty()) return;
    
    optimizer_->optimizeForExcelAccess(cell_buffer_);
    
    // 重建索引
    coord_to_index_.clear();
    for (size_t i = 0; i < cell_buffer_.size; ++i) {
        coord_to_index_[cell_buffer_.coordinates[i]] = i;
    }
}

size_t TXInMemorySheet::compressSparseData() {
    size_t removed = TXBatchSIMDProcessor::compressSparseData(cell_buffer_);
    
    if (removed > 0) {
        // 重建索引
        coord_to_index_.clear();
        for (size_t i = 0; i < cell_buffer_.size; ++i) {
            coord_to_index_[cell_buffer_.coordinates[i]] = i;
        }
    }
    
    return removed;
}

void TXInMemorySheet::reserve(size_t estimated_cells) {
    cell_buffer_.reserve(estimated_cells);
    coord_to_index_.reserve(estimated_cells);
}

void TXInMemorySheet::shrink_to_fit() {
    cell_buffer_.shrink_to_fit();
}

// ==================== 序列化和导出 ====================

TXResult<void> TXInMemorySheet::serializeToMemory(TXZeroCopySerializer& serializer) const {
    return serializer.serializeWorksheet(*this);
}

TXResult<std::string> TXInMemorySheet::exportToCSV(const TXRange* range) const {
    try {
        std::ostringstream csv;
        
        TXRange export_range = range ? *range : getUsedRange();
        
        for (uint32_t row = export_range.getStart().getRow().index(); row <= export_range.getEnd().getRow().index(); ++row) {
            bool first_col = true;
            for (uint32_t col = export_range.getStart().getCol().index(); col <= export_range.getEnd().getCol().index(); ++col) {
                if (!first_col) csv << ",";
                first_col = false;
                
                TXCoordinate coord{row_t(row), column_t(col)};
                auto value_result = getValue(coord);
                if (value_result.isOk()) {
                    const auto& variant = value_result.value();
                    if (variant.getType() == TXVariant::Type::String) {
                        csv << "\"" << variant.getString() << "\"";
                    } else if (variant.getType() == TXVariant::Type::Number) {
                        csv << variant.getNumber();
                    }
                }
            }
            csv << "\n";
        }
        
        return TXResult<std::string>(csv.str());
        
    } catch (const std::exception& e) {
        return TXResult<std::string>(TXError(TXErrorCode::SerializationError, 
                                           std::string("CSV导出失败: ") + e.what()));
    }
}

// ==================== 元数据和属性 ====================

TXRange TXInMemorySheet::getUsedRange() const {
    if (cell_buffer_.empty()) {
        return TXRange(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(1), column_t(1)));
    }
    
    return TXRange(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(max_row_), column_t(max_col_)));
}

// ==================== 性能监控 ====================

TXInMemorySheet::SheetPerformanceStats TXInMemorySheet::getPerformanceStats() const {
    SheetPerformanceStats stats;
    stats.total_cells = stats_.total_cells;
    stats.batch_operations = stats_.batch_operations;
    stats.avg_operation_time = stats_.batch_operations > 0 ? 
        stats_.total_operation_time / stats_.batch_operations : 0.0;
    stats.cache_hit_ratio = stats_.cache_hits + stats_.cache_misses > 0 ?
        static_cast<double>(stats_.cache_hits) / (stats_.cache_hits + stats_.cache_misses) : 0.0;
    stats.memory_usage = cell_buffer_.capacity * sizeof(double) * 5; // 粗略估算
    stats.compression_ratio = cell_buffer_.size > 0 ? 
        static_cast<double>(cell_buffer_.size) / cell_buffer_.capacity : 1.0;
    
    return stats;
}

void TXInMemorySheet::resetPerformanceStats() {
    stats_ = {};
}

std::vector<TXRowGroup> TXInMemorySheet::generateRowGroups() const {
    return TXMemoryLayoutOptimizer::generateRowGroups(cell_buffer_);
}

// ==================== 内部辅助方法 ====================

void TXInMemorySheet::updateBounds(const TXCoordinate& coord) {
    max_row_ = std::max(max_row_, coord.getRow().index());
    max_col_ = std::max(max_col_, coord.getCol().index());
}

void TXInMemorySheet::updateIndex(const TXCoordinate& coord, size_t buffer_index) {
    uint32_t key = coordToKey(coord);
    coord_to_index_[key] = buffer_index;
}

void TXInMemorySheet::removeFromIndex(const TXCoordinate& coord) {
    uint32_t key = coordToKey(coord);
    coord_to_index_.erase(key);
}

void TXInMemorySheet::maybeOptimize() {
    if (!auto_optimize_) return;
    
    if (cell_buffer_.size >= OPTIMIZATION_THRESHOLD) {
        optimizeMemoryLayout();
    }
}

void TXInMemorySheet::updateStats(size_t cells_processed, double time_ms) const {
    stats_.total_cells += cells_processed;
    stats_.batch_operations++;
    stats_.total_operation_time += time_ms;
    stats_.cache_hits++; // 简化统计
}

uint32_t TXInMemorySheet::coordToKey(const TXCoordinate& coord) {
    return (coord.getRow().index() << 16) | coord.getCol().index();
}

TXCoordinate TXInMemorySheet::keyToCoord(uint32_t key) {
    uint32_t row = key >> 16;
    uint32_t col = key & 0xFFFF;
    return TXCoordinate(row_t(row), column_t(col));
}

// ==================== TXInMemoryWorkbook 实现 ====================

std::unique_ptr<TXInMemoryWorkbook> TXInMemoryWorkbook::create(const std::string& filename) {
    return std::make_unique<TXInMemoryWorkbook>(filename);
}

TXInMemoryWorkbook::TXInMemoryWorkbook(const std::string& filename)
    : filename_(filename), string_pool_(TXGlobalStringPool::instance()) {
    // 🚀 不再需要独立的memory_manager_，直接使用全局实例
}

TXInMemorySheet& TXInMemoryWorkbook::createSheet(const std::string& name) {
    // 🚀 使用全局内存管理器 - 这是关键！
    auto sheet = std::make_unique<TXInMemorySheet>(name, GlobalUnifiedMemoryManager::getInstance(), string_pool_);
    TXInMemorySheet& sheet_ref = *sheet;
    sheets_.push_back(std::move(sheet));
    return sheet_ref;
}

TXInMemorySheet& TXInMemoryWorkbook::getSheet(const std::string& name) {
    for (auto& sheet : sheets_) {
        if (sheet->getName() == name) {
            return *sheet;
        }
    }
    throw std::runtime_error("Sheet not found: " + name);
}

TXInMemorySheet& TXInMemoryWorkbook::getSheet(size_t index) {
    if (index >= sheets_.size()) {
        throw std::runtime_error("Sheet index out of range");
    }
    return *sheets_[index];
}

bool TXInMemoryWorkbook::removeSheet(const std::string& name) {
    auto it = std::find_if(sheets_.begin(), sheets_.end(),
        [&name](const std::unique_ptr<TXInMemorySheet>& sheet) {
            return sheet->getName() == name;
        });
    
    if (it != sheets_.end()) {
        sheets_.erase(it);
        return true;
    }
    return false;
}

TXResult<void> TXInMemoryWorkbook::saveToFile(const std::string& filename) {
    std::string output_filename = filename.empty() ? filename_ : filename;
    
    try {
        // 🚀 创建序列化器 - 使用全局内存管理器
        TXZeroCopySerializer serializer(GlobalUnifiedMemoryManager::getInstance());
        
        // 🚀 高性能序列化：预分配ZIP缓冲区
        TXZipArchiveWriter zip_writer;

        // 预估文件大小并预分配缓冲区
        size_t estimated_size = estimateFileSize();
        // 注意：TXZipArchiveWriter可能不支持setCompressionLevel，跳过此设置

        auto open_result = zip_writer.open(output_filename);
        if (open_result.isError()) {
            return TXResult<void>(open_result.error());
        }
        
        // 🚀 批量序列化优化：预分配所有数据
        std::vector<std::pair<std::string, std::vector<uint8_t>>> batch_data;
        batch_data.reserve(sheets_.size() + 10); // 预留额外空间

        for (size_t i = 0; i < sheets_.size(); ++i) {
            auto& sheet = *sheets_[i];

            // 🚀 使用独立的序列化器避免清理开销
            TXZeroCopySerializer sheet_serializer(GlobalUnifiedMemoryManager::getInstance());
            auto result = sheet.serializeToMemory(sheet_serializer);
            if (!result.isOk()) {
                return result;
            }

            // 🚀 批量收集数据，稍后一次性写入
            std::string sheet_filename = "xl/worksheets/sheet" + std::to_string(i + 1) + ".xml";
            batch_data.emplace_back(sheet_filename, std::move(sheet_serializer).getResult());
        }
        
        // 🚀 批量序列化其他文件

        // 序列化共享字符串
        if (string_pool_.size() > 0) {
            TXZeroCopySerializer shared_serializer(GlobalUnifiedMemoryManager::getInstance());
            auto shared_strings_result = shared_serializer.serializeSharedStrings(string_pool_);
            if (shared_strings_result.isOk()) {
                batch_data.emplace_back("xl/sharedStrings.xml", std::move(shared_serializer).getResult());
            }
        }

        // 序列化工作簿
        TXZeroCopySerializer workbook_serializer(GlobalUnifiedMemoryManager::getInstance());
        std::vector<std::string> sheet_names;
        sheet_names.reserve(sheets_.size());
        for (const auto& sheet : sheets_) {
            sheet_names.push_back(sheet->getName());
        }
        auto workbook_result = workbook_serializer.serializeWorkbook(sheet_names);
        if (workbook_result.isOk()) {
            batch_data.emplace_back("xl/workbook.xml", std::move(workbook_serializer).getResult());
        }

        // 🚀 批量添加XLSX结构文件到批量数据
        auto structure_result = addXLSXStructureFilesToBatch(batch_data, sheets_.size());
        if (!structure_result.isOk()) {
            return structure_result;
        }

        // 🚀 一次性批量写入所有文件到ZIP
        for (const auto& [filename, data] : batch_data) {
            auto write_result = zip_writer.write(filename, data);
            if (write_result.isError()) {
                return TXResult<void>(write_result.error());
            }
        }

        // 关闭ZIP文件
        zip_writer.close();
        
        // 标记所有工作表为已保存
        for (auto& sheet : sheets_) {
            sheet->markClean();
        }
        
        return TXResult<void>();
        
    } catch (const std::exception& e) {
        return TXResult<void>(TXError(TXErrorCode::OperationFailed, 
                                    std::string("保存文件失败: ") + e.what()));
    }
}

TXResult<std::vector<uint8_t>> TXInMemoryWorkbook::serializeToMemory() {
    try {
        // 使用临时文件进行序列化
        std::string temp_filename = std::tmpnam(nullptr);
        temp_filename += ".xlsx";

        // 先保存到临时文件
        auto save_result = saveToFile(temp_filename);
        if (!save_result.isOk()) {
            return TXResult<std::vector<uint8_t>>(save_result.error());
        }

        // 读取临时文件内容
        std::ifstream file(temp_filename, std::ios::binary);
        if (!file) {
            return TXResult<std::vector<uint8_t>>(TXError(TXErrorCode::OperationFailed, "无法读取临时文件"));
        }

        // 获取文件大小
        file.seekg(0, std::ios::end);
        size_t file_size = file.tellg();
        file.seekg(0, std::ios::beg);

        // 读取文件内容
        std::vector<uint8_t> excel_data(file_size);
        file.read(reinterpret_cast<char*>(excel_data.data()), file_size);
        file.close();

        // 删除临时文件
        std::remove(temp_filename.c_str());

        return TXResult<std::vector<uint8_t>>(std::move(excel_data));

    } catch (const std::exception& e) {
        return TXResult<std::vector<uint8_t>>(TXError(TXErrorCode::SerializationError,
                                                    std::string("内存序列化失败: ") + e.what()));
    }
}

TXResult<void> TXInMemoryWorkbook::addXLSXStructureFilesToBatch(
    std::vector<std::pair<std::string, std::vector<uint8_t>>>& batch_data, size_t sheet_count) {
    try {
        // 🚀 批量添加所有结构文件到batch_data

        // 1. [Content_Types].xml
        std::string content_types = generateContentTypesXML(sheet_count);
        batch_data.emplace_back("[Content_Types].xml",
                               std::vector<uint8_t>(content_types.begin(), content_types.end()));

        // 2. _rels/.rels
        std::string main_rels(TXCompiledXMLTemplates::MAIN_RELS);
        batch_data.emplace_back("_rels/.rels",
                               std::vector<uint8_t>(main_rels.begin(), main_rels.end()));

        // 3. xl/_rels/workbook.xml.rels
        std::string workbook_rels = generateWorkbookRelsXML(sheet_count);
        batch_data.emplace_back("xl/_rels/workbook.xml.rels",
                               std::vector<uint8_t>(workbook_rels.begin(), workbook_rels.end()));

        // 4. docProps/app.xml
        std::string app_props(TXCompiledXMLTemplates::APP_PROPERTIES);
        batch_data.emplace_back("docProps/app.xml",
                               std::vector<uint8_t>(app_props.begin(), app_props.end()));

        // 5. docProps/core.xml
        std::string timestamp = TXCompiledXMLTemplates::getCurrentTimestamp();
        std::string core_props = TXCompiledXMLTemplates::applyTemplate(
            TXCompiledXMLTemplates::CORE_PROPERTIES, timestamp, timestamp);
        batch_data.emplace_back("docProps/core.xml",
                               std::vector<uint8_t>(core_props.begin(), core_props.end()));

        return TXResult<void>();

    } catch (const std::exception& e) {
        return TXResult<void>(TXError(TXErrorCode::OperationFailed,
                                    std::string("添加XLSX结构文件失败: ") + e.what()));
    }
}

std::string TXInMemoryWorkbook::generateContentTypesXML(size_t sheet_count) {
    std::string content = std::string(TXCompiledXMLTemplates::CONTENT_TYPES_HEADER);

    // 添加工作表内容类型
    for (size_t i = 1; i <= sheet_count; ++i) {
        content += TXCompiledXMLTemplates::applyTemplate(
            TXCompiledXMLTemplates::WORKSHEET_CONTENT_TYPE, i);
    }

    // 添加共享字符串内容类型（如果有的话）
    if (string_pool_.size() > 0) {
        content += TXCompiledXMLTemplates::SHARED_STRINGS_CONTENT_TYPE;
    }

    content += TXCompiledXMLTemplates::CONTENT_TYPES_FOOTER;
    return content;
}

std::string TXInMemoryWorkbook::generateWorkbookRelsXML(size_t sheet_count) {
    std::string rels = std::string(TXCompiledXMLTemplates::WORKBOOK_RELS_HEADER);

    // 添加工作表关系
    for (size_t i = 1; i <= sheet_count; ++i) {
        rels += TXCompiledXMLTemplates::applyTemplate(
            TXCompiledXMLTemplates::WORKSHEET_REL, i, i);
    }

    // 添加共享字符串关系（如果有的话）
    if (string_pool_.size() > 0) {
        rels += TXCompiledXMLTemplates::applyTemplate(
            TXCompiledXMLTemplates::SHARED_STRINGS_REL, sheet_count + 1);
    }

    rels += TXCompiledXMLTemplates::WORKBOOK_RELS_FOOTER;
    return rels;
}

size_t TXInMemoryWorkbook::estimateFileSize() const {
    // 🚀 快速估算文件大小以优化内存分配
    size_t total_cells = 0;
    for (const auto& sheet : sheets_) {
        total_cells += sheet->getCellCount();
    }

    // 估算：每个单元格约50字节XML + 压缩率约30%
    size_t estimated_xml_size = total_cells * 50;
    size_t estimated_compressed_size = estimated_xml_size * 3 / 10;

    // 加上固定开销（结构文件等）
    return estimated_compressed_size + 10240; // 10KB固定开销
}

} // namespace TinaXlsx