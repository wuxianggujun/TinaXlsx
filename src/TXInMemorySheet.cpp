//
// @file TXInMemorySheet.cpp
// @brief 内存优先工作表实现 - 完全内存中操作，极致性能
//

#include "TinaXlsx/TXInMemorySheet.hpp"
#include "TinaXlsx/TXZeroCopySerializer.hpp"
#include "TinaXlsx/TXBatchSIMDProcessor.hpp"
#include "TinaXlsx/TXGlobalStringPool.hpp"
#include <algorithm>
#include <chrono>
#include <sstream>
#include <fstream>
#include <map>

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
) : memory_manager_(memory_manager)
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
        // 转换坐标格式
        std::vector<uint32_t> packed_coords(coords.size());
        for (size_t i = 0; i < coords.size(); ++i) {
            packed_coords[i] = coordToKey(coords[i]);
            updateBounds(coords[i]);
        }
        
        // 使用SIMD批量处理
        size_t old_size = cell_buffer_.size;
        TXBatchSIMDProcessor::batchCreateNumberCells(
            values.data(), cell_buffer_, packed_coords.data(), coords.size());
        
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
}

TXInMemorySheet& TXInMemoryWorkbook::createSheet(const std::string& name) {
    auto sheet = std::make_unique<TXInMemorySheet>(name, memory_manager_, string_pool_);
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
        // 创建序列化器
        TXZeroCopySerializer serializer(memory_manager_);
        
        // 序列化所有工作表
        TXStreamingZipWriter zip_writer;
        
        for (size_t i = 0; i < sheets_.size(); ++i) {
            auto& sheet = *sheets_[i];
            
            // 序列化工作表
            serializer.clear();
            auto result = sheet.serializeToMemory(serializer);
            if (!result.isOk()) {
                return result;
            }
            
            // 添加到ZIP
            std::string sheet_filename = "xl/worksheets/sheet" + std::to_string(i + 1) + ".xml";
            auto serialized_data = std::move(serializer).getResult();
            zip_writer.addFile(sheet_filename, static_cast<std::vector<uint8_t>>(std::move(serialized_data)));
        }
        
        // 序列化共享字符串
        serializer.clear();
        auto shared_strings_result = serializer.serializeSharedStrings(string_pool_);
        if (shared_strings_result.isOk()) {
            auto shared_strings_data = std::move(serializer).getResult();
            zip_writer.addFile("xl/sharedStrings.xml", static_cast<std::vector<uint8_t>>(std::move(shared_strings_data)));
        }
        
        // 序列化工作簿
        serializer.clear();
        std::vector<std::string> sheet_names;
        for (const auto& sheet : sheets_) {
            sheet_names.push_back(sheet->getName());
        }
        auto workbook_result = serializer.serializeWorkbook(sheet_names);
        if (workbook_result.isOk()) {
            auto workbook_data = std::move(serializer).getResult();
            zip_writer.addFile("xl/workbook.xml", static_cast<std::vector<uint8_t>>(std::move(workbook_data)));
        }
        
        // 生成最终Excel文件
        auto excel_data = zip_writer.generateZip();
        
        // 写入文件
        std::ofstream file(output_filename, std::ios::binary);
        if (!file) {
            return TXResult<void>(TXError(TXErrorCode::OperationFailed, "无法创建输出文件"));
        }
        
        file.write(reinterpret_cast<const char*>(excel_data.data()), excel_data.size());
        
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
        TXZeroCopySerializer serializer(memory_manager_);
        TXStreamingZipWriter zip_writer;
        
        // 序列化所有组件（同saveToFile逻辑）
        for (size_t i = 0; i < sheets_.size(); ++i) {
            serializer.clear();
            auto result = sheets_[i]->serializeToMemory(serializer);
            if (!result.isOk()) {
                return TXResult<std::vector<uint8_t>>(result.error());
            }
            
            std::string filename = "xl/worksheets/sheet" + std::to_string(i + 1) + ".xml";
            auto data = std::move(serializer).getResult();
            zip_writer.addFile(filename, static_cast<std::vector<uint8_t>>(std::move(data)));
        }
        
        auto excel_data = zip_writer.generateZip();
        return TXResult<std::vector<uint8_t>>(std::move(excel_data));
        
    } catch (const std::exception& e) {
        return TXResult<std::vector<uint8_t>>(TXError(TXErrorCode::SerializationError, 
                                                    std::string("内存序列化失败: ") + e.what()));
    }
}

} // namespace TinaXlsx