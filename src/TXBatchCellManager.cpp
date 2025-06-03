//
// @file TXBatchCellManager.cpp
// @brief 批处理单元格管理器实现
//

#include "TinaXlsx/TXBatchCellManager.hpp"
#include "TinaXlsx/TXOptimizedSIMD.hpp"
#include <algorithm>
#include <cstring>

namespace TinaXlsx {

// ==================== TXMemoryChunk 实现 ====================

TXMemoryChunk::TXMemoryChunk() {
    // 预分配第一个块
    chunks_[0].data = std::make_unique<char[]>(CHUNK_SIZE);
    current_chunk_ = 0;
    total_allocated_ = CHUNK_SIZE;
}

TXMemoryChunk::~TXMemoryChunk() {
    clear();
}

void* TXMemoryChunk::allocate(size_t size) {
    if (size > CHUNK_SIZE) {
        return nullptr; // 单次分配不能超过块大小
    }
    
    std::lock_guard<std::mutex> lock(mutex_);
    
    size_t current = current_chunk_.load();
    
    // 检查当前块是否有足够空间
    if (chunks_[current].used + size <= CHUNK_SIZE) {
        void* ptr = chunks_[current].data.get() + chunks_[current].used;
        chunks_[current].used += size;
        return ptr;
    }
    
    // 需要新块
    if (current + 1 >= MAX_CHUNKS) {
        return nullptr; // 达到最大块数限制
    }
    
    // 检查内存限制
    if (!checkMemoryLimit(CHUNK_SIZE)) {
        return nullptr;
    }
    
    // 分配新块
    size_t new_chunk = current + 1;
    chunks_[new_chunk].data = std::make_unique<char[]>(CHUNK_SIZE);
    chunks_[new_chunk].used = size;
    
    current_chunk_ = new_chunk;
    total_allocated_ += CHUNK_SIZE;
    
    return chunks_[new_chunk].data.get();
}

void TXMemoryChunk::clear() {
    std::lock_guard<std::mutex> lock(mutex_);
    
    for (auto& chunk : chunks_) {
        chunk.data.reset();
        chunk.used = 0;
    }
    
    current_chunk_ = 0;
    total_allocated_ = 0;
    
    // 重新分配第一个块
    chunks_[0].data = std::make_unique<char[]>(CHUNK_SIZE);
    total_allocated_ = CHUNK_SIZE;
}

bool TXMemoryChunk::checkMemoryLimit(size_t requested_size) const {
    return total_allocated_.load() + requested_size <= MAX_MEMORY;
}

// ==================== TXStringBuffer 实现 ====================

TXStringBuffer::TXStringBuffer() {
    buffer_.reserve(16 * 1024 * 1024); // 预分配16MB
    
    // 添加空字符串在偏移量0
    buffer_.push_back('\0');
    offset_map_[""] = 0;
}

uint32_t TXStringBuffer::addString(const std::string& str) {
    if (str.empty()) {
        return 0; // 空字符串总是偏移量0
    }
    
    std::lock_guard<std::mutex> lock(mutex_);
    
    // 检查是否已存在
    auto it = offset_map_.find(str);
    if (it != offset_map_.end()) {
        return it->second;
    }
    
    // 添加新字符串
    uint32_t offset = static_cast<uint32_t>(buffer_.size());
    buffer_.insert(buffer_.end(), str.begin(), str.end());
    buffer_.push_back('\0'); // 添加终止符
    
    offset_map_[str] = offset;
    return offset;
}

std::string_view TXStringBuffer::getString(uint32_t offset) const {
    std::lock_guard<std::mutex> lock(mutex_);
    
    if (offset >= buffer_.size()) {
        return "";
    }
    
    const char* start = buffer_.data() + offset;
    size_t len = strlen(start);
    return std::string_view(start, len);
}

void TXStringBuffer::reserve(size_t size) {
    std::lock_guard<std::mutex> lock(mutex_);
    buffer_.reserve(size);
}

void TXStringBuffer::clear() {
    std::lock_guard<std::mutex> lock(mutex_);
    buffer_.clear();
    offset_map_.clear();
    
    // 重新添加空字符串
    buffer_.push_back('\0');
    offset_map_[""] = 0;
}

// ==================== TXBatchCellManager 实现 ====================

TXBatchCellManager::TXBatchCellManager() 
    : memory_chunk_(std::make_unique<TXMemoryChunk>())
    , string_buffer_(std::make_unique<TXStringBuffer>()) {
    
    cells_.reserve(100000); // 预分配10万个单元格
    coordinate_index_.reserve(100000);
    
    resetStats();
}

TXBatchCellManager::~TXBatchCellManager() = default;

size_t TXBatchCellManager::setBatchCells(const std::vector<CellData>& cells) {
    if (cells.empty()) {
        return 0;
    }
    
    auto start_time = std::chrono::steady_clock::now();
    
    std::lock_guard<std::mutex> lock(data_mutex_);
    
    size_t processed = 0;
    
    // 预估内存需求
    size_t estimated_memory = cells.size() * sizeof(UltraCompactCell);
    if (!checkMemoryLimit(estimated_memory)) {
        return 0; // 内存不足
    }
    
    // 批量处理
    for (const auto& cell_data : cells) {
        try {
            setCell(cell_data);
            ++processed;
        } catch (const std::exception&) {
            // 处理失败，跳过这个单元格
            continue;
        }
    }
    
    auto end_time = std::chrono::steady_clock::now();
    updateStats(processed, start_time, end_time);
    
    return processed;
}

std::vector<CellData> TXBatchCellManager::getBatchCells(const CellRange& range) const {
    std::lock_guard<std::mutex> lock(data_mutex_);
    
    std::vector<CellData> result;
    result.reserve(range.getCellCount());
    
    for (uint16_t row = range.start_row; row <= range.end_row; ++row) {
        for (uint16_t col = range.start_col; col <= range.end_col; ++col) {
            row_t row_obj{static_cast<u32>(row)};
            column_t col_obj{static_cast<u32>(col)};
            TXCoordinate coord{row_obj, col_obj};
            size_t index = findCellIndex(coord);

            if (index < cells_.size()) {
                result.push_back(decodeCellData(cells_[index]));
            } else {
                // 空单元格
                CellData empty_cell;
                empty_cell.value = std::monostate{};
                empty_cell.coordinate = coord;
                result.push_back(empty_cell);
            }
        }
    }
    
    return result;
}

CellData TXBatchCellManager::getCell(const TXCoordinate& coord) const {
    std::lock_guard<std::mutex> lock(data_mutex_);
    
    size_t index = findCellIndex(coord);
    if (index < cells_.size()) {
        return decodeCellData(cells_[index]);
    }
    
    return CellData(std::monostate{}, coord);
}

void TXBatchCellManager::setCell(const CellData& data) {
    uint32_t key = coordinateToKey(data.coordinate);
    auto it = coordinate_index_.find(key);
    
    if (it != coordinate_index_.end()) {
        // 更新现有单元格
        updateExistingCell(it->second, data);
    } else {
        // 添加新单元格
        size_t index = addNewCell(data);
        coordinate_index_[key] = index;
    }
}

void TXBatchCellManager::compactMemory() {
    std::lock_guard<std::mutex> lock(data_mutex_);
    
    // 移除空单元格
    auto new_end = std::remove_if(cells_.begin(), cells_.end(),
        [](const UltraCompactCell& cell) {
            return cell.getType() == UltraCompactCell::CellType::Empty;
        });
    
    cells_.erase(new_end, cells_.end());
    
    // 重建索引
    coordinate_index_.clear();
    for (size_t i = 0; i < cells_.size(); ++i) {
        uint32_t key = coordinateToKey(cells_[i].getCoordinate());
        coordinate_index_[key] = i;
    }
    
    // 压缩字符串缓冲区（简化实现）
    string_buffer_->clear();
    
    // 重新添加所有字符串
    for (auto& cell : cells_) {
        if (cell.getType() == UltraCompactCell::CellType::String ||
            cell.getType() == UltraCompactCell::CellType::Formula) {
            // 这里需要重新构建字符串映射
            // 简化实现：暂时跳过
        }
    }
}

void TXBatchCellManager::clear() {
    std::lock_guard<std::mutex> lock(data_mutex_);
    
    cells_.clear();
    coordinate_index_.clear();
    string_buffer_->clear();
    memory_chunk_->clear();
    
    resetStats();
}

size_t TXBatchCellManager::getMemoryUsage() const {
    size_t total = 0;

    // 单元格内存（实际使用的内存）
    total += cells_.size() * sizeof(UltraCompactCell);

    // 索引内存（实际使用的内存）
    total += coordinate_index_.size() * (sizeof(uint32_t) + sizeof(size_t));

    // 字符串缓冲区（实际使用的内存）
    total += string_buffer_->getSize();

    // 不计算预分配的内存块，只计算实际使用的内存
    // total += memory_chunk_->getCurrentUsage();

    return total;
}

bool TXBatchCellManager::checkMemoryLimit(size_t additional_size) const {
    return memory_chunk_->checkMemoryLimit(additional_size);
}

TXBatchCellManager::BatchStats TXBatchCellManager::getStats() const {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    return stats_;
}

void TXBatchCellManager::resetStats() {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    stats_ = BatchStats{};
    stats_.start_time = std::chrono::steady_clock::now();
}

void TXBatchCellManager::startTiming() {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    stats_.start_time = std::chrono::steady_clock::now();
}

void TXBatchCellManager::endTiming() {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    stats_.end_time = std::chrono::steady_clock::now();
}

// ==================== 内部辅助方法实现 ====================

uint32_t TXBatchCellManager::coordinateToKey(const TXCoordinate& coord) const {
    return (static_cast<uint32_t>(coord.getRow().index()) << 16) | coord.getCol().index();
}

TXCoordinate TXBatchCellManager::keyToCoordinate(uint32_t key) const {
    uint16_t row = static_cast<uint16_t>(key >> 16);
    uint16_t col = static_cast<uint16_t>(key & 0xFFFF);
    row_t row_obj{static_cast<u32>(row)};
    column_t col_obj{static_cast<u32>(col)};
    return TXCoordinate{row_obj, col_obj};
}

size_t TXBatchCellManager::findCellIndex(const TXCoordinate& coord) const {
    uint32_t key = coordinateToKey(coord);
    auto it = coordinate_index_.find(key);
    return (it != coordinate_index_.end()) ? it->second : SIZE_MAX;
}

size_t TXBatchCellManager::addNewCell(const CellData& data) {
    UltraCompactCell cell = encodeCellData(data);
    cells_.push_back(cell);
    return cells_.size() - 1;
}

void TXBatchCellManager::updateExistingCell(size_t index, const CellData& data) {
    if (index < cells_.size()) {
        cells_[index] = encodeCellData(data);
    }
}

UltraCompactCell TXBatchCellManager::encodeCellData(const CellData& data) {
    UltraCompactCell cell;
    
    // 设置坐标
    cell.setCoordinate(data.coordinate);
    
    // 设置样式
    cell.setStyleIndex(data.style_index);
    
    // 设置公式状态
    cell.setIsFormula(data.is_formula);
    
    // 根据值类型编码
    std::visit([&cell, &data, this](const auto& value) {
        using T = std::decay_t<decltype(value)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            cell.setType(UltraCompactCell::CellType::Empty);
        } else if constexpr (std::is_same_v<T, std::string>) {
            uint32_t offset = string_buffer_->addString(value);
            cell = UltraCompactCell(value, offset);
            cell.setCoordinate(data.coordinate);
            cell.setStyleIndex(data.style_index);
        } else if constexpr (std::is_same_v<T, double>) {
            cell = UltraCompactCell(value);
            cell.setCoordinate(data.coordinate);
            cell.setStyleIndex(data.style_index);
        } else if constexpr (std::is_same_v<T, int64_t>) {
            cell = UltraCompactCell(value);
            cell.setCoordinate(data.coordinate);
            cell.setStyleIndex(data.style_index);
        } else if constexpr (std::is_same_v<T, bool>) {
            cell = UltraCompactCell(value);
            cell.setCoordinate(data.coordinate);
            cell.setStyleIndex(data.style_index);
        }
    }, data.value);
    
    return cell;
}

CellData TXBatchCellManager::decodeCellData(const UltraCompactCell& cell) const {
    CellData data;
    data.coordinate = cell.getCoordinate();
    data.style_index = cell.getStyleIndex();
    data.is_formula = cell.isFormula();
    
    switch (cell.getType()) {
        case UltraCompactCell::CellType::Empty:
            data.value = std::monostate{};
            break;
        case UltraCompactCell::CellType::String:
        case UltraCompactCell::CellType::Formula: {
            uint32_t offset = cell.getStringOffset();
            std::string_view str_view = string_buffer_->getString(offset);
            data.value = std::string(str_view);
            break;
        }
        case UltraCompactCell::CellType::Number:
            data.value = cell.getNumberValue();
            break;
        case UltraCompactCell::CellType::Integer:
            data.value = cell.getIntegerValue();
            break;
        case UltraCompactCell::CellType::Boolean:
            data.value = cell.getBooleanValue();
            break;
        default:
            data.value = std::monostate{};
            break;
    }
    
    return data;
}

void TXBatchCellManager::updateStats(size_t cells_count, 
                                    std::chrono::steady_clock::time_point start,
                                    std::chrono::steady_clock::time_point end) {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    double time_per_cell = static_cast<double>(duration.count()) / cells_count;
    
    stats_.cells_processed += cells_count;
    stats_.avg_time_per_cell = (stats_.avg_time_per_cell * (stats_.cells_processed - cells_count) + 
                               time_per_cell * cells_count) / stats_.cells_processed;
    stats_.memory_used = getMemoryUsage();
    stats_.memory_efficiency = static_cast<double>(cells_.size() * 16) / stats_.memory_used;
    stats_.string_pool_size = string_buffer_->getSize();
    stats_.end_time = end;
}

} // namespace TinaXlsx
