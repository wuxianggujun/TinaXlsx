//
// @file TXZeroCopySerializer.cpp
// @brief 零拷贝序列化器实现 - 极速XML生成核心
//

#include "TinaXlsx/TXZeroCopySerializer.hpp"
#include "TinaXlsx/TXGlobalStringPool.hpp"
#include <fmt/format.h>
#include <algorithm>
#include <chrono>
#include <thread>
#include <sstream>

namespace TinaXlsx {

// 静态成员初始化
TXXMLTemplate TXZeroCopySerializer::worksheet_template_;
TXXMLTemplate TXZeroCopySerializer::shared_strings_template_;
TXXMLTemplate TXZeroCopySerializer::workbook_template_;
bool TXZeroCopySerializer::templates_initialized_ = false;

// ==================== TXZeroCopySerializer 实现 ====================

TXZeroCopySerializer::TXZeroCopySerializer(
    TXUnifiedMemoryManager& memory_manager,
    const TXSerializationOptions& options
) : memory_manager_(memory_manager), options_(options) {
    
    if (!templates_initialized_) {
        initializeTemplates();
    }
    
    // 预分配输出缓冲区
    output_buffer_.reserve(options_.buffer_size);
}

TXZeroCopySerializer::~TXZeroCopySerializer() = default;

TXZeroCopySerializer::TXZeroCopySerializer(TXZeroCopySerializer&& other) noexcept
    : memory_manager_(other.memory_manager_)
    , output_buffer_(std::move(other.output_buffer_))
    , current_pos_(other.current_pos_)
    , options_(other.options_)
    , stats_(other.stats_) {
    other.current_pos_ = 0;
}

TXZeroCopySerializer& TXZeroCopySerializer::operator=(TXZeroCopySerializer&& other) noexcept {
    if (this != &other) {
        output_buffer_ = std::move(other.output_buffer_);
        current_pos_ = other.current_pos_;
        options_ = other.options_;
        stats_ = other.stats_;
        other.current_pos_ = 0;
    }
    return *this;
}

// ==================== 核心序列化方法 ====================

TXResult<void> TXZeroCopySerializer::serializeWorksheet(const TXInMemorySheet& sheet) {
    auto start_time = std::chrono::high_resolution_clock::now();
    
    try {
        // 预估大小并预分配内存
        size_t estimated_size = estimateWorksheetSize(sheet);
        reserve(estimated_size);
        
        // 写入XML声明和工作表开始
        writeXMLDeclaration();
        writeWorksheetStart();
        writeSheetDataStart();
        
        // 获取单元格数据和行分组
        const auto& cell_buffer = sheet.getCellBuffer();
        auto row_groups = sheet.generateRowGroups();
        
        // 批量序列化单元格数据
        size_t serialized_cells = 0;
        if (options_.enable_parallel && cell_buffer.size >= options_.parallel_threshold) {
            auto result = serializeParallel(cell_buffer, row_groups);
            if (!result.isSuccess()) {
                return TXResult<void>::error(result.getError());
            }
            serialized_cells = cell_buffer.size;
        } else {
            serialized_cells = serializeCellDataBatch(cell_buffer, row_groups);
        }
        
        // 写入结束标签
        writeSheetDataEnd();
        writeWorksheetEnd();
        
        auto end_time = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
        updateStats(serialized_cells, current_pos_, duration.count() / 1000.0);
        
        return TXResult<void>::success();
        
    } catch (const std::exception& e) {
        return TXResult<void>::error(TXError::SerializationError, 
                                   fmt::format("工作表序列化失败: {}", e.what()));
    }
}

TXResult<void> TXZeroCopySerializer::serializeSharedStrings(const TXGlobalStringPool& string_pool) {
    try {
        // 写入共享字符串头部
        writeStringView(TXCompiledXMLTemplates::SHARED_STRINGS_HEADER);
        
        // 序列化所有字符串
        auto strings = string_pool.getAllStrings();
        for (const auto& str : strings) {
            std::string escaped = escapeXMLString(str);
            std::string item = TXCompiledXMLTemplates::applyTemplate(
                TXCompiledXMLTemplates::SHARED_STRING_ITEM, escaped);
            writeString(item);
        }
        
        // 写入尾部
        writeStringView(TXCompiledXMLTemplates::SHARED_STRINGS_FOOTER);
        
        return TXResult<void>::success();
        
    } catch (const std::exception& e) {
        return TXResult<void>::error(TXError::SerializationError, 
                                   fmt::format("共享字符串序列化失败: {}", e.what()));
    }
}

TXResult<void> TXZeroCopySerializer::serializeWorkbook(const std::vector<std::string>& sheet_names) {
    try {
        // 写入工作簿头部
        writeStringView(TXCompiledXMLTemplates::WORKBOOK_HEADER);
        
        // 序列化工作表条目
        for (size_t i = 0; i < sheet_names.size(); ++i) {
            std::string escaped_name = escapeXMLString(sheet_names[i]);
            std::string sheet_entry = TXCompiledXMLTemplates::applyTemplate(
                TXCompiledXMLTemplates::SHEET_ENTRY, 
                escaped_name, i + 1, i + 1);
            writeString(sheet_entry);
        }
        
        // 写入尾部
        writeStringView(TXCompiledXMLTemplates::WORKBOOK_FOOTER);
        
        return TXResult<void>::success();
        
    } catch (const std::exception& e) {
        return TXResult<void>::error(TXError::SerializationError, 
                                   fmt::format("工作簿序列化失败: {}", e.what()));
    }
}

// ==================== 批量序列化实现 ====================

size_t TXZeroCopySerializer::serializeCellDataBatch(
    const TXCompactCellBuffer& buffer,
    const std::vector<TXRowGroup>& row_groups
) {
    size_t total_serialized = 0;
    
    // 按行分组序列化
    for (const auto& row_group : row_groups) {
        size_t row_serialized = serializeRowBatch(buffer, row_group);
        total_serialized += row_serialized;
    }
    
    return total_serialized;
}

size_t TXZeroCopySerializer::serializeRowBatch(
    const TXCompactCellBuffer& buffer,
    const TXRowGroup& row_group
) {
    // 写入行开始标签
    writeRowStart(row_group.row_index + 1); // Excel行号从1开始
    
    size_t serialized_count = 0;
    
    // 序列化该行的所有单元格
    for (size_t i = row_group.start_cell_index; 
         i < row_group.start_cell_index + row_group.cell_count; ++i) {
        
        // 获取单元格坐标字符串
        std::string coord_str = coordToString(buffer.coordinates[i]);
        
        // 根据类型序列化单元格
        uint8_t cell_type = buffer.cell_types[i];
        
        if (cell_type == static_cast<uint8_t>(TXCellType::Number)) {
            writeNumberCell(coord_str, buffer.number_values[i]);
        } else if (cell_type == static_cast<uint8_t>(TXCellType::String)) {
            // 这里应该从字符串池获取实际字符串，简化版本直接使用索引
            writeInlineStringCell(coord_str, fmt::format("String_{}", buffer.string_indices[i]));
        }
        // 其他类型可以在这里添加
        
        ++serialized_count;
    }
    
    // 写入行结束标签
    writeRowEnd();
    
    return serialized_count;
}

TXResult<void> TXZeroCopySerializer::serializeParallel(
    const TXCompactCellBuffer& buffer,
    const std::vector<TXRowGroup>& row_groups
) {
    try {
        // 简单的并行实现：将行分组分割给不同线程
        const size_t num_threads = std::thread::hardware_concurrency();
        const size_t groups_per_thread = (row_groups.size() + num_threads - 1) / num_threads;
        
        std::vector<std::thread> threads;
        std::vector<std::string> thread_outputs(num_threads);
        
        for (size_t t = 0; t < num_threads; ++t) {
            size_t start_group = t * groups_per_thread;
            size_t end_group = std::min(start_group + groups_per_thread, row_groups.size());
            
            if (start_group >= row_groups.size()) break;
            
            threads.emplace_back([&, t, start_group, end_group]() {
                std::ostringstream oss;
                
                for (size_t g = start_group; g < end_group; ++g) {
                    const auto& row_group = row_groups[g];
                    
                    // 写入行开始标签
                    oss << fmt::format(TXCompiledXMLTemplates::ROW_START, row_group.row_index + 1);
                    
                    // 序列化该行的所有单元格
                    for (size_t i = row_group.start_cell_index; 
                         i < row_group.start_cell_index + row_group.cell_count; ++i) {
                        
                        std::string coord_str = coordToString(buffer.coordinates[i]);
                        uint8_t cell_type = buffer.cell_types[i];
                        
                        if (cell_type == static_cast<uint8_t>(TXCellType::Number)) {
                            oss << fmt::format(TXCompiledXMLTemplates::CELL_NUMBER, 
                                             coord_str, buffer.number_values[i]);
                        } else if (cell_type == static_cast<uint8_t>(TXCellType::String)) {
                            std::string text = fmt::format("String_{}", buffer.string_indices[i]);
                            oss << fmt::format(TXCompiledXMLTemplates::CELL_INLINE_STRING, 
                                             coord_str, escapeXMLString(text));
                        }
                    }
                    
                    // 写入行结束标签
                    oss << TXCompiledXMLTemplates::ROW_END;
                }
                
                thread_outputs[t] = oss.str();
            });
        }
        
        // 等待所有线程完成
        for (auto& thread : threads) {
            thread.join();
        }
        
        // 合并结果
        for (const auto& output : thread_outputs) {
            writeString(output);
        }
        
        return TXResult<void>::success();
        
    } catch (const std::exception& e) {
        return TXResult<void>::error(TXError::SerializationError, 
                                   fmt::format("并行序列化失败: {}", e.what()));
    }
}

// ==================== 高性能写入方法 ====================

void TXZeroCopySerializer::reserve(size_t estimated_size) {
    if (estimated_size > output_buffer_.capacity()) {
        output_buffer_.reserve(estimated_size);
    }
}

void TXZeroCopySerializer::writeRaw(const void* data, size_t size) {
    ensureCapacity(size);
    std::memcpy(output_buffer_.data() + current_pos_, data, size);
    current_pos_ += size;
}

void TXZeroCopySerializer::writeString(const std::string& str) {
    writeStringView(str);
}

void TXZeroCopySerializer::writeStringView(std::string_view str) {
    ensureCapacity(str.size());
    std::memcpy(output_buffer_.data() + current_pos_, str.data(), str.size());
    current_pos_ += str.size();
}

void TXZeroCopySerializer::writeStringsBatch(const std::vector<std::string>& strings) {
    // 预计算总大小
    size_t total_size = 0;
    for (const auto& str : strings) {
        total_size += str.size();
    }
    
    ensureCapacity(total_size);
    
    // 批量写入
    for (const auto& str : strings) {
        std::memcpy(output_buffer_.data() + current_pos_, str.data(), str.size());
        current_pos_ += str.size();
    }
}

template<typename... Args>
void TXZeroCopySerializer::applyTemplate(const std::string& template_str, Args&&... args) {
    std::string result = fmt::format(template_str, std::forward<Args>(args)...);
    writeString(result);
}

// ==================== 快速单元格写入 ====================

void TXZeroCopySerializer::writeNumberCell(const std::string& coord_str, double value) {
    std::string cell_xml = TXCompiledXMLTemplates::applyTemplate(
        TXCompiledXMLTemplates::CELL_NUMBER, coord_str, value);
    writeString(cell_xml);
}

void TXZeroCopySerializer::writeStringCell(const std::string& coord_str, const std::string& value) {
    std::string escaped_value = escapeXMLString(value);
    std::string cell_xml = TXCompiledXMLTemplates::applyTemplate(
        TXCompiledXMLTemplates::CELL_STRING, coord_str, escaped_value);
    writeString(cell_xml);
}

void TXZeroCopySerializer::writeInlineStringCell(const std::string& coord_str, const std::string& value) {
    std::string escaped_value = escapeXMLString(value);
    std::string cell_xml = TXCompiledXMLTemplates::applyTemplate(
        TXCompiledXMLTemplates::CELL_INLINE_STRING, coord_str, escaped_value);
    writeString(cell_xml);
}

void TXZeroCopySerializer::writeNumberCellsBatch(
    const std::vector<std::string>& coords,
    const double* values,
    size_t count
) {
    // 预计算总大小
    size_t estimated_size = count * 50; // 预估每个数值单元格50字节
    ensureCapacity(estimated_size);
    
    for (size_t i = 0; i < count; ++i) {
        writeNumberCell(coords[i], values[i]);
    }
}

// ==================== XML结构写入 ====================

void TXZeroCopySerializer::writeXMLDeclaration() {
    writeStringView("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
}

void TXZeroCopySerializer::writeWorksheetStart() {
    writeStringView("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
}

void TXZeroCopySerializer::writeWorksheetEnd() {
    writeStringView("</worksheet>");
}

void TXZeroCopySerializer::writeRowStart(uint32_t row_index) {
    std::string row_start = fmt::format(TXCompiledXMLTemplates::ROW_START, row_index);
    writeString(row_start);
}

void TXZeroCopySerializer::writeRowEnd() {
    writeStringView(TXCompiledXMLTemplates::ROW_END);
}

void TXZeroCopySerializer::writeSheetDataStart() {
    writeStringView("<sheetData>");
}

void TXZeroCopySerializer::writeSheetDataEnd() {
    writeStringView("</sheetData>");
}

// ==================== 性能优化 ====================

size_t TXZeroCopySerializer::estimateWorksheetSize(const TXInMemorySheet& sheet) {
    size_t cell_count = sheet.getCellCount();
    if (cell_count == 0) return 1024; // 最小大小
    
    // 预估：每个单元格平均50字节XML
    // 加上头部、尾部和行标签的开销
    size_t estimated_cells_size = cell_count * 50;
    size_t estimated_overhead = 1024 + (sheet.getMaxRow() + 1) * 20; // 行标签开销
    
    return estimated_cells_size + estimated_overhead;
}

size_t TXZeroCopySerializer::estimateCellsSize(size_t cell_count, size_t avg_string_length) {
    // 数值单元格: ~40字节
    // 字符串单元格: ~60字节 + 字符串长度
    size_t avg_cell_size = 40 + avg_string_length * 0.5; // 假设一半是字符串
    return cell_count * avg_cell_size;
}

void TXZeroCopySerializer::initializeTemplates() {
    if (templates_initialized_) return;
    
    // 工作表模板
    worksheet_template_.header = TXCompiledXMLTemplates::WORKSHEET_HEADER;
    worksheet_template_.footer = TXCompiledXMLTemplates::WORKSHEET_FOOTER;
    worksheet_template_.cell_number_template = TXCompiledXMLTemplates::CELL_NUMBER;
    worksheet_template_.cell_string_template = TXCompiledXMLTemplates::CELL_INLINE_STRING;
    worksheet_template_.estimated_size_per_cell = 50;
    worksheet_template_.is_compiled = true;
    
    // 共享字符串模板
    shared_strings_template_.header = TXCompiledXMLTemplates::SHARED_STRINGS_HEADER;
    shared_strings_template_.footer = TXCompiledXMLTemplates::SHARED_STRINGS_FOOTER;
    shared_strings_template_.is_compiled = true;
    
    // 工作簿模板
    workbook_template_.header = TXCompiledXMLTemplates::WORKBOOK_HEADER;
    workbook_template_.footer = TXCompiledXMLTemplates::WORKBOOK_FOOTER;
    workbook_template_.is_compiled = true;
    
    templates_initialized_ = true;
}

void TXZeroCopySerializer::optimizeBuffer() {
    // 确保缓冲区大小合理
    if (output_buffer_.capacity() > current_pos_ * 2) {
        output_buffer_.shrink_to_fit();
    }
}

double TXZeroCopySerializer::compressOutput() {
    if (!options_.enable_compression) return 1.0;
    
    // 简化版本：这里应该使用zlib或其他压缩算法
    // 暂时返回模拟的压缩比
    return 0.7; // 假设70%压缩比
}

// ==================== 结果获取 ====================

std::vector<uint8_t> TXZeroCopySerializer::getResult() && {
    output_buffer_.resize(current_pos_);
    return std::move(output_buffer_);
}

std::span<const uint8_t> TXZeroCopySerializer::getResultView() const {
    return std::span<const uint8_t>(output_buffer_.data(), current_pos_);
}

void TXZeroCopySerializer::clear() {
    current_pos_ = 0;
    stats_ = {};
}

// ==================== 性能监控 ====================

TXZeroCopySerializer::SerializationStats TXZeroCopySerializer::getPerformanceStats() const {
    SerializationStats stats;
    stats.total_cells = stats_.total_cells_serialized;
    stats.total_bytes = stats_.total_bytes_written;
    stats.serialization_time_ms = stats_.total_time_ms;
    
    if (stats_.total_time_ms > 0) {
        stats.throughput_cells_per_sec = stats_.total_cells_serialized / (stats_.total_time_ms / 1000.0);
        stats.throughput_mb_per_sec = (stats_.total_bytes_written / 1024.0 / 1024.0) / (stats_.total_time_ms / 1000.0);
    }
    
    stats.template_cache_hits = stats_.template_cache_hits;
    stats.compression_ratio = stats_.compression_ratio_percent / 100.0;
    stats.memory_usage_bytes = output_buffer_.capacity();
    
    return stats;
}

void TXZeroCopySerializer::resetPerformanceStats() {
    stats_ = {};
}

// ==================== 内部辅助方法 ====================

void TXZeroCopySerializer::ensureCapacity(size_t additional_size) {
    size_t required_size = current_pos_ + additional_size;
    if (required_size > output_buffer_.size()) {
        output_buffer_.resize(std::max(required_size, output_buffer_.capacity() * 2));
    }
}

void TXZeroCopySerializer::appendToOutput(const void* data, size_t size) {
    writeRaw(data, size);
}

void TXZeroCopySerializer::appendToOutput(std::string_view str) {
    writeStringView(str);
}

std::string TXZeroCopySerializer::coordToString(uint32_t coord) const {
    uint32_t row = coord >> 16;
    uint32_t col = coord & 0xFFFF;
    return rowColToString(row, col);
}

std::string TXZeroCopySerializer::rowColToString(uint32_t row, uint32_t col) const {
    // 将行列转换为Excel格式 (如: A1, B2, AA10)
    std::string result;
    
    // 列转换为字母
    uint32_t temp_col = col;
    do {
        result = char('A' + (temp_col % 26)) + result;
        temp_col /= 26;
    } while (temp_col > 0);
    
    // 行号 (1-based)
    result += std::to_string(row + 1);
    
    return result;
}

void TXZeroCopySerializer::numberToString(double value, char* buffer, size_t buffer_size) {
    snprintf(buffer, buffer_size, "%.15g", value);
}

void TXZeroCopySerializer::numberToStringBatch(const double* values, std::vector<std::string>& output, size_t count) {
    output.reserve(count);
    char buffer[32];
    
    for (size_t i = 0; i < count; ++i) {
        numberToString(values[i], buffer, sizeof(buffer));
        output.emplace_back(buffer);
    }
}

std::string TXZeroCopySerializer::escapeXMLString(const std::string& str) {
    std::string result;
    result.reserve(str.size() * 1.1); // 预留一些空间给转义字符
    
    for (char c : str) {
        switch (c) {
            case '<': result += "&lt;"; break;
            case '>': result += "&gt;"; break;
            case '&': result += "&amp;"; break;
            case '"': result += "&quot;"; break;
            case '\'': result += "&apos;"; break;
            default: result += c; break;
        }
    }
    
    return result;
}

void TXZeroCopySerializer::escapeXMLStringInPlace(std::string& str) {
    str = escapeXMLString(str);
}

void TXZeroCopySerializer::updateStats(size_t cells_processed, size_t bytes_written, double time_ms) {
    stats_.total_cells_serialized += cells_processed;
    stats_.total_bytes_written += bytes_written;
    stats_.total_time_ms += time_ms;
    stats_.template_cache_hits++;
}

// ==================== TXStreamingZipWriter 实现 ====================

TXStreamingZipWriter::TXStreamingZipWriter() {
    zip_buffer_.reserve(1024 * 1024); // 1MB初始容量
}

void TXStreamingZipWriter::addFile(const std::string& filename, std::vector<uint8_t> data) {
    ZipEntry entry;
    entry.filename = filename;
    entry.uncompressed_size = data.size();
    entry.crc32 = calculateCRC32(data.data(), data.size());
    
    if (data.size() > 1024) { // 只压缩大于1KB的文件
        entry.data = compressData(data.data(), data.size());
        entry.compressed_size = entry.data.size();
    } else {
        entry.data = std::move(data);
        entry.compressed_size = entry.uncompressed_size;
    }
    
    entries_.push_back(std::move(entry));
}

void TXStreamingZipWriter::addFile(const std::string& filename, std::span<const uint8_t> data) {
    std::vector<uint8_t> data_copy(data.begin(), data.end());
    addFile(filename, std::move(data_copy));
}

std::vector<uint8_t> TXStreamingZipWriter::generateZip() {
    std::vector<uint8_t> zip_data;
    
    // 简化版本：这里应该生成完整的ZIP文件格式
    // 包括文件头、中央目录等
    
    // 为每个文件写入数据
    for (const auto& entry : entries_) {
        // ZIP文件头 (简化版本)
        zip_data.insert(zip_data.end(), entry.data.begin(), entry.data.end());
    }
    
    return zip_data;
}

size_t TXStreamingZipWriter::getZipSize() const {
    size_t total_size = 0;
    for (const auto& entry : entries_) {
        total_size += entry.compressed_size + 30; // 30字节的文件头开销
    }
    return total_size + 1024; // 加上中央目录开销
}

uint32_t TXStreamingZipWriter::calculateCRC32(const uint8_t* data, size_t size) {
    // 简化版本：这里应该实现标准的CRC32算法
    uint32_t crc = 0xFFFFFFFF;
    for (size_t i = 0; i < size; ++i) {
        crc ^= data[i];
        for (int j = 0; j < 8; ++j) {
            crc = (crc >> 1) ^ (0xEDB88320 & (-(crc & 1)));
        }
    }
    return ~crc;
}

std::vector<uint8_t> TXStreamingZipWriter::compressData(const uint8_t* data, size_t size) {
    // 简化版本：这里应该使用zlib压缩
    // 暂时返回原始数据
    return std::vector<uint8_t>(data, data + size);
}

} // namespace TinaXlsx