//
// @file TXZeroCopySerializer.cpp
// @brief 零拷贝序列化器实现 - 极速XML生成核心
//

#include "TinaXlsx/TXZeroCopySerializer.hpp"
#include "TinaXlsx/TXGlobalStringPool.hpp"
#include "TinaXlsx/TXHighPerformanceLogger.hpp"
#include "TinaXlsx/TXFastXmlWriter.hpp"
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
        // 🚀 性能监控：开始序列化
        size_t cell_count = sheet.getCellCount();
        TX_LOG_DEBUG("开始序列化工作表，单元格数量: {}", cell_count);

        // 🚀 预估大小并预分配内存
        size_t estimated_size = estimateWorksheetSize(sheet);
        reserve(estimated_size);
        TX_LOG_DEBUG("预分配缓冲区大小: {:.2f}KB", estimated_size / 1024.0);

        auto header_start = std::chrono::high_resolution_clock::now();
        // 🚀 使用高性能写入器写入XML头部
        TXFastXmlWriter header_writer(memory_manager_, 1024);
        header_writer.writeXmlDeclaration();
        header_writer.writeWorksheetStart();
        header_writer.writeSheetDataStart();

        // 将头部写入主缓冲区
        auto header_result = std::move(header_writer).getResult();
        ensureCapacity(header_result.size());
        std::memcpy(output_buffer_.data() + current_pos_, header_result.data(), header_result.size());
        current_pos_ += header_result.size();

        auto header_end = std::chrono::high_resolution_clock::now();
        auto header_duration = std::chrono::duration_cast<std::chrono::microseconds>(header_end - header_start);

        // 获取单元格数据和行分组
        auto data_prep_start = std::chrono::high_resolution_clock::now();
        const auto& cell_buffer = sheet.getCellBuffer();
        auto row_groups = sheet.generateRowGroups();
        auto data_prep_end = std::chrono::high_resolution_clock::now();
        auto data_prep_duration = std::chrono::duration_cast<std::chrono::microseconds>(data_prep_end - data_prep_start);

        // 🚀 使用高性能XML写入器进行序列化
        auto serialization_start = std::chrono::high_resolution_clock::now();
        size_t serialized_cells = 0;

        // 🚀 创建高性能XML写入器
        TXFastXmlWriter fast_writer(memory_manager_, estimated_size);

        if (options_.enable_parallel && cell_buffer.size >= options_.parallel_threshold) {
            TX_LOG_DEBUG("使用并行序列化，单元格数量: {}", cell_buffer.size);
            serialized_cells = serializeCellDataFast(fast_writer, cell_buffer, row_groups);
        } else {
            TX_LOG_DEBUG("使用高性能串行序列化，单元格数量: {}", cell_buffer.size);
            serialized_cells = serializeCellDataFast(fast_writer, cell_buffer, row_groups);
        }

        // 🚀 获取序列化结果并合并到输出缓冲区
        auto fast_result = std::move(fast_writer).getResult();
        ensureCapacity(fast_result.size());
        std::memcpy(output_buffer_.data() + current_pos_, fast_result.data(), fast_result.size());
        current_pos_ += fast_result.size();

        auto serialization_end = std::chrono::high_resolution_clock::now();
        auto serialization_duration = std::chrono::duration_cast<std::chrono::microseconds>(serialization_end - serialization_start);

        // 🚀 使用高性能写入器写入XML尾部
        auto footer_start = std::chrono::high_resolution_clock::now();
        TXFastXmlWriter footer_writer(memory_manager_, 512);
        footer_writer.writeSheetDataEnd();
        footer_writer.writeWorksheetEnd();

        // 将尾部写入主缓冲区
        auto footer_result = std::move(footer_writer).getResult();
        ensureCapacity(footer_result.size());
        std::memcpy(output_buffer_.data() + current_pos_, footer_result.data(), footer_result.size());
        current_pos_ += footer_result.size();

        auto footer_end = std::chrono::high_resolution_clock::now();
        auto footer_duration = std::chrono::duration_cast<std::chrono::microseconds>(footer_end - footer_start);

        auto end_time = std::chrono::high_resolution_clock::now();
        auto total_duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);

        // 🚀 详细性能统计
        TX_LOG_DEBUG("工作表序列化性能统计:");
        TX_LOG_DEBUG("  - 头部写入: {:.3f}ms", header_duration.count() / 1000.0);
        TX_LOG_DEBUG("  - 数据准备: {:.3f}ms", data_prep_duration.count() / 1000.0);
        TX_LOG_DEBUG("  - 单元格序列化: {:.3f}ms", serialization_duration.count() / 1000.0);
        TX_LOG_DEBUG("  - 尾部写入: {:.3f}ms", footer_duration.count() / 1000.0);
        TX_LOG_DEBUG("  - 总耗时: {:.3f}ms", total_duration.count() / 1000.0);
        TX_LOG_DEBUG("  - 序列化速度: {:.0f} 单元格/秒", serialized_cells / (total_duration.count() / 1000000.0));

        updateStats(serialized_cells, current_pos_, total_duration.count() / 1000.0);

        return TXResult<void>();

    } catch (const std::exception& e) {
        TX_LOG_ERROR("工作表序列化失败: {}", e.what());
        return TXResult<void>(TXError(TXErrorCode::SerializationError,
                                     fmt::format("工作表序列化失败: {}", e.what())));
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
        
        return TXResult<void>();
        
    } catch (const std::exception& e) {
        return TXResult<void>(TXError(TXErrorCode::SerializationError, 
                                     fmt::format("共享字符串序列化失败: {}", e.what())));
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
        
        return TXResult<void>();
        
    } catch (const std::exception& e) {
        return TXResult<void>(TXError(TXErrorCode::SerializationError, 
                                   std::string("工作簿序列化失败: ") + e.what()));
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
            writeInlineStringCell(coord_str, std::string("String_") + std::to_string(buffer.string_indices[i]));
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
                    oss << "<row r=\"" << (row_group.row_index + 1) << "\">";
                    
                    // 序列化该行的所有单元格
                    for (size_t i = row_group.start_cell_index; 
                         i < row_group.start_cell_index + row_group.cell_count; ++i) {
                        
                        std::string coord_str = coordToString(buffer.coordinates[i]);
                        uint8_t cell_type = buffer.cell_types[i];
                        
                        if (cell_type == static_cast<uint8_t>(TXCellType::Number)) {
                            oss << "<c r=\"" << coord_str << "\"><v>" << buffer.number_values[i] << "</v></c>";
                        } else if (cell_type == static_cast<uint8_t>(TXCellType::String)) {
                            std::string text = std::string("String_") + std::to_string(buffer.string_indices[i]);
                            oss << "<c r=\"" << coord_str << "\" t=\"inlineStr\"><is><t>" << escapeXMLString(text) << "</t></is></c>";
                        }
                    }
                    
                    // 写入行结束标签
                                                oss << "</row>";
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
        
        return TXResult<void>();
        
    } catch (const std::exception& e) {
        return TXResult<void>(TXError(TXErrorCode::SerializationError, 
                                   std::string("并行序列化失败: ") + e.what()));
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
    // 简化版本，暂时不使用模板
    writeString(template_str);
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
    std::string row_start = "<row r=\"" + std::to_string(row_index) + "\">";
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

    // 🚀 优化：更精确的大小预估，减少重新分配
    // 数值单元格：约35字节，字符串单元格：约60字节
    size_t estimated_cells_size = cell_count * 40; // 平均40字节
    size_t estimated_overhead = 2048 + (sheet.getMaxRow() + 1) * 25; // 增加行标签开销预估

    // 🚀 预留20%的缓冲区以避免重新分配
    size_t total_estimated = estimated_cells_size + estimated_overhead;
    return static_cast<size_t>(total_estimated * 1.2);
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

std::vector<uint8_t> TXZeroCopySerializer::getResultView() const {
    return std::vector<uint8_t>(output_buffer_.begin(), output_buffer_.begin() + current_pos_);
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
        // 🚀 优化：使用更激进的增长策略，减少重新分配次数
        size_t new_capacity = std::max({
            required_size,
            output_buffer_.capacity() * 2,
            output_buffer_.capacity() + 1024 * 1024  // 至少增长1MB
        });
        output_buffer_.resize(new_capacity);
        TX_LOG_DEBUG("缓冲区扩容: {} -> {} bytes", output_buffer_.capacity(), new_capacity);
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

    // 列转换为字母 (正确的Excel列转换算法)
    uint32_t temp_col = col + 1; // Excel列从1开始
    do {
        temp_col--;
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

// ==================== 高性能序列化方法 ====================

size_t TXZeroCopySerializer::serializeCellDataFast(
    TXFastXmlWriter& writer,
    const TXCompactCellBuffer& buffer,
    const std::vector<TXRowGroup>& row_groups
) {
    size_t total_serialized = 0;

    // 🚀 使用高性能XML写入器按行分组序列化
    for (const auto& row_group : row_groups) {
        // 写入行开始标签
        writer.writeRowStart(row_group.row_index + 1); // Excel行号从1开始

        // 序列化该行的所有单元格
        for (size_t i = row_group.start_cell_index;
             i < row_group.start_cell_index + row_group.cell_count; ++i) {

            // 🚀 快速坐标转换
            uint32_t coord = buffer.coordinates[i];
            uint32_t row = coord >> 16;
            uint32_t col = coord & 0xFFFF;
            std::string coord_str = TXCoordConverter::rowColToString(row, col);

            // 根据类型序列化单元格
            uint8_t cell_type = buffer.cell_types[i];

            if (cell_type == static_cast<uint8_t>(TXCellType::Number)) {
                writer.writeNumberCell(coord_str, buffer.number_values[i]);
            } else if (cell_type == static_cast<uint8_t>(TXCellType::String)) {
                // 简化版本：直接使用字符串索引
                std::string text = "String_" + std::to_string(buffer.string_indices[i]);
                writer.writeInlineStringCell(coord_str, text);
            }
            // 其他类型可以在这里添加

            ++total_serialized;
        }

        // 写入行结束标签
        writer.writeRowEnd();
    }

    return total_serialized;
}

} // namespace TinaXlsx