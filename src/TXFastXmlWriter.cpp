//
// @file TXFastXmlWriter.cpp
// @brief 超高性能XML写入器实现
//

#include "TinaXlsx/TXFastXmlWriter.hpp"
#include <algorithm>
#include <cstdio>
#include <cstring>

namespace TinaXlsx {

// 线程局部存储
thread_local char TXFastXmlWriter::coord_buffer_[TXFastXmlWriter::MAX_COORD_LENGTH];
thread_local char TXFastXmlWriter::number_buffer_[TXFastXmlWriter::MAX_NUMBER_LENGTH];

// ==================== TXFastXmlWriter 实现 ====================

TXFastXmlWriter::TXFastXmlWriter(TXUnifiedMemoryManager& memory_manager, size_t initial_capacity)
    : memory_manager_(memory_manager), current_pos_(0) {
    buffer_.reserve(initial_capacity);
    buffer_.resize(initial_capacity);
    TX_LOG_DEBUG("TXFastXmlWriter 初始化，初始容量: {:.2f}KB", initial_capacity / 1024.0);
}

void TXFastXmlWriter::reserve(size_t capacity) {
    if (capacity > buffer_.capacity()) {
        buffer_.reserve(capacity);
        buffer_.resize(capacity);
        TX_LOG_DEBUG("TXFastXmlWriter 扩容至: {:.2f}KB", capacity / 1024.0);
    }
}

void TXFastXmlWriter::ensureCapacity(size_t additional_size) {
    size_t required_size = current_pos_ + additional_size;
    if (required_size > buffer_.size()) {
        size_t new_capacity = std::max({
            required_size,
            buffer_.capacity() * 2,
            buffer_.capacity() + 512 * 1024  // 至少增长512KB
        });
        buffer_.resize(new_capacity);
    }
}

void TXFastXmlWriter::writeRaw(const void* data, size_t size) {
    ensureCapacity(size);
    std::memcpy(buffer_.data() + current_pos_, data, size);
    current_pos_ += size;
}

void TXFastXmlWriter::writeConstant(std::string_view str) {
    ensureCapacity(str.size());
    std::memcpy(buffer_.data() + current_pos_, str.data(), str.size());
    current_pos_ += str.size();
}

void TXFastXmlWriter::writeRowStart(uint32_t row_number) {
    // 🚀 优化：直接写入，避免字符串拼接
    constexpr std::string_view prefix = "<row r=\"";
    constexpr std::string_view suffix = "\">";
    
    ensureCapacity(prefix.size() + 16 + suffix.size()); // 预留足够空间
    
    // 写入前缀
    std::memcpy(buffer_.data() + current_pos_, prefix.data(), prefix.size());
    current_pos_ += prefix.size();
    
    // 写入行号
    size_t number_len = TXFastNumberConverter::uint32ToString(row_number, number_buffer_, MAX_NUMBER_LENGTH);
    std::memcpy(buffer_.data() + current_pos_, number_buffer_, number_len);
    current_pos_ += number_len;
    
    // 写入后缀
    std::memcpy(buffer_.data() + current_pos_, suffix.data(), suffix.size());
    current_pos_ += suffix.size();
}

void TXFastXmlWriter::writeNumberCell(std::string_view coord, double value) {
    // 🚀 优化：预编译模板 <c r="COORD"><v>VALUE</v></c>
    constexpr std::string_view prefix = "<c r=\"";
    constexpr std::string_view middle = "\"><v>";
    constexpr std::string_view suffix = "</v></c>";
    
    size_t total_size = prefix.size() + coord.size() + middle.size() + MAX_NUMBER_LENGTH + suffix.size();
    ensureCapacity(total_size);
    
    // 写入前缀
    std::memcpy(buffer_.data() + current_pos_, prefix.data(), prefix.size());
    current_pos_ += prefix.size();
    
    // 写入坐标
    std::memcpy(buffer_.data() + current_pos_, coord.data(), coord.size());
    current_pos_ += coord.size();
    
    // 写入中间部分
    std::memcpy(buffer_.data() + current_pos_, middle.data(), middle.size());
    current_pos_ += middle.size();
    
    // 写入数值
    size_t number_len = TXFastNumberConverter::doubleToString(value, number_buffer_, MAX_NUMBER_LENGTH);
    std::memcpy(buffer_.data() + current_pos_, number_buffer_, number_len);
    current_pos_ += number_len;
    
    // 写入后缀
    std::memcpy(buffer_.data() + current_pos_, suffix.data(), suffix.size());
    current_pos_ += suffix.size();
}

void TXFastXmlWriter::writeInlineStringCell(std::string_view coord, std::string_view value) {
    // 🚀 优化：预编译模板 <c r="COORD" t="inlineStr"><is><t>VALUE</t></is></c>
    constexpr std::string_view prefix = "<c r=\"";
    constexpr std::string_view middle = "\" t=\"inlineStr\"><is><t>";
    constexpr std::string_view suffix = "</t></is></c>";
    
    size_t total_size = prefix.size() + coord.size() + middle.size() + value.size() * 2 + suffix.size(); // *2 for escaping
    ensureCapacity(total_size);
    
    // 写入前缀
    std::memcpy(buffer_.data() + current_pos_, prefix.data(), prefix.size());
    current_pos_ += prefix.size();
    
    // 写入坐标
    std::memcpy(buffer_.data() + current_pos_, coord.data(), coord.size());
    current_pos_ += coord.size();
    
    // 写入中间部分
    std::memcpy(buffer_.data() + current_pos_, middle.data(), middle.size());
    current_pos_ += middle.size();
    
    // 写入转义的字符串值
    writeEscapedString(value);
    
    // 写入后缀
    std::memcpy(buffer_.data() + current_pos_, suffix.data(), suffix.size());
    current_pos_ += suffix.size();
}

void TXFastXmlWriter::writeEscapedString(std::string_view str) {
    // 🚀 快速XML转义
    for (char c : str) {
        switch (c) {
            case '<':
                writeConstant("&lt;");
                break;
            case '>':
                writeConstant("&gt;");
                break;
            case '&':
                writeConstant("&amp;");
                break;
            case '"':
                writeConstant("&quot;");
                break;
            case '\'':
                writeConstant("&apos;");
                break;
            default:
                ensureCapacity(1);
                buffer_[current_pos_++] = c;
                break;
        }
    }
}

void TXFastXmlWriter::writeNumberCellsBatch(const std::vector<std::string>& coords, 
                                            const double* values, size_t count) {
    // 🚀 批量写入优化
    size_t estimated_size = count * 80; // 每个数值单元格约80字节
    ensureCapacity(estimated_size);
    
    for (size_t i = 0; i < count; ++i) {
        writeNumberCell(coords[i], values[i]);
    }
}

std::vector<uint8_t> TXFastXmlWriter::getResult() && {
    buffer_.resize(current_pos_);
    return std::move(buffer_);
}

void TXFastXmlWriter::clear() {
    current_pos_ = 0;
}

// ==================== TXCoordConverter 实现 ====================

void TXCoordConverter::rowColToString(uint32_t row, uint32_t col, char* buffer, size_t buffer_size) {
    size_t pos = 0;
    
    // 列转换为字母
    columnToLetters(col, buffer, pos);
    
    // 行号转换
    uint32_t row_1based = row + 1;
    size_t number_len = TXFastNumberConverter::uint32ToString(row_1based, buffer + pos, buffer_size - pos);
    pos += number_len;
    
    buffer[pos] = '\0';
}

std::string TXCoordConverter::rowColToString(uint32_t row, uint32_t col) {
    char buffer[16];
    rowColToString(row, col, buffer, sizeof(buffer));
    return std::string(buffer);
}

void TXCoordConverter::columnToLetters(uint32_t col, char* buffer, size_t& pos) {
    uint32_t temp_col = col + 1; // Excel列从1开始
    char temp_buffer[8];
    size_t temp_pos = 0;
    
    do {
        temp_col--;
        temp_buffer[temp_pos++] = 'A' + (temp_col % 26);
        temp_col /= 26;
    } while (temp_col > 0);
    
    // 反转字符串
    for (size_t i = 0; i < temp_pos; ++i) {
        buffer[pos + i] = temp_buffer[temp_pos - 1 - i];
    }
    pos += temp_pos;
}

// ==================== TXFastNumberConverter 实现 ====================

constexpr double TXFastNumberConverter::POWERS_OF_10[];

size_t TXFastNumberConverter::doubleToString(double value, char* buffer, size_t buffer_size) {
    // 🚀 快速双精度转换，优化精度
    return snprintf(buffer, buffer_size, "%.15g", value);
}

size_t TXFastNumberConverter::uint32ToString(uint32_t value, char* buffer, size_t buffer_size) {
    // 🚀 快速整数转换
    return snprintf(buffer, buffer_size, "%u", value);
}

} // namespace TinaXlsx
