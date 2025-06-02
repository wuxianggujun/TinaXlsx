#include "TinaXlsx/TXStreamXmlWriter.hpp"
#include <algorithm>
#include <stdexcept>

namespace TinaXlsx {

// ==================== TXStreamXmlWriter 实现 ====================

TXStreamXmlWriter::TXStreamXmlWriter(std::size_t bufferSize)
    : outputMode_(OutputMode::String)
    , bufferCapacity_(bufferSize)
    , totalBytesWritten_(0)
    , outputBuffer_(nullptr)
    , elementStarted_(false)
    , attributesAllowed_(false) {
    buffer_.reserve(bufferCapacity_);
}

TXStreamXmlWriter::~TXStreamXmlWriter() {
    try {
        flush();
        if (fileStream_.is_open()) {
            fileStream_.close();
        }
    } catch (...) {
        // 析构函数中不抛出异常
    }
}

bool TXStreamXmlWriter::startStringOutput() {
    outputMode_ = OutputMode::String;
    stringStream_.clear();
    stringStream_.str("");
    buffer_.clear();
    totalBytesWritten_ = 0;
    elementStack_.clear();
    elementStarted_ = false;
    attributesAllowed_ = false;
    return true;
}

bool TXStreamXmlWriter::startFileOutput(const std::string& filename) {
    outputMode_ = OutputMode::File;
    fileStream_.open(filename, std::ios::binary | std::ios::out);
    if (!fileStream_.is_open()) {
        return false;
    }
    buffer_.clear();
    totalBytesWritten_ = 0;
    elementStack_.clear();
    elementStarted_ = false;
    attributesAllowed_ = false;
    return true;
}

bool TXStreamXmlWriter::startBufferOutput(std::vector<uint8_t>& buffer) {
    outputMode_ = OutputMode::Buffer;
    outputBuffer_ = &buffer;
    outputBuffer_->clear();
    buffer_.clear();
    totalBytesWritten_ = 0;
    elementStack_.clear();
    elementStarted_ = false;
    attributesAllowed_ = false;
    return true;
}

void TXStreamXmlWriter::writeXmlDeclaration(const std::string& encoding) {
    writeInternal("<?xml version=\"1.0\" encoding=\"" + encoding + "\"?>\n");
}

void TXStreamXmlWriter::startElement(const std::string& name) {
    finishElementStart();
    
    buffer_ += '<';
    buffer_ += name;
    elementStack_.push_back(name);
    elementStarted_ = true;
    attributesAllowed_ = true;
}

void TXStreamXmlWriter::addAttribute(const std::string& name, const std::string& value) {
    if (!attributesAllowed_) {
        throw std::runtime_error("Attributes can only be added immediately after startElement");
    }
    
    buffer_ += ' ';
    buffer_ += name;
    buffer_ += "=\"";
    buffer_ += escapeXml(value);
    buffer_ += '"';
}

void TXStreamXmlWriter::endElement(const std::string& name) {
    if (elementStack_.empty() || elementStack_.back() != name) {
        throw std::runtime_error("Element mismatch: expected " + 
                                (elementStack_.empty() ? "no element" : elementStack_.back()) + 
                                ", got " + name);
    }
    
    if (elementStarted_) {
        // 自闭合标签
        buffer_ += "/>";
        elementStarted_ = false;
        attributesAllowed_ = false;
    } else {
        // 完整结束标签
        buffer_ += "</";
        buffer_ += name;
        buffer_ += '>';
    }
    
    elementStack_.pop_back();
    
    // 检查是否需要刷新缓冲区
    if (buffer_.size() >= bufferCapacity_) {
        flush();
    }
}

void TXStreamXmlWriter::writeText(const std::string& text, bool escapeXml) {
    finishElementStart();
    
    if (escapeXml) {
        buffer_ += this->escapeXml(text);
    } else {
        buffer_ += text;
    }
}

void TXStreamXmlWriter::writeSimpleElement(const std::string& name, const std::string& text,
                                          const std::vector<std::pair<std::string, std::string>>& attributes) {
    startElement(name);
    for (const auto& attr : attributes) {
        addAttribute(attr.first, attr.second);
    }
    if (!text.empty()) {
        writeText(text);
    }
    endElement(name);
}

void TXStreamXmlWriter::flush() {
    if (buffer_.empty()) {
        return;
    }
    
    writeInternal(buffer_);
    buffer_.clear();
}

std::string TXStreamXmlWriter::finish() {
    flush();
    
    if (outputMode_ == OutputMode::String) {
        return stringStream_.str();
    } else if (outputMode_ == OutputMode::File && fileStream_.is_open()) {
        fileStream_.close();
    }
    
    return "";
}

void TXStreamXmlWriter::writeInternal(const std::string& data) {
    totalBytesWritten_ += data.size();
    
    switch (outputMode_) {
        case OutputMode::String:
            stringStream_ << data;
            break;
        case OutputMode::File:
            if (fileStream_.is_open()) {
                fileStream_ << data;
            }
            break;
        case OutputMode::Buffer:
            if (outputBuffer_) {
                const char* dataPtr = data.c_str();
                outputBuffer_->insert(outputBuffer_->end(), dataPtr, dataPtr + data.size());
            }
            break;
    }
}

std::string TXStreamXmlWriter::escapeXml(const std::string& text) const {
    std::string result;
    result.reserve(text.size() * 1.1); // 预留一些空间给转义字符
    
    for (char c : text) {
        switch (c) {
            case '<':
                result += "&lt;";
                break;
            case '>':
                result += "&gt;";
                break;
            case '&':
                result += "&amp;";
                break;
            case '"':
                result += "&quot;";
                break;
            case '\'':
                result += "&apos;";
                break;
            default:
                result += c;
                break;
        }
    }
    
    return result;
}

void TXStreamXmlWriter::finishElementStart() {
    if (elementStarted_) {
        buffer_ += '>';
        elementStarted_ = false;
        attributesAllowed_ = false;
    }
}

// ==================== TXWorksheetStreamWriter 实现 ====================

TXWorksheetStreamWriter::TXWorksheetStreamWriter(std::size_t bufferSize)
    : writer_(bufferSize), inRow_(false) {
}

void TXWorksheetStreamWriter::startWorksheet(const std::string& usedRangeRef) {
    writer_.startStringOutput();
    writer_.writeXmlDeclaration();
    
    writer_.startElement("worksheet");
    writer_.addAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    writer_.addAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    // 添加维度信息
    writer_.writeSimpleElement("dimension", "", {{"ref", usedRangeRef}});
    
    // 开始sheetData
    writer_.startElement("sheetData");
}

void TXWorksheetStreamWriter::startRow(u32 rowNumber) {
    if (inRow_) {
        endRow();
    }
    
    writer_.startElement("row");
    writer_.addAttribute("r", std::to_string(rowNumber));
    inRow_ = true;
}

void TXWorksheetStreamWriter::writeCell(const std::string& cellRef, const std::string& value,
                                       const std::string& cellType, u32 styleIndex) {
    if (!inRow_) {
        throw std::runtime_error("writeCell called outside of row");
    }
    
    writer_.startElement("c");
    writer_.addAttribute("r", cellRef);
    
    if (!cellType.empty()) {
        writer_.addAttribute("t", cellType);
    }
    
    if (styleIndex > 0) {
        writer_.addAttribute("s", std::to_string(styleIndex));
    }
    
    if (!value.empty()) {
        if (cellType == "inlineStr") {
            writer_.startElement("is");
            writer_.writeSimpleElement("t", value);
            writer_.endElement("is");
        } else {
            writer_.writeSimpleElement("v", value);
        }
    }
    
    writer_.endElement("c");
}

void TXWorksheetStreamWriter::endRow() {
    if (inRow_) {
        writer_.endElement("row");
        inRow_ = false;
    }
}

void TXWorksheetStreamWriter::endWorksheet() {
    if (inRow_) {
        endRow();
    }
    
    writer_.endElement("sheetData"); // 结束sheetData
    writer_.endElement("worksheet"); // 结束worksheet
}

std::string TXWorksheetStreamWriter::getXml() {
    return writer_.finish();
}

} // namespace TinaXlsx
