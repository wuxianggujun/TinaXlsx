//
// @file TXSharedStringsStreamWriter.cpp
// @brief 共享字符串流式写入器实现
//

#include "TinaXlsx/TXSharedStringsStreamWriter.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include <sstream>
#include <iomanip>

namespace TinaXlsx {

// TXSharedStringsStreamWriter 实现
TXSharedStringsStreamWriter::TXSharedStringsStreamWriter(size_t bufferSize)
    : writer_(std::make_unique<TXBufferedXmlWriter>(bufferSize))
    , stringCount_(0)
    , documentStarted_(false) {
}

void TXSharedStringsStreamWriter::startDocument(size_t estimatedCount) {
    if (documentStarted_) {
        return; // 已经开始
    }
    
    writer_->clear();
    
    // 写入XML声明
    const char* xmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
    writer_->write(xmlDecl, strlen(xmlDecl));
    
    // 开始sst元素
    std::ostringstream oss;
    oss << "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"";
    
    // 如果有预估数量，添加count属性（可以提高Excel加载性能）
    if (estimatedCount > 0) {
        oss << " count=\"" << estimatedCount << "\" uniqueCount=\"" << estimatedCount << "\"";
    }
    
    oss << ">\n";
    
    std::string sstStart = oss.str();
    writer_->write(sstStart.c_str(), sstStart.length());
    
    documentStarted_ = true;
    stringCount_ = 0;
}

void TXSharedStringsStreamWriter::writeString(const std::string& text, bool preserveSpace) {
    if (!documentStarted_) {
        startDocument(0); // 自动开始文档
    }
    
    // 开始si元素（shared string item）
    const char* siStart = "<si>";
    writer_->write(siStart, strlen(siStart));
    
    // 开始t元素（text）
    if (preserveSpace) {
        const char* tStartWithSpace = "<t xml:space=\"preserve\">";
        writer_->write(tStartWithSpace, strlen(tStartWithSpace));
    } else {
        const char* tStart = "<t>";
        writer_->write(tStart, strlen(tStart));
    }
    
    // 写入转义后的文本内容
    std::string escapedText = escapeXmlText(text);
    writer_->write(escapedText.c_str(), escapedText.length());
    
    // 结束t和si元素
    const char* tEnd = "</t></si>\n";
    writer_->write(tEnd, strlen(tEnd));
    
    stringCount_++;
}

void TXSharedStringsStreamWriter::endDocument() {
    if (!documentStarted_) {
        return;
    }
    
    // 结束sst元素
    const char* sstEnd = "</sst>\n";
    writer_->write(sstEnd, strlen(sstEnd));
    
    documentStarted_ = false;
}

TXResult<void> TXSharedStringsStreamWriter::writeToZip(TXZipArchiveWriter& zipWriter, const std::string& partName) {
    if (documentStarted_) {
        endDocument(); // 确保文档已结束
    }
    
    auto writeResult = zipWriter.write(partName, writer_->getBuffer());
    if (writeResult.isError()) {
        return Err<void>(writeResult.error().getCode(), 
                       "Failed to write " + partName + ": " + writeResult.error().getMessage());
    }
    
    return Ok();
}

void TXSharedStringsStreamWriter::reset() {
    writer_->clear();
    stringCount_ = 0;
    documentStarted_ = false;
}

std::string TXSharedStringsStreamWriter::escapeXmlText(const std::string& text) {
    // 快速检查是否需要转义
    bool needsEscape = false;
    for (char c : text) {
        if (c == '<' || c == '>' || c == '&' || c == '"' || c == '\'') {
            needsEscape = true;
            break;
        }
    }

    // 如果不需要转义，直接返回原字符串
    if (!needsEscape) {
        return text;
    }

    // 需要转义时，预分配更大的空间
    std::string result;
    result.reserve(text.length() * 1.2); // 预留20%空间给转义字符

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

// TXSharedStringsWriterFactory 实现
std::unique_ptr<TXSharedStringsStreamWriter> TXSharedStringsWriterFactory::createWriter(size_t stringCount) {
    // 根据预估字符串数量调整缓冲区大小
    size_t bufferSize = 32 * 1024; // 默认32KB

    if (stringCount > 20000) {
        bufferSize = 512 * 1024; // 512KB
    } else if (stringCount > 5000) {
        bufferSize = 256 * 1024; // 256KB
    } else if (stringCount > 2000) {
        bufferSize = 128 * 1024; // 128KB
    } else if (stringCount > 1000) {
        bufferSize = 64 * 1024; // 64KB
    }

    return std::make_unique<TXSharedStringsStreamWriter>(bufferSize);
}

} // namespace TinaXlsx
