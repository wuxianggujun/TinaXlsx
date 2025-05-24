/**
 * @file Workbook.cpp
 * @brief Excel工作簿类实现 - 统一的读写接口
 */

#include "TinaXlsx/Workbook.hpp"
#include <xlsxwriter.h>
#include <filesystem>

namespace TinaXlsx {

Workbook::Workbook(const std::string& filePath, Mode mode) 
    : filePath_(filePath), mode_(mode) {
    
    // 检查文件路径不能为空
    if (filePath.empty()) {
        throw Exception("File path cannot be empty");
    }
    
    switch (mode_) {
        case Mode::Read:
            if (!std::filesystem::exists(filePath)) {
                throw FileException("File not found: " + filePath);
            }
            reader_ = std::make_unique<Reader>(filePath);
            break;
            
        case Mode::Write:
            writer_ = std::make_unique<Writer>(filePath);
            break;
            
        case Mode::ReadWrite:
            throw Exception("ReadWrite mode not supported yet");
            break;
    }
}

Workbook::Workbook(Workbook&& other) noexcept 
    : filePath_(std::move(other.filePath_))
    , mode_(other.mode_)
    , reader_(std::move(other.reader_))
    , writer_(std::move(other.writer_)) {
}

Workbook& Workbook::operator=(Workbook&& other) noexcept {
    if (this != &other) {
        filePath_ = std::move(other.filePath_);
        mode_ = other.mode_;
        reader_ = std::move(other.reader_);
        writer_ = std::move(other.writer_);
    }
    return *this;
}

Workbook::~Workbook() = default;

Reader& Workbook::getReader() {
    if (!reader_) {
        throw Exception("Workbook not in read mode");
    }
    return *reader_;
}

Writer& Workbook::getWriter() {
    if (!writer_) {
        throw Exception("Workbook not in write mode");
    }
    return *writer_;
}

bool Workbook::canRead() const {
    return reader_ != nullptr;
}

bool Workbook::canWrite() const {
    return writer_ != nullptr;
}

bool Workbook::save() {
    if (writer_) {
        return writer_->save();
    }
    return false;
}

bool Workbook::close() {
    bool result = true;
    
    if (writer_) {
        result = writer_->close();
        writer_.reset();
    }
    
    if (reader_) {
        reader_.reset();
    }
    
    return result;
}

bool Workbook::isClosed() const {
    if (writer_) {
        return writer_->isClosed();
    }
    
    // 对于读取模式，没有"关闭"的概念
    return reader_ == nullptr;
}

std::unique_ptr<Workbook> Workbook::openForRead(const std::string& filePath) {
    return std::make_unique<Workbook>(filePath, Mode::Read);
}

std::unique_ptr<Workbook> Workbook::createForWrite(const std::string& filePath) {
    return std::make_unique<Workbook>(filePath, Mode::Write);
}

} // namespace TinaXlsx 
