/**
 * @file ZipReader.cpp
 * @brief ZIP文件读取器实现
 */

#include "TinaXlsx/ZipReader.hpp"
#include <fstream>
#include <algorithm>
#include <regex>

// minizip-ng 包含
#include <mz.h>
#include <mz_strm.h>
#include <mz_strm_mem.h>
#include <mz_zip.h>
#include <mz_zip_rw.h>

namespace TinaXlsx {

// =============================================================================
// ZipReader 实现
// =============================================================================

ZipReader::ZipReader(const std::string& filePath) {
    openFromFile(filePath);
}

ZipReader::ZipReader(const void* data, size_t dataSize) {
    openFromMemory(data, dataSize);
}

ZipReader::~ZipReader() {
    cleanup();
}

ZipReader::ZipReader(ZipReader&& other) noexcept
    : zipHandle_(other.zipHandle_), zipStream_(other.zipStream_),
      fileBuffer_(std::move(other.fileBuffer_)), filePath_(std::move(other.filePath_)),
      entries_(std::move(other.entries_)), entriesCached_(other.entriesCached_) {
    other.zipHandle_ = nullptr;
    other.zipStream_ = nullptr;
    other.entriesCached_ = false;
}

ZipReader& ZipReader::operator=(ZipReader&& other) noexcept {
    if (this != &other) {
        cleanup();
        
        zipHandle_ = other.zipHandle_;
        zipStream_ = other.zipStream_;
        fileBuffer_ = std::move(other.fileBuffer_);
        filePath_ = std::move(other.filePath_);
        entries_ = std::move(other.entries_);
        entriesCached_ = other.entriesCached_;
        
        other.zipHandle_ = nullptr;
        other.zipStream_ = nullptr;
        other.entriesCached_ = false;
    }
    return *this;
}

bool ZipReader::isValid() const {
    return zipHandle_ != nullptr;
}

const std::vector<ZipReader::EntryInfo>& ZipReader::getEntries() const {
    if (!entriesCached_) {
        cacheEntries();
    }
    return entries_;
}

bool ZipReader::hasEntry(const std::string& entryName) const {
    if (!zipHandle_ || !isValidEntryName(entryName)) {
        return false;
    }
    
    return mz_zip_reader_locate_entry(zipHandle_, entryName.c_str(), 0) == MZ_OK;
}

std::optional<std::string> ZipReader::readEntry(const std::string& entryName) const {
    auto binaryData = readEntryBinary(entryName);
    if (!binaryData) {
        return std::nullopt;
    }
    
    return std::string(binaryData->begin(), binaryData->end());
}

std::optional<std::vector<uint8_t>> ZipReader::readEntryBinary(const std::string& entryName) const {
    if (!zipHandle_ || !isValidEntryName(entryName)) {
        return std::nullopt;
    }
    
    // 定位文件
    if (mz_zip_reader_locate_entry(zipHandle_, entryName.c_str(), 0) != MZ_OK) {
        return std::nullopt;
    }
    
    // 获取文件信息
    mz_zip_file* fileInfo = nullptr;
    if (mz_zip_reader_entry_get_info(zipHandle_, &fileInfo) != MZ_OK || !fileInfo) {
        return std::nullopt;
    }
    
    // 打开文件
    if (mz_zip_reader_entry_open(zipHandle_) != MZ_OK) {
        return std::nullopt;
    }
    
    // 读取文件内容
    std::vector<uint8_t> buffer(fileInfo->uncompressed_size);
    int32_t bytesRead = mz_zip_reader_entry_read(zipHandle_, 
        buffer.data(), static_cast<int32_t>(buffer.size()));
    
    // 关闭文件
    mz_zip_reader_entry_close(zipHandle_);
    
    if (bytesRead != static_cast<int32_t>(fileInfo->uncompressed_size)) {
        return std::nullopt;
    }
    
    return buffer;
}

std::optional<ZipReader::EntryInfo> ZipReader::getEntryInfo(const std::string& entryName) const {
    if (!zipHandle_ || !isValidEntryName(entryName)) {
        return std::nullopt;
    }
    
    // 定位文件
    if (mz_zip_reader_locate_entry(zipHandle_, entryName.c_str(), 0) != MZ_OK) {
        return std::nullopt;
    }
    
    // 获取文件信息
    mz_zip_file* fileInfo = nullptr;
    if (mz_zip_reader_entry_get_info(zipHandle_, &fileInfo) != MZ_OK || !fileInfo) {
        return std::nullopt;
    }
    
    EntryInfo info;
    info.filename = fileInfo->filename ? fileInfo->filename : "";
    info.compressedSize = fileInfo->compressed_size;
    info.uncompressedSize = fileInfo->uncompressed_size;
    info.isDirectory = info.filename.back() == '/';
    
    return info;
}

std::vector<std::string> ZipReader::listDirectory(const std::string& dirPath) const {
    const auto& allEntries = getEntries();
    std::vector<std::string> result;
    
    std::string normalizedDir = dirPath;
    if (!normalizedDir.empty() && normalizedDir.back() != '/') {
        normalizedDir += '/';
    }
    
    for (const auto& entry : allEntries) {
        if (entry.filename.size() >= normalizedDir.size() && 
            entry.filename.substr(0, normalizedDir.size()) == normalizedDir) {
            std::string relativePath = entry.filename.substr(normalizedDir.length());
            // 只包含直接子项（不包含子目录中的文件）
            if (relativePath.find('/') == std::string::npos || 
                (relativePath.find('/') == relativePath.length() - 1)) {
                result.push_back(entry.filename);
            }
        }
    }
    
    return result;
}

size_t ZipReader::getFileSize() const {
    return fileBuffer_.size();
}

std::vector<std::string> ZipReader::findEntries(const std::string& pattern) const {
    const auto& allEntries = getEntries();
    std::vector<std::string> result;
    
    try {
        // 将通配符模式转换为正则表达式
        std::string regexPattern = pattern;
        std::replace(regexPattern.begin(), regexPattern.end(), '*', '.');
        regexPattern = ".*" + regexPattern + ".*";
        
        std::regex regex(regexPattern, std::regex_constants::icase);
        
        for (const auto& entry : allEntries) {
            if (std::regex_match(entry.filename, regex)) {
                result.push_back(entry.filename);
            }
        }
    } catch (const std::regex_error&) {
        // 正则表达式错误，进行简单的字符串匹配
        for (const auto& entry : allEntries) {
            if (entry.filename.find(pattern) != std::string::npos) {
                result.push_back(entry.filename);
            }
        }
    }
    
    return result;
}

void ZipReader::openFromFile(const std::string& filePath) {
    filePath_ = filePath;
    
    // 检查文件是否存在
    if (!std::filesystem::exists(filePath)) {
        throw FileException("File not found: " + filePath);
    }
    
    // 读取文件到内存
    std::ifstream file(filePath, std::ios::binary | std::ios::ate);
    if (!file.is_open()) {
        throw FileException("Cannot open file: " + filePath);
    }
    
    auto fileSize = file.tellg();
    file.seekg(0, std::ios::beg);
    
    fileBuffer_.resize(static_cast<size_t>(fileSize));
    if (!file.read(reinterpret_cast<char*>(fileBuffer_.data()), fileSize)) {
        throw FileException("Failed to read file: " + filePath);
    }
    
    openFromMemory(fileBuffer_.data(), fileBuffer_.size());
}

void ZipReader::openFromMemory(const void* data, size_t dataSize) {
    cleanup();
    
    // 创建内存流
    zipStream_ = mz_stream_mem_create();
    mz_stream_mem_set_buffer(zipStream_, const_cast<void*>(data), static_cast<int32_t>(dataSize));
    
    if (mz_stream_open(zipStream_, nullptr, MZ_OPEN_MODE_READ) != MZ_OK) {
        throw FileException("Failed to open memory stream");
    }
    
    // 创建ZIP读取器
    zipHandle_ = mz_zip_reader_create();
    
    if (mz_zip_reader_open(zipHandle_, zipStream_) != MZ_OK) {
        cleanup();
        throw FileException("Failed to open ZIP archive");
    }
}

void ZipReader::cleanup() {
    if (zipHandle_) {
        mz_zip_reader_close(zipHandle_);
        mz_zip_reader_delete(&zipHandle_);
        zipHandle_ = nullptr;
    }
    
    if (zipStream_) {
        mz_stream_close(zipStream_);
        mz_stream_mem_delete(&zipStream_);
        zipStream_ = nullptr;
    }
    
    entries_.clear();
    entriesCached_ = false;
}

void ZipReader::cacheEntries() const {
    entries_.clear();
    
    if (!zipHandle_) {
        entriesCached_ = true;
        return;
    }
    
    // 移动到第一个条目
    if (mz_zip_reader_goto_first_entry(zipHandle_) != MZ_OK) {
        entriesCached_ = true;
        return;
    }
    
    do {
        mz_zip_file* fileInfo = nullptr;
        if (mz_zip_reader_entry_get_info(zipHandle_, &fileInfo) == MZ_OK && fileInfo) {
            EntryInfo info;
            info.filename = fileInfo->filename ? fileInfo->filename : "";
            info.compressedSize = fileInfo->compressed_size;
            info.uncompressedSize = fileInfo->uncompressed_size;
            info.isDirectory = info.filename.back() == '/';
            
            entries_.push_back(info);
        }
    } while (mz_zip_reader_goto_next_entry(zipHandle_) == MZ_OK);
    
    entriesCached_ = true;
}

bool ZipReader::isValidEntryName(const std::string& entryName) {
    return !entryName.empty() && entryName.length() < 1024; // 合理的长度限制
}

// =============================================================================
// ExcelZipReader 实现
// =============================================================================

ExcelZipReader::ExcelZipReader(const std::string& filePath) : ZipReader(filePath) {
    parseFileType();
    findWorkbookPath();
}

ExcelZipReader::ExcelZipReader(const void* data, size_t dataSize) : ZipReader(data, dataSize) {
    parseFileType();
    findWorkbookPath();
}

std::optional<std::string> ExcelZipReader::readWorkbook() const {
    if (workbookPath_.empty()) {
        return std::nullopt;
    }
    return readEntry(workbookPath_);
}

std::optional<std::string> ExcelZipReader::readSharedStrings() const {
    // 查找共享字符串文件
    auto relsContent = readWorkbookRelationships();
    if (!relsContent) {
        return std::nullopt;
    }
    
    // 简单解析关系文件，查找共享字符串
    size_t pos = relsContent->find("sharedStrings.xml");
    if (pos == std::string::npos) {
        return std::nullopt;
    }
    
    // 向前查找Target属性
    size_t targetStart = relsContent->rfind("Target=\"", pos);
    if (targetStart == std::string::npos) {
        return std::nullopt;
    }
    
    targetStart += 8; // strlen("Target=\"")
    size_t targetEnd = relsContent->find("\"", targetStart);
    if (targetEnd == std::string::npos) {
        return std::nullopt;
    }
    
    std::string target = relsContent->substr(targetStart, targetEnd - targetStart);
    target = normalizePath("xl/" + target);
    
    return readEntry(target);
}

std::optional<std::string> ExcelZipReader::readWorkbookRelationships() const {
    return readEntry("xl/_rels/workbook.xml.rels");
}

std::optional<std::string> ExcelZipReader::readWorksheet(const std::string& worksheetPath) const {
    return readEntry(worksheetPath);
}

std::vector<std::string> ExcelZipReader::getWorksheetPaths() const {
    auto entries = findEntries("xl/worksheets/sheet*.xml");
    std::vector<std::string> result;
    
    for (const auto& entry : entries) {
        if (entry.find("sheet") != std::string::npos && 
            entry.find(".xml") != std::string::npos) {
            result.push_back(entry);
        }
    }
    
    return result;
}

bool ExcelZipReader::isValidExcelFile() const {
    return fileType_ != ExcelFileType::Unknown && 
           hasEntry("[Content_Types].xml") && 
           !workbookPath_.empty();
}

void ExcelZipReader::parseFileType() {
    auto contentTypes = readEntry("[Content_Types].xml");
    if (!contentTypes) {
        fileType_ = ExcelFileType::Unknown;
        return;
    }
    
    if (contentTypes->find("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml") != std::string::npos) {
        fileType_ = ExcelFileType::XLSX;
    } else if (contentTypes->find("application/vnd.ms-excel.sheet.macroEnabled.main+xml") != std::string::npos) {
        fileType_ = ExcelFileType::XLSM;
    } else if (contentTypes->find("application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml") != std::string::npos) {
        fileType_ = ExcelFileType::XLTX;
    } else if (contentTypes->find("application/vnd.ms-excel.template.macroEnabled.main+xml") != std::string::npos) {
        fileType_ = ExcelFileType::XLTM;
    } else {
        fileType_ = ExcelFileType::Unknown;
    }
}

void ExcelZipReader::findWorkbookPath() {
    auto contentTypes = readEntry("[Content_Types].xml");
    if (!contentTypes) {
        return;
    }
    
    // 查找工作簿的主文件
    std::vector<std::string> patterns = {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
        "application/vnd.ms-excel.template.macroEnabled.main+xml"
    };
    
    for (const auto& pattern : patterns) {
        size_t pos = contentTypes->find(pattern);
        if (pos != std::string::npos) {
            // 向前查找PartName属性
            size_t partNameStart = contentTypes->rfind("PartName=\"", pos);
            if (partNameStart != std::string::npos) {
                partNameStart += 10; // strlen("PartName=\"")
                size_t partNameEnd = contentTypes->find("\"", partNameStart);
                if (partNameEnd != std::string::npos) {
                    workbookPath_ = contentTypes->substr(partNameStart, partNameEnd - partNameStart);
                    workbookPath_ = normalizePath(workbookPath_);
                    return;
                }
            }
        }
    }
    
    // 如果没有找到，使用默认路径
    workbookPath_ = "xl/workbook.xml";
}

std::string ExcelZipReader::normalizePath(const std::string& path) {
    std::string result = path;
    
    // 移除开头的斜杠
    if (!result.empty() && result[0] == '/') {
        result = result.substr(1);
    }
    
    // 统一使用正斜杠
    std::replace(result.begin(), result.end(), '\\', '/');
    
    return result;
}

} // namespace TinaXlsx 