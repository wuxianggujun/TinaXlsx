#include "TinaXlsx/TXZipHandler.hpp"
#include <mz.h>
#include <mz_strm.h>
#include <mz_strm_buf.h>
#include <mz_strm_mem.h>
#include <mz_zip.h>
#include <mz_zip_rw.h>
#include <algorithm>
#include <sstream>
#include <cstring>

namespace TinaXlsx {

class TXZipHandler::Impl {
public:
    Impl() : zip_reader_(nullptr), zip_writer_(nullptr), is_open_(false), last_error_("") {}
    
    ~Impl() {
        close();
    }

    bool open(const std::string& filename, TXZipHandler::OpenMode mode) {
        close();
        
        if (mode == TXZipHandler::OpenMode::Read) {
            zip_reader_ = mz_zip_reader_create();
            if (!zip_reader_) {
                last_error_ = "Failed to create zip reader";
                return false;
            }
            
            int32_t result = mz_zip_reader_open_file(zip_reader_, filename.c_str());
            if (result != MZ_OK) {
                last_error_ = "Failed to open zip file for reading: " + std::to_string(result);
                mz_zip_reader_delete(&zip_reader_);
                return false;
            }
        } else {
            zip_writer_ = mz_zip_writer_create();
            if (!zip_writer_) {
                last_error_ = "Failed to create zip writer";
                return false;
            }
            
            int32_t result;
            if (mode == TXZipHandler::OpenMode::Write) {
                result = mz_zip_writer_open_file(zip_writer_, filename.c_str(), 0, 0);
            } else { // Append
                result = mz_zip_writer_open_file(zip_writer_, filename.c_str(), 1, 0);
            }
            
            if (result != MZ_OK) {
                last_error_ = "Failed to open zip file for writing: " + std::to_string(result);
                mz_zip_writer_delete(&zip_writer_);
                return false;
            }
        }

        filename_ = filename;
        mode_ = mode;
        is_open_ = true;
        last_error_.clear();
        return true;
    }

    void close() {
        if (is_open_) {
            if (zip_reader_) {
                mz_zip_reader_close(zip_reader_);
                mz_zip_reader_delete(&zip_reader_);
                zip_reader_ = nullptr;
            }
            if (zip_writer_) {
                mz_zip_writer_close(zip_writer_);
                mz_zip_writer_delete(&zip_writer_);
                zip_writer_ = nullptr;
            }
            is_open_ = false;
        }
    }

    bool isOpen() const {
        return is_open_;
    }

    std::vector<TXZipHandler::ZipEntry> getEntries() const {
        std::vector<TXZipHandler::ZipEntry> entries;
        if (!is_open_ || !zip_reader_) {
            return entries;
        }

        int32_t result = mz_zip_reader_goto_first_entry(zip_reader_);
        while (result == MZ_OK) {
            mz_zip_file* file_info = nullptr;
            result = mz_zip_reader_entry_get_info(zip_reader_, &file_info);
            if (result == MZ_OK && file_info) {
                TXZipHandler::ZipEntry entry;
                entry.filename = file_info->filename ? file_info->filename : "";
                entry.uncompressed_size = file_info->uncompressed_size;
                entry.compressed_size = file_info->compressed_size;
                entry.modified_date = file_info->modified_date;
                entry.is_directory = (file_info->external_fa & 0x10) != 0;
                entries.push_back(entry);
            }
            result = mz_zip_reader_goto_next_entry(zip_reader_);
        }

        return entries;
    }

    bool hasFile(const std::string& filename) const {
        if (!is_open_ || !zip_reader_) {
            return false;
        }

        int32_t result = mz_zip_reader_locate_entry(zip_reader_, filename.c_str(), 1);
        return result == MZ_OK;
    }

    std::string readFileToString(const std::string& filename) const {
        auto data = readFileToBytes(filename);
        if (data.empty()) {
            return "";
        }
        return std::string(data.begin(), data.end());
    }

    std::vector<uint8_t> readFileToBytes(const std::string& filename) const {
        if (!is_open_ || !zip_reader_) {
            const_cast<Impl*>(this)->last_error_ = "Zip file not open for reading";
            return {};
        }

        int32_t result = mz_zip_reader_locate_entry(zip_reader_, filename.c_str(), 1);
        if (result != MZ_OK) {
            const_cast<Impl*>(this)->last_error_ = "File not found in zip: " + filename;
            return {};
        }

        result = mz_zip_reader_entry_open(zip_reader_);
        if (result != MZ_OK) {
            const_cast<Impl*>(this)->last_error_ = "Failed to open entry: " + filename;
            return {};
        }

        mz_zip_file* file_info = nullptr;
        result = mz_zip_reader_entry_get_info(zip_reader_, &file_info);
        if (result != MZ_OK || !file_info) {
            mz_zip_reader_entry_close(zip_reader_);
            const_cast<Impl*>(this)->last_error_ = "Failed to get entry info: " + filename;
            return {};
        }

        std::vector<uint8_t> buffer(file_info->uncompressed_size);
        int32_t bytes_read = mz_zip_reader_entry_read(zip_reader_, buffer.data(), buffer.size());
        mz_zip_reader_entry_close(zip_reader_);

        if (bytes_read < 0) {
            const_cast<Impl*>(this)->last_error_ = "Failed to read entry: " + filename;
            return {};
        }

        buffer.resize(bytes_read);
        const_cast<Impl*>(this)->last_error_.clear();
        return buffer;
    }

    bool writeFile(const std::string& filename, const std::string& content, int compression_level) {
        return writeFile(filename, std::vector<uint8_t>(content.begin(), content.end()), compression_level);
    }

    bool writeFile(const std::string& filename, const std::vector<uint8_t>& data, int compression_level) {
        if (!is_open_ || !zip_writer_ || mode_ == TXZipHandler::OpenMode::Read) {
            last_error_ = "Zip file not open for writing";
            return false;
        }

        mz_zip_file file_info = {};
        file_info.filename = filename.c_str();
        file_info.compression_method = MZ_COMPRESS_METHOD_DEFLATE;
        file_info.uncompressed_size = data.size();
        
        time_t current_time = time(nullptr);
        file_info.modified_date = current_time;

        mz_zip_writer_set_compress_level(zip_writer_, compression_level);
        
        int32_t result = mz_zip_writer_entry_open(zip_writer_, &file_info);
        if (result != MZ_OK) {
            last_error_ = "Failed to open entry for writing: " + filename;
            return false;
        }

        if (!data.empty()) {
            result = mz_zip_writer_entry_write(zip_writer_, data.data(), data.size());
            if (result != MZ_OK) {
                mz_zip_writer_entry_close(zip_writer_);
                last_error_ = "Failed to write entry: " + filename;
                return false;
            }
        }

        result = mz_zip_writer_entry_close(zip_writer_);
        if (result != MZ_OK) {
            last_error_ = "Failed to close entry: " + filename;
            return false;
        }

        last_error_.clear();
        return true;
    }

    bool removeFile(const std::string& filename) {
        last_error_ = "Remove file operation not supported in current implementation";
        return false;
    }

    std::size_t readMultipleFiles(
        const std::vector<std::string>& filenames,
        std::function<void(const std::string&, const std::string&)> callback
    ) const {
        std::size_t success_count = 0;
        for (const auto& filename : filenames) {
            auto content = readFileToString(filename);
            if (!content.empty() || last_error_.empty()) {
                callback(filename, content);
                ++success_count;
            }
        }
        return success_count;
    }

    std::size_t writeMultipleFiles(
        const std::unordered_map<std::string, std::string>& files,
        int compression_level
    ) {
        std::size_t success_count = 0;
        for (const auto& pair : files) {
            if (writeFile(pair.first, pair.second, compression_level)) {
                ++success_count;
            }
        }
        return success_count;
    }

    const std::string& getLastError() const {
        return last_error_;
    }

private:
    void* zip_reader_;
    void* zip_writer_;
    bool is_open_;
    std::string filename_;
    TXZipHandler::OpenMode mode_;
    std::string last_error_;
};

// TXZipHandler 实现
TXZipHandler::TXZipHandler() : pImpl(std::make_unique<Impl>()) {}

TXZipHandler::~TXZipHandler() = default;

TXZipHandler::TXZipHandler(TXZipHandler&& other) noexcept : pImpl(std::move(other.pImpl)) {}

TXZipHandler& TXZipHandler::operator=(TXZipHandler&& other) noexcept {
    if (this != &other) {
        pImpl = std::move(other.pImpl);
    }
    return *this;
}

bool TXZipHandler::open(const std::string& filename, OpenMode mode) {
    return pImpl->open(filename, mode);
}

void TXZipHandler::close() {
    pImpl->close();
}

bool TXZipHandler::isOpen() const {
    return pImpl->isOpen();
}

std::vector<TXZipHandler::ZipEntry> TXZipHandler::getEntries() const {
    return pImpl->getEntries();
}

bool TXZipHandler::hasFile(const std::string& filename) const {
    return pImpl->hasFile(filename);
}

std::string TXZipHandler::readFileToString(const std::string& filename) const {
    return pImpl->readFileToString(filename);
}

std::vector<uint8_t> TXZipHandler::readFileToBytes(const std::string& filename) const {
    return pImpl->readFileToBytes(filename);
}

bool TXZipHandler::writeFile(const std::string& filename, const std::string& content, int compression_level) {
    return pImpl->writeFile(filename, content, compression_level);
}

bool TXZipHandler::writeFile(const std::string& filename, const std::vector<uint8_t>& data, int compression_level) {
    return pImpl->writeFile(filename, data, compression_level);
}

bool TXZipHandler::removeFile(const std::string& filename) {
    return pImpl->removeFile(filename);
}

const std::string& TXZipHandler::getLastError() const {
    return pImpl->getLastError();
}

std::size_t TXZipHandler::readMultipleFiles(
    const std::vector<std::string>& filenames,
    std::function<void(const std::string&, const std::string&)> callback
) const {
    return pImpl->readMultipleFiles(filenames, callback);
}

std::size_t TXZipHandler::writeMultipleFiles(
    const std::unordered_map<std::string, std::string>& files,
    int compression_level
) {
    return pImpl->writeMultipleFiles(files, compression_level);
}

} // namespace TinaXlsx 