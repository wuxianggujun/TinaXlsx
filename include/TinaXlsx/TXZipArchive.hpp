#pragma once

#include <mz_strm.h>
#include <mz.h>
#include <mz_os.h>
#include <mz_zip.h>
#include <mz_zip_rw.h>
#include <mz_strm_os.h>
#include <mz_strm_mem.h>

#include <string>
#include <vector>
#include <unordered_map>
#include <cstdint>
#include <ctime>
#include <functional>
#include <cstring>

#include "TXResult.hpp"

namespace TinaXlsx
{
    // ────────────────────────────────────────────────────────────────────────────
    //  Shared types
    // ────────────────────────────────────────────────────────────────────────────
    struct ZipEntry
    {
        std::string filename;
        std::size_t uncompressed_size = 0;
        std::size_t compressed_size = 0;
        uint64_t modified_date = 0; // DOS epoch time (local)
        bool is_directory = false;
    };

    // ────────────────────────────────────────────────────────────────────────────
    //  RAII handle wrapper (minizip uses C pointers)
    // ────────────────────────────────────────────────────────────────────────────

    template <void* (*Create)(), void (*Destroy)(void**)>
    class unique_mz_handle
    {
    public:
        unique_mz_handle() : h_(Create())
        {
        }

        ~unique_mz_handle() { reset(); }

        // 移动构造和移动赋值
        unique_mz_handle(unique_mz_handle&& o) noexcept : h_(o.h_) { o.h_ = nullptr; }

        unique_mz_handle& operator=(unique_mz_handle&& o) noexcept
        {
            if (this != &o)
            {
                reset();
                h_ = o.h_;
                o.h_ = nullptr;
            }
            return *this;
        }

        // 禁止拷贝构造和拷贝赋值
        unique_mz_handle(const unique_mz_handle&) = delete;
        unique_mz_handle& operator=(const unique_mz_handle&) = delete;

        void* get() const { return h_; }

        void reset()
        {
            if (h_) Destroy(&h_);
            h_ = nullptr;
        }

        explicit operator bool() const { return h_ != nullptr; }

    private:
        void* h_ = nullptr;
    };

    // Helper to convert std::time_t → MS‑DOS timestamp expected by minizip.
    inline static uint64_t to_dos_datetime(std::time_t t)
    {
        struct tm tm_buf;
        struct tm* lt;
#ifdef _WIN32
        if (localtime_s(&tm_buf, &t) != 0) return 0;
        lt = &tm_buf;
#else
    lt = std::localtime_r(&t, &tm_buf);
    if (!lt) return 0;
#endif
        // MS-DOS 日期格式: 年份 (从1980起)，月份 (1-12)，日期 (1-31)
        uint16_t dos_date = static_cast<uint16_t>(((lt->tm_year + 1900 - 1980) << 9) | // 年份从1900开始，MS-DOS从1980开始
            ((lt->tm_mon + 1) << 5) | // 月份从0开始，MS-DOS从1开始
            lt->tm_mday);
        // MS-DOS 时间格式: 小时 (0-23)，分钟 (0-59)，秒/2 (0-29)
        uint16_t dos_time = static_cast<uint16_t>((lt->tm_hour << 11) |
            (lt->tm_min << 5) |
            (lt->tm_sec / 2));
        return (static_cast<uint64_t>(dos_date) << 16) | dos_time;
    }

    // ────────────────────────────────────────────────────────────────────────────
    //  Reader implementation
    // ────────────────────────────────────────────────────────────────────────────
    class TXZipArchiveReader
    {
    public:
        TXZipArchiveReader() = default;
        ~TXZipArchiveReader() { close(); }

        TXZipArchiveReader(const TXZipArchiveReader&) = delete;
        TXZipArchiveReader& operator=(const TXZipArchiveReader&) = delete;

        TXZipArchiveReader(TXZipArchiveReader&& o) noexcept { *this = std::move(o); }

        TXZipArchiveReader& operator=(TXZipArchiveReader&& o) noexcept
        {
            if (this != &o)
            {
                close();
                reader_ = std::move(o.reader_);
                is_open_ = o.is_open_;
                filename_ = std::move(o.filename_);
                o.is_open_ = false;
            }
            return *this;
        }

        [[nodiscard]] TXResult<void> open(const std::string& file)
        {
            close();
            reader_ = ReaderHandle();
            if (!reader_)
            {
                return Err(TX_ERROR_CREATE(TXErrorCode::ZipOpenFailed,
                                           "mz_zip_reader_create failed (internal minizip-ng error)"));
            }
            int32_t err = mz_zip_reader_open_file(reader_.get(), file.c_str());
            if (err != MZ_OK)
            {
                std::string error_message = "Cannot open ZIP archive: " + file +
                    " (minizip-ng error code: " + std::to_string(err) + ")";
                return Err<void>(TX_ERROR_CREATE(TXErrorCode::ZipOpenFailed, error_message));
            }
            filename_ = file;
            is_open_ = true;
            return Ok();
        }

        void close()
        {
            if (reader_ && is_open_)
            {
                mz_zip_reader_close(reader_.get());
            }
            is_open_ = false;
            filename_.clear();
        }

        [[nodiscard]] bool isOpen() const { return is_open_; }

        // Enumerate entries – O(n) with no allocations apart from vector growth.
        [[nodiscard]] TXResult<std::vector<ZipEntry>> entries()
        {
            auto open_check = ensureOpen();
            if (open_check.isError()) return Err<std::vector<ZipEntry>>(open_check.error());

            std::vector<ZipEntry> out_entries;
            int32_t err = mz_zip_reader_goto_first_entry(reader_.get());

            if (err == MZ_END_OF_LIST)
            {
                return Ok(std::move(out_entries));
            }

            if (err != MZ_OK)
            {
                // 其他错误
                std::string error_message = "Failed to go to first entry in ZIP (minizip-ng error code: " +
                    std::to_string(err) + ")";
                return Err<std::vector<ZipEntry>>(TX_ERROR_CREATE(TXErrorCode::OperationFailed, error_message));
            }

            do
            {
                mz_zip_file* info = nullptr;
                err = mz_zip_reader_entry_get_info(reader_.get(), &info);
                if (err == MZ_OK && info)
                {
                    ZipEntry e;
                    e.filename = info->filename ? info->filename : "";
                    e.uncompressed_size = static_cast<std::size_t>(info->uncompressed_size);
                    e.compressed_size = static_cast<std::size_t>(info->compressed_size);
                    e.modified_date = info->modified_date;
                    e.is_directory = (info->filename && info->filename[strlen(info->filename) - 1] == '/');
                    out_entries.emplace_back(std::move(e));
                }
                else
                {
                    std::string error_message = "Failed to get entry info during ZIP iteration (minizip-ng error code: "
                        + std::to_string(err) + ")";
                    return Err<std::vector<ZipEntry>>(TX_ERROR_CREATE(TXErrorCode::OperationFailed, error_message));
                }
                err = mz_zip_reader_goto_next_entry(reader_.get());
            }
            while (err == MZ_OK);

            if (err != MZ_END_OF_LIST)
            {
                // 如果循环不是因为到达列表末尾而正常结束
                std::string error_message = "Error iterating ZIP entries (minizip-ng error code: " + std::to_string(err)
                    + ")";
                return Err<std::vector<ZipEntry>>(TX_ERROR_CREATE(TXErrorCode::OperationFailed, error_message));
            }

            return Ok(std::move(out_entries));
        }

        [[nodiscard]] TXResult<bool> has(const std::string& entry_name)
        {
            auto open_check = ensureOpen();
            if (open_check.isError())
            {
                return Err<bool>(open_check.error());
            }
            int32_t err = mz_zip_reader_locate_entry(reader_.get(), entry_name.c_str(), 1);
            if (err == MZ_OK)
            {
                return Ok(true);
            }
            // 其他错误码表示定位操作本身失败
            std::string error_message = "Failed to locate entry '" + entry_name + "' (minizip-ng error code: " +
                std::to_string(err) + ")";
            return Err<bool>(TX_ERROR_CREATE(TXErrorCode::OperationFailed, error_message));
        }

        [[nodiscard]] TXResult<std::vector<uint8_t>> read(const std::string& entry_name)
        {
            auto open_check = ensureOpen();
            if (open_check.isError())
            {
                return Err<std::vector<uint8_t>>(open_check.error());
            }
            int32_t err = mz_zip_reader_locate_entry(reader_.get(), entry_name.c_str(), 1);

            if (err != MZ_OK)
            {
                std::string error_message = "Failed to locate entry '" + entry_name +
                    "' for reading (minizip-ng error code: " + std::to_string(err) + ")";
                return Err<std::vector<uint8_t>>(TX_ERROR_CREATE(TXErrorCode::ZipReadEntryFailed, error_message));
            }
            err = mz_zip_reader_entry_open(reader_.get());

            if (err != MZ_OK)
            {
                std::string error_message = "Cannot open ZIP entry '" + entry_name +
                    "' for reading (minizip-ng error code: " + std::to_string(err) + ")";
                return Err<std::vector<uint8_t>>(TX_ERROR_CREATE(TXErrorCode::ZipReadEntryFailed, error_message));
            }

            mz_zip_file* info = nullptr;
            err = mz_zip_reader_entry_get_info(reader_.get(), &info);
            if (err != MZ_OK || !info)
            {
                mz_zip_reader_entry_close(reader_.get());
                std::string error_message = "Failed to get ZIP entry info for '" + entry_name +
                    "' (minizip-ng error code: " + std::to_string(err) + ")";
                return Err<std::vector<uint8_t>>(TX_ERROR_CREATE(TXErrorCode::ZipReadEntryFailed, error_message));
            }
            std::vector<uint8_t> data_buffer(static_cast<std::size_t>(info->uncompressed_size));
            int32_t bytes_read = mz_zip_reader_entry_read(reader_.get(), data_buffer.data(),
                                                          static_cast<int32_t>(data_buffer.size()));

            int32_t close_err = mz_zip_reader_entry_close(reader_.get());
            if (close_err != MZ_OK && bytes_read >= 0)
            {
                // 如果读取成功但关闭失败，也报告问题
                std::string error_message = "Failed to close ZIP entry '" + entry_name +
                    "' after reading (minizip-ng error code: " + std::to_string(close_err) + ")";
                return Err<std::vector<uint8_t>>(TX_ERROR_CREATE(TXErrorCode::ZipReadEntryFailed, error_message));
            }

            if (bytes_read < 0)
            {
                std::string error_message = "Failed to read data from ZIP entry '" + entry_name +
                    "' (minizip-ng error code: " + std::to_string(bytes_read) + ")";
                return Err<std::vector<uint8_t>>(TX_ERROR_CREATE(TXErrorCode::ZipReadEntryFailed, error_message));
            }

            if (static_cast<std::size_t>(bytes_read) != data_buffer.size())
            {
                return Err<std::vector<uint8_t>>(TX_ERROR_CREATE(TXErrorCode::ZipOpenFailed,
                                                                 "Incomplete read from ZIP entry '" + entry_name +
                                                                 "'. Expected " + std::to_string(data_buffer.size()) +
                                                                 ", got " + std::to_string(bytes_read)));
            }
            return Ok(std::move(data_buffer));
        }

        [[nodiscard]] TXResult<std::string> readString(const std::string& entry_name)
        {
            TXResult<std::vector<uint8_t>> read_result = read(entry_name);
            if (read_result.isError())
            {
                return Err<std::string>(read_result.error());
            }
            const auto& bytes = read_result.value();
            return Ok(std::string(bytes.begin(), bytes.end()));
        }

    private:
        using ReaderHandle = unique_mz_handle<mz_zip_reader_create, mz_zip_reader_delete>;
        ReaderHandle reader_;
        bool is_open_ = false;
        std::string filename_;

        [[nodiscard]] TXResult<void> ensureOpen() const
        {
            if (!is_open_)
            {
                return Err(TXErrorCode::ZipInvalidState, "Archive is not open. Call open() first.");
            }
            return Ok();
        }
    };

    // ────────────────────────────────────────────────────────────────────────────
    //  TXZipArchiveWriter: ZIP 归档写入器实现
    // ────────────────────────────────────────────────────────────────────────────
    class TXZipArchiveWriter
    {
    public:
        TXZipArchiveWriter() = default;
        ~TXZipArchiveWriter() { close(); }

        // 禁止拷贝构造和拷贝赋值
        TXZipArchiveWriter(const TXZipArchiveWriter&) = delete;
        TXZipArchiveWriter& operator=(const TXZipArchiveWriter&) = delete;

        // 支持移动构造和移动赋值
        TXZipArchiveWriter(TXZipArchiveWriter&& o) noexcept { *this = std::move(o); }

        TXZipArchiveWriter& operator=(TXZipArchiveWriter&& o) noexcept
        {
            if (this != &o)
            {
                close(); // 关闭当前实例（如果已打开）
                writer_ = std::move(o.writer_);
                is_open_ = o.is_open_;
                filename_ = std::move(o.filename_);
                // last_error_ 成员是内部实现细节，用于构建 TXError，不参与移动赋值
                o.is_open_ = false; // 源对象置于有效但关闭的状态
                o.filename_.clear();
            }
            return *this;
        }

        /**
         * @brief 打开ZIP归档文件进行写入或追加。
         * @param file 归档路径，用于创建或追加。
         * @param append 如果为true，则打开现有归档并追加；否则，如果文件存在则截断，不存在则创建。
         * @param level Deflate 压缩级别 (0‑9)。0 = 仅存储 (不压缩)。默认级别通常是 6。
         * @return TXResult<void> 成功则Ok()，失败则Err(TXError)。
         */
        [[nodiscard]] TXResult<void> open(const std::string& file, bool append = false,
                                          int16_t level = 6 /* MZ_DEFAULT_COMPRESSION */)
        {
            close(); // 关闭任何先前打开的归档
            
            // 如果不是追加模式且文件存在，先删除它以确保完全重写
            if (!append && mz_os_file_exists(file.c_str()) == MZ_OK) {
                if (mz_os_unlink(file.c_str()) != MZ_OK) {
                    return Err(TX_ERROR_CREATE(TXErrorCode::ZipCreateFailed,
                                               "Failed to delete existing file: " + file));
                }
            }
            
            writer_ = WriterHandle(); // 调用 unique_mz_handle 构造函数, 内部调用 mz_zip_writer_create
            if (!writer_)
            {
                // TX_ERROR_CREATE 宏会自动填充文件/行号等信息
                return Err(TX_ERROR_CREATE(TXErrorCode::ZipCreateFailed,
                                           "mz_zip_writer_create failed (internal minizip-ng error)"));
            }

            mz_zip_writer_set_compress_level(writer_.get(), level);
            
            // 设置覆盖回调以确保正确的覆盖行为
            if (!append) {
                mz_zip_writer_set_overwrite_cb(writer_.get(), nullptr, 
                    [](void*, void*, const char*) -> int32_t { return MZ_OK; });
            }

            // append 参数：MZ_OPEN_MODE_CREATE (0) 表示创建/截断，MZ_OPEN_MODE_APPEND 表示追加
            int32_t open_mode = append ? MZ_OPEN_MODE_APPEND : MZ_OPEN_MODE_CREATE;
            // 第三个参数 disk_size 对于 mz_zip_writer_open_file 意义不大，通常设为0。
            // 第四个参数 zip_cd 用于指定是否创建 ZIP64 中心目录记录，0表示自动。
            int32_t err = mz_zip_writer_open_file(writer_.get(), file.c_str(),
                                                  0 /* disk_size for split archives, not used here */, open_mode);
            if (err != MZ_OK)
            {
                std::string error_message = "Cannot open ZIP archive for writing: " + file +
                    " (minizip-ng error code: " + std::to_string(err) + ")";
                TXErrorCode code_to_use = append ? TXErrorCode::ZipOpenFailed : TXErrorCode::ZipCreateFailed;
                return Err(TX_ERROR_CREATE(code_to_use, error_message));
            }
            filename_ = file;
            is_open_ = true;
            return Ok(); // 成功打开
        }

        /**
         * @brief 关闭当前打开的ZIP归档，完成所有写入操作。
         * 如果未打开，则此操作无效果。
         */
        void close()
        {
            if (writer_ && is_open_)
            {
                // 确保句柄有效且归档已打开
                mz_zip_writer_close(writer_.get());
            }
            is_open_ = false;
            filename_.clear();
        }

        /**
         * @brief 检查归档是否已成功打开以供写入。
         * @return 如果归档已打开，则为 true；否则为 false。
         */
        [[nodiscard]] bool isOpen() const { return is_open_; }

        /**
         * @brief 将内存中的字节向量作为条目写入ZIP归档。
         * @param entry_name 要在归档中创建的条目名称 (UTF‑8 编码)。
         * @param data 包含要写入数据的字节向量。
         * @param mtimeSec 条目的UNIX时间戳（秒）；如果为0，则使用当前系统时间。
         * @return TXResult<void> 成功则Ok()，失败则Err(TXError)。
         */
        [[nodiscard]] TXResult<void> write(const std::string& entry_name,
                                           const std::vector<uint8_t>& data,
                                           std::time_t mtimeSec = 0)
        {
            return write(entry_name, data.data(), data.size(), mtimeSec);
        }

        /**
         * @brief 将内存中的原始缓冲区作为条目写入ZIP归档。
         * @param entry_name 要在归档中创建的条目名称 (UTF‑8 编码)。
         * @param buf 指向要写入数据的缓冲区的指针。
         * @param size 要写入数据的大小（字节）。
         * @param mtimeSec 条目的UNIX时间戳（秒）；如果为0，则使用当前系统时间。
         * @return TXResult<void> 成功则Ok()，失败则Err(TXError)。
         */
        [[nodiscard]] TXResult<void> write(const std::string& entry_name,
                                           const void* buf,
                                           std::size_t size,
                                           std::time_t mtimeSec = 0)
        {
            auto open_check = ensureOpen_(); // 内部检查归档是否打开
            if (open_check.isError())
            {
                return open_check; // 直接返回打开错误
            }

            mz_zip_file file_info{}; // 初始化结构体
            file_info.filename = entry_name.c_str();
            file_info.version_madeby = MZ_VERSION_MADEBY; // 使用 minizip-ng 定义的版本

            file_info.compression_method = static_cast<uint8_t>(MZ_COMPRESS_METHOD_DEFLATE); // Deflate 压缩

            file_info.modified_date = to_dos_datetime(mtimeSec != 0 ? mtimeSec : std::time(nullptr));
            file_info.flag = MZ_ZIP_FLAG_UTF8; // 推荐使用 UTF-8 编码文件名
            file_info.uncompressed_size = static_cast<int64_t>(size);

            // mz_zip_writer_add_buffer(void *handle, const void *buf, uint32_t len, const mz_zip_file *file_info)
            int32_t err = mz_zip_writer_add_buffer(writer_.get(), const_cast<void*>(buf), static_cast<uint32_t>(size),
                                                   &file_info);
            if (err != MZ_OK)
            {
                std::string error_message = "Failed to add buffer to ZIP as entry '" + entry_name +
                    "' (minizip-ng error code: " + std::to_string(err) + ")";
                // 假设您已添加 TXErrorCode::ZipWriteEntryFailed
                return Err(TX_ERROR_CREATE(TXErrorCode::ZipWriteEntryFailed, error_message));
            }
            return Ok(); // 成功写入
        }

        /**
         * @brief 将磁盘上的文件直接添加为ZIP归档中的一个条目（可能使用流式处理，效率较高）。
         * @param entry_name 要在归档中创建的条目名称。
         * @param src_path 源文件的磁盘路径。
         * @return TXResult<void> 成功则Ok()，失败则Err(TXError)。
         */
        [[nodiscard]] TXResult<void> writeFile(const std::string& entry_name, const std::string& src_path)
        {
            auto open_check = ensureOpen_();
            if (open_check.isError())
            {
                return open_check;
            }

            // mz_zip_writer_add_file(void *handle, const char *path_to_add, const char *filename_in_zip)
            // path_to_add 是磁盘上的源文件路径, filename_in_zip 是在归档中的名称
            int32_t err = mz_zip_writer_add_file(writer_.get(), src_path.c_str(), entry_name.c_str());
            if (err != MZ_OK)
            {
                std::string error_message = "Failed to add file '" + src_path + "' as entry '" + entry_name +
                    "' to ZIP (minizip-ng error code: " + std::to_string(err) + ")";
                return Err(TX_ERROR_CREATE(TXErrorCode::ZipWriteEntryFailed, error_message));
            }
            return Ok(); // 成功添加文件
        }

    private:
        using WriterHandle = unique_mz_handle<mz_zip_writer_create, mz_zip_writer_delete>; ///< minizip 写入器句柄的 RAII 包装。
        WriterHandle writer_; ///< 底层 minizip-ng 写入器句柄。
        bool is_open_ = false; ///< 标记归档是否已打开以供写入。
        std::string filename_; ///< 当前打开的归档文件名。

        /**
         * @brief 内部辅助函数，确保归档当前已打开以供写入。
         * @return TXResult<void> 如果归档已打开则Ok()；否则Err(TXError)。
         */
        [[nodiscard]] TXResult<void> ensureOpen_() const
        {
            // 标记为 const，因为它不修改对象状态
            if (!is_open_)
            {
                // 假设您已添加 TXErrorCode::ZipInvalidState
                return Err(TXErrorCode::ZipInvalidState, "Archive is not open for writing. Call open() first.");
            }
            return Ok();
        }
    };
} // namespace TinaXlsx
