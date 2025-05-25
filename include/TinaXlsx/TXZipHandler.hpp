#pragma once

#include <string>
#include <vector>
#include <memory>
#include <unordered_map>
#include <functional>

namespace TinaXlsx {

/**
 * @brief 高性能ZIP文件处理器
 * 
 * 基于minizip-ng库实现的ZIP文件压缩解压功能，
 * 专门优化用于处理XLSX文件格式（实际上是ZIP文件）
 */
class TXZipHandler {
public:
    /**
     * @brief ZIP文件条目信息
     */
    struct ZipEntry {
        std::string filename;       ///< 文件名
        std::size_t uncompressed_size;   ///< 未压缩大小
        std::size_t compressed_size;     ///< 压缩后大小
        uint64_t modified_date;     ///< 修改时间
        bool is_directory;          ///< 是否为目录
    };

    /**
     * @brief 读取模式枚举
     */
    enum class OpenMode {
        Read,      ///< 只读模式
        Write,     ///< 写入模式
        Append     ///< 追加模式
    };

public:
    TXZipHandler();
    ~TXZipHandler();

    // 禁用拷贝构造和赋值
    TXZipHandler(const TXZipHandler&) = delete;
    TXZipHandler& operator=(const TXZipHandler&) = delete;

    // 支持移动构造和赋值
    TXZipHandler(TXZipHandler&& other) noexcept;
    TXZipHandler& operator=(TXZipHandler&& other) noexcept;

    /**
     * @brief 打开ZIP文件
     * @param filename ZIP文件路径
     * @param mode 打开模式
     * @return 成功返回true，失败返回false
     */
    bool open(const std::string& filename, OpenMode mode = OpenMode::Read);

    /**
     * @brief 关闭ZIP文件
     */
    void close();

    /**
     * @brief 检查文件是否已打开
     * @return 已打开返回true，否则返回false
     */
    bool isOpen() const;

    /**
     * @brief 获取ZIP文件中所有条目列表
     * @return ZIP条目列表
     */
    std::vector<ZipEntry> getEntries() const;

    /**
     * @brief 检查文件是否存在
     * @param filename 文件名
     * @return 存在返回true，否则返回false
     */
    bool hasFile(const std::string& filename) const;

    /**
     * @brief 读取文件内容到字符串
     * @param filename 文件名
     * @return 文件内容，如果失败返回空字符串
     */
    std::string readFileToString(const std::string& filename) const;

    /**
     * @brief 读取文件内容到字节数组
     * @param filename 文件名
     * @return 文件内容字节数组，如果失败返回空向量
     */
    std::vector<uint8_t> readFileToBytes(const std::string& filename) const;

    /**
     * @brief 写入字符串到ZIP文件
     * @param filename 文件名
     * @param content 文件内容
     * @param compression_level 压缩级别 (0-9, 6为默认)
     * @return 成功返回true，失败返回false
     */
    bool writeFile(const std::string& filename, const std::string& content, int compression_level = 6);

    /**
     * @brief 写入字节数组到ZIP文件
     * @param filename 文件名
     * @param data 字节数据
     * @param compression_level 压缩级别 (0-9, 6为默认)
     * @return 成功返回true，失败返回false
     */
    bool writeFile(const std::string& filename, const std::vector<uint8_t>& data, int compression_level = 6);

    /**
     * @brief 删除ZIP文件中的条目
     * @param filename 要删除的文件名
     * @return 成功返回true，失败返回false
     */
    bool removeFile(const std::string& filename);

    /**
     * @brief 获取最后的错误信息
     * @return 错误信息字符串
     */
    const std::string& getLastError() const;

    /**
     * @brief 批量读取文件（高性能版本）
     * @param filenames 要读取的文件名列表
     * @param callback 读取完成后的回调函数
     * @return 成功读取的文件数量
     */
    std::size_t readMultipleFiles(
        const std::vector<std::string>& filenames,
        std::function<void(const std::string&, const std::string&)> callback
    ) const;

    /**
     * @brief 批量写入文件（高性能版本）
     * @param files 文件名到内容的映射
     * @param compression_level 压缩级别
     * @return 成功写入的文件数量
     */
    std::size_t writeMultipleFiles(
        const std::unordered_map<std::string, std::string>& files,
        int compression_level = 6
    );

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

} // namespace TinaXlsx 