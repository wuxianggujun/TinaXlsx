/**
 * @file ZipReader.hpp
 * @brief ZIP文件读取器 - 将ZIP文件操作职责从Reader中抽离
 */

#pragma once

#include "Exception.hpp"
#include <memory>
#include <string>
#include <vector>
#include <optional>
#include <filesystem>

// 前向声明，避免在头文件中包含minizip-ng具体实现
extern "C" {
    struct mz_zip_reader_s;
    typedef struct mz_zip_reader_s mz_zip_reader;
}

namespace TinaXlsx {

/**
 * @brief ZIP文件读取器
 * 封装minizip-ng的使用，专门负责ZIP文件的操作
 */
class ZipReader {
public:
    /**
     * @brief ZIP文件条目信息
     */
    struct EntryInfo {
        std::string filename;
        uint64_t compressedSize;
        uint64_t uncompressedSize;
        bool isDirectory;
    };

private:
    void* zipHandle_ = nullptr;
    void* zipStream_ = nullptr;
    std::vector<uint8_t> fileBuffer_;
    std::string filePath_;
    mutable std::vector<EntryInfo> entries_;
    mutable bool entriesCached_ = false;

public:
    /**
     * @brief 从文件路径构造
     * @param filePath ZIP文件路径
     * @throws FileException 如果文件无法打开
     */
    explicit ZipReader(const std::string& filePath);
    
    /**
     * @brief 从内存数据构造
     * @param data 内存中的ZIP数据
     * @param dataSize 数据大小
     * @throws FileException 如果数据无效
     */
    ZipReader(const void* data, size_t dataSize);
    
    /**
     * @brief 析构函数
     */
    ~ZipReader();
    
    // 禁用拷贝
    ZipReader(const ZipReader&) = delete;
    ZipReader& operator=(const ZipReader&) = delete;
    
    // 启用移动
    ZipReader(ZipReader&& other) noexcept;
    ZipReader& operator=(ZipReader&& other) noexcept;
    
    /**
     * @brief 检查ZIP文件是否有效
     * @return bool 是否有效
     */
    [[nodiscard]] bool isValid() const;
    
    /**
     * @brief 获取ZIP文件中的所有条目
     * @return std::vector<EntryInfo> 条目列表
     */
    [[nodiscard]] const std::vector<EntryInfo>& getEntries() const;
    
    /**
     * @brief 检查指定条目是否存在
     * @param entryName 条目名称
     * @return bool 是否存在
     */
    [[nodiscard]] bool hasEntry(const std::string& entryName) const;
    
    /**
     * @brief 读取指定条目的内容
     * @param entryName 条目名称
     * @return std::optional<std::string> 条目内容，如果不存在则返回空
     */
    [[nodiscard]] std::optional<std::string> readEntry(const std::string& entryName) const;
    
    /**
     * @brief 读取指定条目的二进制内容
     * @param entryName 条目名称
     * @return std::optional<std::vector<uint8_t>> 条目内容，如果不存在则返回空
     */
    [[nodiscard]] std::optional<std::vector<uint8_t>> readEntryBinary(const std::string& entryName) const;
    
    /**
     * @brief 获取指定条目的信息
     * @param entryName 条目名称
     * @return std::optional<EntryInfo> 条目信息，如果不存在则返回空
     */
    [[nodiscard]] std::optional<EntryInfo> getEntryInfo(const std::string& entryName) const;
    
    /**
     * @brief 列出指定目录下的所有条目
     * @param dirPath 目录路径（以'/'结尾）
     * @return std::vector<std::string> 条目名称列表
     */
    [[nodiscard]] std::vector<std::string> listDirectory(const std::string& dirPath) const;
    
    /**
     * @brief 获取ZIP文件路径（如果是从文件打开的）
     * @return const std::string& 文件路径
     */
    [[nodiscard]] const std::string& getFilePath() const { return filePath_; }
    
    /**
     * @brief 获取ZIP文件总大小
     * @return size_t 文件大小
     */
    [[nodiscard]] size_t getFileSize() const;
    
    /**
     * @brief 查找匹配模式的条目
     * @param pattern 匹配模式（支持通配符）
     * @return std::vector<std::string> 匹配的条目名称列表
     */
    [[nodiscard]] std::vector<std::string> findEntries(const std::string& pattern) const;

private:
    /**
     * @brief 从文件打开ZIP
     * @param filePath 文件路径
     */
    void openFromFile(const std::string& filePath);
    
    /**
     * @brief 从内存打开ZIP
     * @param data 内存数据
     * @param dataSize 数据大小
     */
    void openFromMemory(const void* data, size_t dataSize);
    
    /**
     * @brief 清理资源
     */
    void cleanup();
    
    /**
     * @brief 缓存条目列表
     */
    void cacheEntries() const;
    
    /**
     * @brief 验证条目名称
     * @param entryName 条目名称
     * @return bool 是否有效
     */
    [[nodiscard]] static bool isValidEntryName(const std::string& entryName);
};

/**
 * @brief Excel专用ZIP读取器
 * 在ZipReader基础上添加Excel特定的功能
 */
class ExcelZipReader : public ZipReader {
public:
    /**
     * @brief Excel文件类型
     */
    enum class ExcelFileType {
        Unknown,
        XLSX,   // Excel 2007+
        XLSM,   // Excel 2007+ with macros
        XLTX,   // Excel template
        XLTM    // Excel template with macros
    };

private:
    ExcelFileType fileType_ = ExcelFileType::Unknown;
    std::string workbookPath_;

public:
    /**
     * @brief 构造函数
     * @param filePath Excel文件路径
     */
    explicit ExcelZipReader(const std::string& filePath);
    
    /**
     * @brief 构造函数
     * @param data 内存中的Excel数据
     * @param dataSize 数据大小
     */
    ExcelZipReader(const void* data, size_t dataSize);
    
    /**
     * @brief 获取Excel文件类型
     * @return ExcelFileType 文件类型
     */
    [[nodiscard]] ExcelFileType getFileType() const { return fileType_; }
    
    /**
     * @brief 获取工作簿文件路径
     * @return const std::string& 工作簿路径
     */
    [[nodiscard]] const std::string& getWorkbookPath() const { return workbookPath_; }
    
    /**
     * @brief 读取工作簿XML内容
     * @return std::optional<std::string> 工作簿内容
     */
    [[nodiscard]] std::optional<std::string> readWorkbook() const;
    
    /**
     * @brief 读取共享字符串XML内容
     * @return std::optional<std::string> 共享字符串内容
     */
    [[nodiscard]] std::optional<std::string> readSharedStrings() const;
    
    /**
     * @brief 读取工作簿关系XML内容
     * @return std::optional<std::string> 关系内容
     */
    [[nodiscard]] std::optional<std::string> readWorkbookRelationships() const;
    
    /**
     * @brief 读取指定工作表的XML内容
     * @param worksheetPath 工作表路径
     * @return std::optional<std::string> 工作表内容
     */
    [[nodiscard]] std::optional<std::string> readWorksheet(const std::string& worksheetPath) const;
    
    /**
     * @brief 获取所有工作表文件路径
     * @return std::vector<std::string> 工作表路径列表
     */
    [[nodiscard]] std::vector<std::string> getWorksheetPaths() const;
    
    /**
     * @brief 检查是否是有效的Excel文件
     * @return bool 是否有效
     */
    [[nodiscard]] bool isValidExcelFile() const;

private:
    /**
     * @brief 解析Excel文件类型
     */
    void parseFileType();
    
    /**
     * @brief 查找工作簿路径
     */
    void findWorkbookPath();
    
    /**
     * @brief 标准化路径（移除开头的斜杠，统一分隔符）
     * @param path 原始路径
     * @return std::string 标准化后的路径
     */
    [[nodiscard]] static std::string normalizePath(const std::string& path);
};

} // namespace TinaXlsx 