//
// @file TXExcelIO.hpp
// @brief 🚀 Excel文件I/O处理器 - 高性能Excel文件读写
//

#pragma once

#include <TinaXlsx/user/TXWorkbook.hpp>
#include <TinaXlsx/TXResult.hpp>
#include <TinaXlsx/TXError.hpp>
#include <TinaXlsx/TXUnifiedMemoryManager.hpp>
#include <TinaXlsx/TXVector.hpp>
#include <string>
#include <memory>

namespace TinaXlsx {

/**
 * @brief 🚀 Excel文件I/O处理器
 * 
 * 设计理念：
 * - 高性能Excel文件读写
 * - 支持.xlsx格式（基于OpenXML）
 * - 流式处理，内存友好
 * - 与高性能内存引擎无缝集成
 * 
 * 使用示例：
 * ```cpp
 * // 读取Excel文件
 * auto workbook = TXExcelIO::loadFromFile("input.xlsx");
 * if (workbook.isOk()) {
 *     auto wb = workbook.value();
 *     // 使用高性能引擎处理数据
 *     wb->getSheet("Sheet1")->cell("A1").setValue(100.0);
 * }
 * 
 * // 保存Excel文件
 * auto result = TXExcelIO::saveToFile(wb.get(), "output.xlsx");
 * ```
 */
class TXExcelIO {
public:
    // ==================== 文件格式枚举 ====================
    
    enum class FileFormat {
        XLSX,           // Excel 2007+ (.xlsx)
        XLS,            // Excel 97-2003 (.xls) - 暂不支持
        CSV,            // 逗号分隔值 (.csv)
        AUTO_DETECT     // 自动检测格式
    };
    
    // ==================== 读取选项 ====================
    
    struct ReadOptions {
        bool read_formulas = true;          // 是否读取公式
        bool read_styles = false;           // 是否读取样式（暂不支持）
        bool read_comments = false;         // 是否读取注释（暂不支持）
        bool read_hidden_sheets = true;     // 是否读取隐藏工作表
        size_t max_rows = 1048576;          // 最大行数限制
        size_t max_cols = 16384;            // 最大列数限制
        std::string password;               // 密码（暂不支持）
        
        ReadOptions() = default;
    };
    
    // ==================== 写入选项 ====================
    
    struct WriteOptions {
        bool write_formulas = true;         // 是否写入公式
        bool write_styles = false;          // 是否写入样式（暂不支持）
        bool compress = true;               // 是否压缩
        FileFormat format = FileFormat::XLSX; // 输出格式
        std::string creator = "TinaXlsx";   // 创建者信息
        
        WriteOptions() = default;
    };

    // ==================== 静态读取方法 ====================
    
    /**
     * @brief 🚀 从文件加载工作簿
     * @param file_path 文件路径
     * @param options 读取选项
     * @return 工作簿智能指针或错误
     */
    static TXResult<std::unique_ptr<TXWorkbook>> loadFromFile(
        const std::string& file_path,
        const ReadOptions& options = ReadOptions{}
    );
    
    /**
     * @brief 🚀 从内存数据加载工作簿
     * @param data 内存数据
     * @param size 数据大小
     * @param options 读取选项
     * @return 工作簿智能指针或错误
     */
    static TXResult<std::unique_ptr<TXWorkbook>> loadFromMemory(
        const void* data,
        size_t size,
        const ReadOptions& options = ReadOptions{}
    );
    
    /**
     * @brief 🚀 从流加载工作簿
     * @param stream 输入流
     * @param options 读取选项
     * @return 工作簿智能指针或错误
     */
    static TXResult<std::unique_ptr<TXWorkbook>> loadFromStream(
        std::istream& stream,
        const ReadOptions& options = ReadOptions{}
    );

    // ==================== 静态写入方法 ====================
    
    /**
     * @brief 🚀 保存工作簿到文件
     * @param workbook 工作簿指针
     * @param file_path 文件路径
     * @param options 写入选项
     * @return 操作结果
     */
    static TXResult<void> saveToFile(
        const TXWorkbook* workbook,
        const std::string& file_path,
        const WriteOptions& options = WriteOptions{}
    );
    
    /**
     * @brief 🚀 保存工作簿到内存
     * @param workbook 工作簿指针
     * @param options 写入选项
     * @return 内存数据或错误
     */
    static TXResult<TXVector<uint8_t>> saveToMemory(
        const TXWorkbook* workbook,
        const WriteOptions& options = WriteOptions{}
    );
    
    /**
     * @brief 🚀 保存工作簿到流
     * @param workbook 工作簿指针
     * @param stream 输出流
     * @param options 写入选项
     * @return 操作结果
     */
    static TXResult<void> saveToStream(
        const TXWorkbook* workbook,
        std::ostream& stream,
        const WriteOptions& options = WriteOptions{}
    );

    // ==================== 格式检测 ====================
    
    /**
     * @brief 🚀 检测文件格式
     * @param file_path 文件路径
     * @return 文件格式
     */
    static FileFormat detectFormat(const std::string& file_path);
    
    /**
     * @brief 🚀 检测内存数据格式
     * @param data 内存数据
     * @param size 数据大小
     * @return 文件格式
     */
    static FileFormat detectFormat(const void* data, size_t size);
    
    /**
     * @brief 🚀 检查文件是否为有效的Excel文件
     * @param file_path 文件路径
     * @return 是否有效
     */
    static bool isValidExcelFile(const std::string& file_path);

    // ==================== 便捷方法 ====================
    
    /**
     * @brief 🚀 快速读取CSV文件
     * @param file_path CSV文件路径
     * @param delimiter 分隔符（默认逗号）
     * @return 工作簿智能指针或错误
     */
    static TXResult<std::unique_ptr<TXWorkbook>> loadCSV(
        const std::string& file_path,
        char delimiter = ','
    );
    
    /**
     * @brief 🚀 快速保存为CSV文件
     * @param workbook 工作簿指针
     * @param file_path CSV文件路径
     * @param sheet_index 工作表索引（默认第一个）
     * @param delimiter 分隔符（默认逗号）
     * @return 操作结果
     */
    static TXResult<void> saveCSV(
        const TXWorkbook* workbook,
        const std::string& file_path,
        size_t sheet_index = 0,
        char delimiter = ','
    );

    // ==================== 信息获取 ====================
    
    /**
     * @brief 🚀 获取文件信息（不完全加载）
     * @param file_path 文件路径
     * @return 文件信息
     */
    struct FileInfo {
        std::string creator;                // 创建者
        std::string last_modified_by;       // 最后修改者
        std::string created_time;           // 创建时间
        std::string modified_time;          // 修改时间
        TXVector<std::string> sheet_names;  // 工作表名称列表
        size_t total_sheets;                // 工作表总数
        FileFormat format;                  // 文件格式
        size_t file_size;                   // 文件大小
        
        FileInfo(TXUnifiedMemoryManager& manager) 
            : sheet_names(manager), total_sheets(0), format(FileFormat::XLSX), file_size(0) {}
    };
    
    static TXResult<FileInfo> getFileInfo(const std::string& file_path);

    // ==================== 批量操作 ====================
    
    /**
     * @brief 🚀 批量转换文件格式
     * @param input_files 输入文件列表
     * @param output_dir 输出目录
     * @param target_format 目标格式
     * @return 转换结果
     */
    static TXResult<size_t> batchConvert(
        const TXVector<std::string>& input_files,
        const std::string& output_dir,
        FileFormat target_format
    );

private:
    // ==================== 内部实现类 ====================
    
    class XLSXReader;   // XLSX读取器
    class XLSXWriter;   // XLSX写入器
    class CSVReader;    // CSV读取器
    class CSVWriter;    // CSV写入器
    
    // ==================== 内部辅助方法 ====================
    
    /**
     * @brief 验证文件路径
     */
    static TXResult<void> validateFilePath(const std::string& file_path, bool for_writing = false);
    
    /**
     * @brief 验证工作簿
     */
    static TXResult<void> validateWorkbook(const TXWorkbook* workbook);
    
    /**
     * @brief 获取文件扩展名
     */
    static std::string getFileExtension(const std::string& file_path);
    
    /**
     * @brief 创建备份文件
     */
    static TXResult<void> createBackup(const std::string& file_path);
    
    /**
     * @brief 错误处理
     */
    static void handleError(const std::string& operation, const TXError& error);
};

/**
 * @brief 🚀 便捷的Excel文件加载函数
 */
inline TXResult<std::unique_ptr<TXWorkbook>> loadExcel(const std::string& file_path) {
    return TXExcelIO::loadFromFile(file_path);
}

/**
 * @brief 🚀 便捷的Excel文件保存函数
 */
inline TXResult<void> saveExcel(const TXWorkbook* workbook, const std::string& file_path) {
    return TXExcelIO::saveToFile(workbook, file_path);
}

} // namespace TinaXlsx
