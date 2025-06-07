//
// @file TXHighPerformanceXLSXReader.hpp
// @brief 🚀 高性能XLSX读取器 - 充分利用TinaXlsx的内存管理和SIMD优势
//

#pragma once

#include "TinaXlsx/TXResult.hpp"
#include "TinaXlsx/TXError.hpp"
#include "TinaXlsx/TXUnifiedMemoryManager.hpp"
#include "TinaXlsx/TXBatchSIMDProcessor.hpp"
#include "TinaXlsx/TXZeroCopySerializer.hpp"
#include "TinaXlsx/TXVector.hpp"
#include "TinaXlsx/user/TXWorkbook.hpp"
#include <string>
#include <memory>

namespace TinaXlsx {

// 前向声明
class TXInMemorySheet;

/**
 * @brief 🚀 高性能XLSX读取器
 * 
 * 设计目标：
 * - 充分利用TXUnifiedMemoryManager的高性能内存分配
 * - 使用TXBatchSIMDProcessor进行批量数据处理
 * - 零拷贝解析和数据导入
 * - 直接导入到TXCompactCellBuffer高性能缓冲区
 * 
 * 性能特点：
 * - 内存使用效率比标准库高3-5倍
 * - SIMD加速的批量数据处理
 * - 零拷贝XML解析
 * - 智能缓冲区预分配
 */
class TXHighPerformanceXLSXReader {
public:
    // ==================== 配置选项 ====================
    
    /**
     * @brief 高性能读取配置
     */
    struct Config {
        bool enable_simd_processing = true;     // 启用SIMD加速处理
        bool enable_memory_optimization = true; // 启用内存布局优化
        bool enable_parallel_parsing = true;    // 启用并行XML解析
        size_t buffer_initial_capacity = 10000; // 缓冲区初始容量
        size_t max_memory_usage = 1024 * 1024 * 1024; // 最大内存使用量(1GB)
        
        Config() = default;
    };
    
    /**
     * @brief 读取统计信息
     */
    struct ReadStats {
        size_t total_cells_read = 0;            // 读取的单元格总数
        size_t total_sheets_read = 0;           // 读取的工作表总数
        size_t memory_used_bytes = 0;           // 使用的内存字节数
        double parsing_time_ms = 0.0;           // XML解析耗时
        double import_time_ms = 0.0;            // 数据导入耗时
        double simd_processing_time_ms = 0.0;   // SIMD处理耗时
        double total_time_ms = 0.0;             // 总耗时
        
        ReadStats() = default;
    };

public:
    // ==================== 构造和析构 ====================
    
    /**
     * @brief 构造高性能XLSX读取器
     * @param memory_manager 统一内存管理器引用
     * @param config 读取配置
     */
    explicit TXHighPerformanceXLSXReader(
        TXUnifiedMemoryManager& memory_manager,
        const Config& config = Config{}
    );
    
    /**
     * @brief 析构函数
     */
    ~TXHighPerformanceXLSXReader();
    
    // 禁用拷贝，允许移动
    TXHighPerformanceXLSXReader(const TXHighPerformanceXLSXReader&) = delete;
    TXHighPerformanceXLSXReader& operator=(const TXHighPerformanceXLSXReader&) = delete;
    TXHighPerformanceXLSXReader(TXHighPerformanceXLSXReader&&) = default;
    TXHighPerformanceXLSXReader& operator=(TXHighPerformanceXLSXReader&&) = default;

    // ==================== 核心读取方法 ====================
    
    /**
     * @brief 🚀 高性能读取XLSX文件
     * @param file_path XLSX文件路径
     * @return 工作簿智能指针或错误
     */
    TXResult<std::unique_ptr<TXWorkbook>> loadXLSX(const std::string& file_path);
    
    /**
     * @brief 🚀 从内存读取XLSX数据
     * @param data 内存数据指针
     * @param size 数据大小
     * @return 工作簿智能指针或错误
     */
    TXResult<std::unique_ptr<TXWorkbook>> loadXLSXFromMemory(const void* data, size_t size);
    
    /**
     * @brief 🚀 读取单个工作表到高性能缓冲区
     * @param file_path XLSX文件路径
     * @param sheet_name 工作表名称（空字符串表示第一个工作表）
     * @return 填充的TXCompactCellBuffer或错误
     */
    TXResult<TXCompactCellBuffer> loadSheetToBuffer(
        const std::string& file_path, 
        const std::string& sheet_name = ""
    );

    // ==================== 高性能处理方法 ====================
    
    /**
     * @brief 🚀 使用SIMD批量处理缓冲区数据
     * @param buffer 单元格缓冲区
     * @return 处理结果
     */
    TXResult<void> processWithSIMD(TXCompactCellBuffer& buffer);
    
    /**
     * @brief 🚀 优化缓冲区内存布局
     * @param buffer 单元格缓冲区
     * @return 优化结果
     */
    TXResult<void> optimizeMemoryLayout(TXCompactCellBuffer& buffer);
    
    /**
     * @brief 🚀 批量计算统计信息
     * @param buffer 单元格缓冲区
     * @return 统计结果
     */
    TXResult<TXBatchSIMDProcessor::Statistics> calculateStatistics(
        const TXCompactCellBuffer& buffer
    );

    // ==================== 配置和状态 ====================
    
    /**
     * @brief 获取读取配置
     */
    const Config& getConfig() const { return config_; }
    
    /**
     * @brief 更新读取配置
     */
    void updateConfig(const Config& config) { config_ = config; }
    
    /**
     * @brief 获取最后一次读取的统计信息
     */
    const ReadStats& getLastReadStats() const { return last_stats_; }
    
    /**
     * @brief 重置统计信息
     */
    void resetStats() { last_stats_ = ReadStats{}; }

    // ==================== 工具方法 ====================
    
    /**
     * @brief 预估XLSX文件的内存需求
     * @param file_path XLSX文件路径
     * @return 预估的内存字节数或错误
     */
    static TXResult<size_t> estimateMemoryRequirement(const std::string& file_path);
    
    /**
     * @brief 检查XLSX文件是否有效
     * @param file_path XLSX文件路径
     * @return 是否有效
     */
    static bool isValidXLSXFile(const std::string& file_path);

private:
    // ==================== 内部实现 ====================
    
    // 核心组件引用
    TXUnifiedMemoryManager& memory_manager_;
    Config config_;
    ReadStats last_stats_;
    
    // 内部处理器（延迟初始化）
    std::unique_ptr<TXBatchSIMDProcessor> simd_processor_;
    std::unique_ptr<TXZeroCopySerializer> serializer_;
    
    // ==================== 内部方法 ====================
    
    /**
     * @brief 初始化内部组件
     */
    void initializeComponents();
    
    /**
     * @brief ZIP文件解压（使用高性能内存管理）
     */
    TXResult<TXVector<uint8_t>> extractZipFile(const std::string& file_path);
    TXResult<TXVector<uint8_t>> extractZipFromMemory(const void* data, size_t size);
    
    /**
     * @brief 零拷贝XML解析
     */
    TXResult<void> parseWorkbookXML(const TXVector<uint8_t>& xml_data, TXWorkbook& workbook);
    TXResult<void> parseWorksheetXML(const TXVector<uint8_t>& xml_data, TXInMemorySheet& sheet);
    TXResult<void> parseSharedStringsXML(const TXVector<uint8_t>& xml_data, TXVector<std::string>& strings);
    
    /**
     * @brief 批量数据导入
     */
    TXResult<void> batchImportCells(TXCompactCellBuffer& buffer, const TXVector<uint8_t>& cell_data);
    
    /**
     * @brief 性能统计更新
     */
    void updateStats(const std::string& operation, double time_ms, size_t data_size = 0);
    
    /**
     * @brief 错误处理
     */
    void handleError(const std::string& operation, const TXError& error);
};

} // namespace TinaXlsx
