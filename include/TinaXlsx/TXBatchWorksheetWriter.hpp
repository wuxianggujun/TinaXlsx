//
// @file TXBatchWorksheetWriter.hpp
// @brief 批量工作表写入器 - 优化多工作表保存性能
//

#pragma once

#include <vector>
#include <memory>
#include <string>
#include <map>
#include "TXTypes.hpp"
#include "TXResult.hpp"

namespace TinaXlsx {

// 前向声明
struct TXWorkbookContext;
class TXZipArchiveWriter;

/**
 * @brief 🚀 批量工作表写入器
 * 
 * 优化多工作表保存性能：
 * - 并行XML生成
 * - 共享字符串池优化
 * - 批量ZIP写入
 */
class TXBatchWorksheetWriter {
public:
    /**
     * @brief 批量写入配置
     */
    struct BatchConfig {
        bool enableParallelGeneration = true;    // 启用并行生成
        bool enableSharedStringOptim = true;     // 启用共享字符串优化
        size_t maxConcurrentThreads = 4;         // 最大并发线程数
        size_t bufferSizePerSheet = 256 * 1024;  // 每个工作表缓冲区大小
    };
    
    explicit TXBatchWorksheetWriter(const BatchConfig& config = BatchConfig{});
    ~TXBatchWorksheetWriter();
    
    /**
     * @brief 批量保存所有工作表
     * @param zipWriter ZIP写入器
     * @param context 工作簿上下文
     * @return 保存结果
     */
    TXResult<void> saveAllWorksheets(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context);
    
    /**
     * @brief 获取批量写入统计信息
     */
    struct BatchStats {
        size_t totalSheets = 0;
        size_t totalCells = 0;
        double xmlGenerationTimeMs = 0.0;
        double zipWriteTimeMs = 0.0;
        double totalTimeMs = 0.0;
        size_t peakMemoryUsage = 0;
    };
    
    const BatchStats& getStats() const { return stats_; }

    /**
     * @brief 工作表写入任务（公开供其他类使用）
     */
    struct WorksheetTask {
        size_t sheetIndex;
        std::string partName;
        std::vector<uint8_t> xmlData;
        size_t cellCount;
        double generationTimeMs;
        
        WorksheetTask(size_t index) : sheetIndex(index), cellCount(0), generationTimeMs(0.0) {}
    };

private:
    BatchConfig config_;
    BatchStats stats_;

    /**
     * @brief 并行生成工作表XML
     * @param context 工作簿上下文
     * @return 生成的任务列表
     */
    std::vector<WorksheetTask> generateWorksheetsParallel(const TXWorkbookContext& context);
    
    /**
     * @brief 串行生成工作表XML（回退方案）
     * @param context 工作簿上下文
     * @return 生成的任务列表
     */
    std::vector<WorksheetTask> generateWorksheetsSerial(const TXWorkbookContext& context);
    
    /**
     * @brief 生成单个工作表XML
     * @param sheetIndex 工作表索引
     * @param context 工作簿上下文
     * @return 工作表任务
     */
    WorksheetTask generateSingleWorksheet(size_t sheetIndex, const TXWorkbookContext& context);
    
    /**
     * @brief 批量写入到ZIP
     * @param zipWriter ZIP写入器
     * @param tasks 工作表任务列表
     * @return 写入结果
     */
    TXResult<void> batchWriteToZip(TXZipArchiveWriter& zipWriter, const std::vector<WorksheetTask>& tasks);
    
    /**
     * @brief 优化共享字符串池
     * @param context 工作簿上下文
     */
    void optimizeSharedStrings(const TXWorkbookContext& context);
    
    /**
     * @brief 预分析工作表数据
     * @param context 工作簿上下文
     * @return 分析结果
     */
    struct WorksheetAnalysis {
        std::vector<size_t> cellCounts;
        std::vector<bool> hasLargeData;
        size_t totalEstimatedSize;
    };
    
    WorksheetAnalysis analyzeWorksheets(const TXWorkbookContext& context);
};

/**
 * @brief 🚀 高性能工作表XML生成器
 * 
 * 专门用于批量生成，避免重复初始化开销
 */
class TXFastWorksheetXmlGenerator {
public:
    explicit TXFastWorksheetXmlGenerator(size_t bufferSize = 256 * 1024);
    
    /**
     * @brief 生成工作表XML
     * @param sheetIndex 工作表索引
     * @param context 工作簿上下文
     * @return XML数据
     */
    TXResult<std::vector<uint8_t>> generateWorksheetXml(size_t sheetIndex, const TXWorkbookContext& context);
    
    /**
     * @brief 重置生成器（复用缓冲区）
     */
    void reset();

private:
    std::unique_ptr<class TXPugiWorksheetWriter> writer_;
    size_t bufferSize_;
    
    /**
     * @brief 快速构建单元格XML
     */
    void buildCellXml(const class TXCompactCell* cell, const std::string& cellRef, 
                     const TXWorkbookContext& context);
    
    /**
     * @brief 快速构建行XML
     */
    void buildRowXml(u32 rowNumber, const std::vector<class TXCompactCell*>& cells, 
                    const TXWorkbookContext& context);
};

/**
 * @brief 🚀 并行任务管理器
 * 
 * 管理工作表生成的并行任务
 */
class TXParallelTaskManager {
public:
    explicit TXParallelTaskManager(size_t maxThreads = 4);
    ~TXParallelTaskManager();
    
    /**
     * @brief 并行执行工作表生成任务
     * @param context 工作簿上下文
     * @param sheetIndices 要生成的工作表索引列表
     * @return 生成结果
     */
    std::vector<TXBatchWorksheetWriter::WorksheetTask> executeParallel(
        const TXWorkbookContext& context, 
        const std::vector<size_t>& sheetIndices);

private:
    size_t maxThreads_;
    std::vector<std::unique_ptr<TXFastWorksheetXmlGenerator>> generators_;
    
    /**
     * @brief 工作线程函数
     */
    void workerThread(const TXWorkbookContext& context, 
                     const std::vector<size_t>& taskIndices,
                     std::vector<TXBatchWorksheetWriter::WorksheetTask>& results,
                     size_t threadId);
};

/**
 * @brief 🚀 内存优化的ZIP批量写入器
 * 
 * 优化大量文件的ZIP写入性能
 */
class TXOptimizedZipBatchWriter {
public:
    explicit TXOptimizedZipBatchWriter(TXZipArchiveWriter& zipWriter);
    
    /**
     * @brief 批量写入文件
     * @param files 文件列表（路径和数据）
     * @return 写入结果
     */
    TXResult<void> batchWrite(const std::vector<std::pair<std::string, std::vector<uint8_t>>>& files);
    
    /**
     * @brief 设置批量写入选项
     */
    struct BatchWriteOptions {
        bool enableCompression = true;      // 启用压缩
        bool enableBuffering = true;        // 启用缓冲
        size_t bufferSize = 1024 * 1024;    // 缓冲区大小
        bool enableParallelCompression = true; // 并行压缩
    };
    
    void setOptions(const BatchWriteOptions& options) { options_ = options; }

private:
    TXZipArchiveWriter& zipWriter_;
    BatchWriteOptions options_;
    
    /**
     * @brief 并行压缩文件
     */
    std::vector<std::vector<uint8_t>> compressFilesParallel(
        const std::vector<std::pair<std::string, std::vector<uint8_t>>>& files);
};

} // namespace TinaXlsx
