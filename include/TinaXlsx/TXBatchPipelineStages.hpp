//
// @file TXBatchPipelineStages.hpp
// @brief 批处理流水线的4个默认阶段实现
//

#pragma once

#include "TXBatchPipeline.hpp"
#include "TXUnifiedMemoryManager.hpp"
#include "TXCompactCell.hpp"
#include <sstream>
#include <algorithm>
#include <map>
#include <fstream>

namespace TinaXlsx {

/**
 * @brief 阶段1：数据预处理
 * 
 * 功能：
 * - 数据验证和清理
 * - 内存优化分配
 * - 数据格式标准化
 * - 批次大小优化
 */
class TXDataPreprocessingStage : public TXPipelineStage {
public:
    struct PreprocessingConfig {
        size_t max_batch_size = 10000;           // 最大批次大小
        size_t min_batch_size = 100;             // 最小批次大小
        bool enable_data_validation = true;      // 启用数据验证
        bool enable_memory_optimization = true;  // 启用内存优化
        bool enable_deduplication = true;        // 启用去重
    };
    
    explicit TXDataPreprocessingStage(TXUnifiedMemoryManager& memory_manager, 
                                     const PreprocessingConfig& config = PreprocessingConfig{});
    
    TXResult<std::unique_ptr<TXBatchData>> process(std::unique_ptr<TXBatchData> input) override;
    std::string getStageName() const override { return "DataPreprocessing"; }
    StageStats getStats() const override;
    void resetStats() override;

private:
    TXUnifiedMemoryManager& memory_manager_;
    PreprocessingConfig config_;
    mutable StageStats stats_;
    mutable std::mutex stats_mutex_;
    
    // 内部处理方法
    TXResult<void> validateData(TXBatchData& batch);
    TXResult<void> optimizeMemoryLayout(TXBatchData& batch);
    TXResult<void> deduplicateStrings(TXBatchData& batch);
    size_t estimateBatchSize(const TXBatchData& batch);
};

/**
 * @brief 阶段2：XML批量生成
 * 
 * 功能：
 * - 高效XML生成
 * - 批量字符串处理
 * - 内存池优化
 * - 流式输出准备
 */
class TXXMLGenerationStage : public TXPipelineStage {
public:
    struct XMLConfig {
        bool enable_streaming = true;            // 启用流式生成
        bool enable_compression_hints = true;    // 启用压缩提示
        bool enable_memory_pooling = true;       // 启用内存池
        size_t xml_buffer_size = 64 * 1024;     // XML缓冲区大小64KB
        std::string xml_encoding = "UTF-8";      // XML编码
    };
    
    explicit TXXMLGenerationStage(TXUnifiedMemoryManager& memory_manager,
                                 const XMLConfig& config = XMLConfig{});
    
    TXResult<std::unique_ptr<TXBatchData>> process(std::unique_ptr<TXBatchData> input) override;
    std::string getStageName() const override { return "XMLGeneration"; }
    StageStats getStats() const override;
    void resetStats() override;

private:
    TXUnifiedMemoryManager& memory_manager_;
    XMLConfig config_;
    mutable StageStats stats_;
    mutable std::mutex stats_mutex_;
    
    // XML生成缓冲区
    thread_local static std::unique_ptr<std::ostringstream> xml_buffer_;
    
    // 内部处理方法
    TXResult<std::string> generateCellXML(const TXCompactCell& cell);
    TXResult<std::string> generateBatchXML(const TXBatchData& batch);
    TXResult<void> optimizeXMLOutput(std::string& xml);
    void prepareCompressionHints(TXBatchData& batch);
};

/**
 * @brief 阶段3：压缩处理
 * 
 * 功能：
 * - 数据压缩
 * - 大块内存优化
 * - 压缩算法选择
 * - 压缩率统计
 */
class TXCompressionStage : public TXPipelineStage {
public:
    enum class CompressionAlgorithm {
        NONE,
        ZLIB,
        LZ4,
        ZSTD
    };
    
    struct CompressionConfig {
        CompressionAlgorithm algorithm = CompressionAlgorithm::ZLIB;
        int compression_level = 6;               // 压缩级别(1-9)
        size_t compression_threshold = 1024;     // 压缩阈值1KB
        bool enable_adaptive_compression = true; // 自适应压缩
        bool enable_parallel_compression = true; // 并行压缩
    };
    
    explicit TXCompressionStage(TXUnifiedMemoryManager& memory_manager,
                               const CompressionConfig& config = CompressionConfig{});
    
    TXResult<std::unique_ptr<TXBatchData>> process(std::unique_ptr<TXBatchData> input) override;
    std::string getStageName() const override { return "Compression"; }
    StageStats getStats() const override;
    void resetStats() override;

private:
    TXUnifiedMemoryManager& memory_manager_;
    CompressionConfig config_;
    mutable StageStats stats_;
    mutable std::mutex stats_mutex_;
    
    // 压缩统计
    mutable size_t total_uncompressed_size_ = 0;
    mutable size_t total_compressed_size_ = 0;
    
    // 内部处理方法
    TXResult<std::vector<uint8_t>> compressData(const std::vector<uint8_t>& data);
    TXResult<std::vector<uint8_t>> compressString(const std::string& str);
    CompressionAlgorithm selectOptimalAlgorithm(const TXBatchData& batch);
    double calculateCompressionRatio() const;
};

/**
 * @brief 阶段4：输出写入
 * 
 * 功能：
 * - 异步文件写入
 * - 内存回收
 * - 写入优化
 * - 完成通知
 */
class TXOutputWriteStage : public TXPipelineStage {
public:
    struct OutputConfig {
        std::string output_directory = "./output";
        std::string file_prefix = "batch_";
        std::string file_extension = ".xlsx";
        bool enable_async_write = true;          // 异步写入
        bool enable_write_verification = true;   // 写入验证
        bool enable_memory_cleanup = true;       // 内存清理
        size_t write_buffer_size = 256 * 1024;  // 写入缓冲区256KB
    };
    
    explicit TXOutputWriteStage(TXUnifiedMemoryManager& memory_manager,
                               const OutputConfig& config = OutputConfig{});
    
    TXResult<std::unique_ptr<TXBatchData>> process(std::unique_ptr<TXBatchData> input) override;
    std::string getStageName() const override { return "OutputWrite"; }
    StageStats getStats() const override;
    void resetStats() override;

private:
    TXUnifiedMemoryManager& memory_manager_;
    OutputConfig config_;
    mutable StageStats stats_;
    mutable std::mutex stats_mutex_;
    
    // 写入统计
    mutable size_t total_bytes_written_ = 0;
    mutable size_t total_files_written_ = 0;
    
    // 内部处理方法
    TXResult<std::string> generateOutputPath(const TXBatchData& batch);
    TXResult<void> writeToFile(const std::string& path, const TXBatchData& batch);
    TXResult<void> verifyWrite(const std::string& path, size_t expected_size);
    void cleanupBatchMemory(TXBatchData& batch);
};

/**
 * @brief 批处理阶段工厂
 */
class TXPipelineStageFactory {
public:
    /**
     * @brief 创建默认的4级流水线阶段
     */
    static std::vector<std::unique_ptr<TXPipelineStage>> createDefaultStages(
        TXUnifiedMemoryManager& memory_manager);
    
    /**
     * @brief 创建自定义阶段
     */
    template<typename StageType, typename... Args>
    static std::unique_ptr<TXPipelineStage> createStage(Args&&... args) {
        return std::make_unique<StageType>(std::forward<Args>(args)...);
    }
    
    /**
     * @brief 创建数据预处理阶段
     */
    static std::unique_ptr<TXPipelineStage> createPreprocessingStage(
        TXUnifiedMemoryManager& memory_manager,
        const TXDataPreprocessingStage::PreprocessingConfig& config = {});
    
    /**
     * @brief 创建XML生成阶段
     */
    static std::unique_ptr<TXPipelineStage> createXMLGenerationStage(
        TXUnifiedMemoryManager& memory_manager,
        const TXXMLGenerationStage::XMLConfig& config = {});
    
    /**
     * @brief 创建压缩阶段
     */
    static std::unique_ptr<TXPipelineStage> createCompressionStage(
        TXUnifiedMemoryManager& memory_manager,
        const TXCompressionStage::CompressionConfig& config = {});
    
    /**
     * @brief 创建输出写入阶段
     */
    static std::unique_ptr<TXPipelineStage> createOutputWriteStage(
        TXUnifiedMemoryManager& memory_manager,
        const TXOutputWriteStage::OutputConfig& config = {});
};

/**
 * @brief 批处理性能分析器
 */
class TXBatchPerformanceAnalyzer {
public:
    /**
     * @brief 性能分析报告
     */
    struct PerformanceReport {
        // 整体性能
        double overall_throughput = 0.0;         // 整体吞吐量
        std::chrono::microseconds avg_latency{0}; // 平均延迟
        double memory_efficiency = 0.0;          // 内存效率
        
        // 各阶段性能
        std::map<std::string, double> stage_throughputs;
        std::map<std::string, std::chrono::microseconds> stage_latencies;
        std::map<std::string, double> stage_cpu_usage;
        
        // 瓶颈分析
        std::string bottleneck_stage;
        std::vector<std::string> optimization_suggestions;
        
        // 资源使用
        size_t peak_memory_usage = 0;
        double avg_cpu_usage = 0.0;
        double io_wait_ratio = 0.0;
    };
    
    /**
     * @brief 分析流水线性能
     */
    static PerformanceReport analyzePipeline(const TXBatchPipeline& pipeline);
    
    /**
     * @brief 生成优化建议
     */
    static std::vector<std::string> generateOptimizationSuggestions(const PerformanceReport& report);
    
    /**
     * @brief 检测性能瓶颈
     */
    static std::string detectBottleneck(const PerformanceReport& report);
};

} // namespace TinaXlsx
