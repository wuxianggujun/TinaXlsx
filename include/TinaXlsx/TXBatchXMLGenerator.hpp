//
// @file TXBatchXMLGenerator.hpp
// @brief 高效批量XML生成器 - 内存优化的XML批量生成
//

#pragma once

#include "TXUnifiedMemoryManager.hpp"
#include "TXCompactCell.hpp"
#include "TXTypes.hpp"
#include "TXResult.hpp"
#include <string>
#include <vector>
#include <memory>
#include <sstream>
#include <unordered_map>
#include <atomic>
#include <mutex>

namespace TinaXlsx {

/**
 * @brief XML生成模板
 */
struct TXXMLTemplate {
    std::string header;
    std::string footer;
    std::string row_template;
    std::string cell_template;
    std::unordered_map<std::string, std::string> placeholders;
};

/**
 * @brief 批量XML生成器
 */
class TXBatchXMLGenerator {
public:
    /**
     * @brief XML生成配置
     */
    struct XMLGeneratorConfig {
        // 内存管理
        bool enable_memory_pooling = true;       // 启用内存池
        size_t initial_buffer_size = 64 * 1024;  // 初始缓冲区64KB
        size_t max_buffer_size = 16 * 1024 * 1024; // 最大缓冲区16MB
        
        // 性能优化
        bool enable_string_interning = true;     // 启用字符串内化
        bool enable_template_caching = true;     // 启用模板缓存
        bool enable_parallel_generation = true;  // 启用并行生成
        size_t parallel_threshold = 1000;        // 并行处理阈值
        
        // XML格式
        std::string encoding = "UTF-8";
        bool pretty_print = false;               // 美化输出
        bool include_xml_declaration = true;     // 包含XML声明
        size_t indent_size = 2;                  // 缩进大小
        
        // 批处理
        size_t batch_size = 10000;              // 批处理大小
        bool enable_streaming = true;            // 启用流式生成
        bool enable_compression_hints = true;    // 启用压缩提示
    };
    
    /**
     * @brief XML生成统计
     */
    struct XMLGeneratorStats {
        size_t total_xml_generated = 0;          // 总生成XML数
        size_t total_cells_processed = 0;        // 总处理单元格数
        size_t total_bytes_generated = 0;        // 总生成字节数
        
        std::chrono::microseconds total_generation_time{0}; // 总生成时间
        std::chrono::microseconds avg_generation_time{0};   // 平均生成时间
        double generation_rate = 0.0;            // 生成速率(单元格/秒)
        
        size_t memory_usage = 0;                 // 内存使用量
        size_t peak_memory_usage = 0;            // 峰值内存使用
        double memory_efficiency = 0.0;          // 内存效率
        
        size_t template_cache_hits = 0;          // 模板缓存命中
        size_t template_cache_misses = 0;        // 模板缓存未命中
        size_t string_intern_hits = 0;           // 字符串内化命中
        
        double compression_ratio = 0.0;          // 压缩比估算
    };
    
    explicit TXBatchXMLGenerator(TXUnifiedMemoryManager& memory_manager,
                                const XMLGeneratorConfig& config = XMLGeneratorConfig{});
    ~TXBatchXMLGenerator();
    
    // ==================== XML生成接口 ====================
    
    /**
     * @brief 生成单个单元格XML
     */
    TXResult<std::string> generateCellXML(const TXCompactCell& cell);
    
    /**
     * @brief 批量生成单元格XML
     */
    TXResult<std::string> generateCellsXML(const std::vector<TXCompactCell>& cells);
    
    /**
     * @brief 生成行XML
     */
    TXResult<std::string> generateRowXML(size_t row_index, const std::vector<TXCompactCell>& cells);
    
    /**
     * @brief 批量生成行XML
     */
    TXResult<std::string> generateRowsXML(const std::vector<std::pair<size_t, std::vector<TXCompactCell>>>& rows);
    
    /**
     * @brief 生成工作表XML
     */
    TXResult<std::string> generateWorksheetXML(const std::string& sheet_name,
                                              const std::vector<std::pair<size_t, std::vector<TXCompactCell>>>& rows);
    
    /**
     * @brief 流式生成XML
     */
    class TXXMLStream {
    public:
        TXXMLStream(TXBatchXMLGenerator& generator, const std::string& root_element);
        ~TXXMLStream();
        
        TXResult<void> writeElement(const std::string& element, const std::string& content);
        TXResult<void> writeCell(const TXCompactCell& cell);
        TXResult<void> writeRow(size_t row_index, const std::vector<TXCompactCell>& cells);
        TXResult<std::string> finalize();
        
    private:
        TXBatchXMLGenerator& generator_;
        std::unique_ptr<std::ostringstream> stream_;
        std::string root_element_;
        bool finalized_ = false;
    };
    
    /**
     * @brief 创建XML流
     */
    std::unique_ptr<TXXMLStream> createXMLStream(const std::string& root_element);
    
    // ==================== 模板管理 ====================
    
    /**
     * @brief 设置XML模板
     */
    void setTemplate(const std::string& template_name, const TXXMLTemplate& xml_template);
    
    /**
     * @brief 获取XML模板
     */
    const TXXMLTemplate* getTemplate(const std::string& template_name) const;
    
    /**
     * @brief 加载默认模板
     */
    void loadDefaultTemplates();
    
    /**
     * @brief 清除模板缓存
     */
    void clearTemplateCache();
    
    // ==================== 性能优化 ====================
    
    /**
     * @brief 预热生成器
     */
    void warmup(size_t warmup_iterations = 1000);
    
    /**
     * @brief 优化内存使用
     */
    void optimizeMemory();
    
    /**
     * @brief 压缩内部缓存
     */
    size_t compactCache();
    
    /**
     * @brief 设置并行处理
     */
    void setParallelProcessing(bool enable, size_t thread_count = 0);
    
    // ==================== 统计和监控 ====================
    
    /**
     * @brief 获取统计信息
     */
    XMLGeneratorStats getStats() const;
    
    /**
     * @brief 重置统计信息
     */
    void resetStats();
    
    /**
     * @brief 生成性能报告
     */
    std::string generatePerformanceReport() const;
    
    /**
     * @brief 获取内存使用情况
     */
    size_t getCurrentMemoryUsage() const;
    
    // ==================== 配置管理 ====================
    
    /**
     * @brief 更新配置
     */
    void updateConfig(const XMLGeneratorConfig& config);
    
    /**
     * @brief 获取配置
     */
    const XMLGeneratorConfig& getConfig() const { return config_; }

private:
    // ==================== 内部数据结构 ====================
    
    TXUnifiedMemoryManager& memory_manager_;
    XMLGeneratorConfig config_;
    
    // 模板缓存
    std::unordered_map<std::string, TXXMLTemplate> templates_;
    mutable std::mutex templates_mutex_;
    
    // 字符串内化
    std::unordered_map<std::string, std::string> interned_strings_;
    mutable std::mutex strings_mutex_;
    
    // 内存缓冲区
    thread_local static std::unique_ptr<std::ostringstream> xml_buffer_;
    thread_local static std::unique_ptr<std::string> temp_string_;
    
    // 统计信息
    mutable XMLGeneratorStats stats_;
    mutable std::mutex stats_mutex_;
    
    // 并行处理
    bool parallel_enabled_ = false;
    size_t parallel_thread_count_ = 0;
    
    // ==================== 内部方法 ====================
    
    /**
     * @brief 获取或创建缓冲区
     */
    std::ostringstream& getXMLBuffer();
    
    /**
     * @brief 获取或创建临时字符串
     */
    std::string& getTempString();
    
    /**
     * @brief 内化字符串
     */
    const std::string& internString(const std::string& str);
    
    /**
     * @brief 应用模板
     */
    TXResult<std::string> applyTemplate(const std::string& template_name,
                                       const std::unordered_map<std::string, std::string>& values);
    
    /**
     * @brief 转义XML字符
     */
    std::string escapeXML(const std::string& str);
    
    /**
     * @brief 格式化单元格值
     */
    std::string formatCellValue(const TXCompactCell& cell);
    
    /**
     * @brief 生成单元格属性
     */
    std::string generateCellAttributes(const TXCompactCell& cell);
    
    /**
     * @brief 并行生成XML
     */
    TXResult<std::string> generateXMLParallel(const std::vector<TXCompactCell>& cells);
    
    /**
     * @brief 串行生成XML
     */
    TXResult<std::string> generateXMLSerial(const std::vector<TXCompactCell>& cells);
    
    /**
     * @brief 更新统计信息
     */
    void updateStats(size_t cells_processed, size_t bytes_generated, 
                    std::chrono::microseconds generation_time);
    
    /**
     * @brief 估算压缩比
     */
    double estimateCompressionRatio(const std::string& xml) const;
    
    /**
     * @brief 优化XML输出
     */
    void optimizeXMLOutput(std::string& xml);
};

/**
 * @brief XML生成器工厂
 */
class TXXMLGeneratorFactory {
public:
    /**
     * @brief 创建工作表XML生成器
     */
    static std::unique_ptr<TXBatchXMLGenerator> createWorksheetGenerator(
        TXUnifiedMemoryManager& memory_manager);
    
    /**
     * @brief 创建共享字符串XML生成器
     */
    static std::unique_ptr<TXBatchXMLGenerator> createSharedStringsGenerator(
        TXUnifiedMemoryManager& memory_manager);
    
    /**
     * @brief 创建样式XML生成器
     */
    static std::unique_ptr<TXBatchXMLGenerator> createStylesGenerator(
        TXUnifiedMemoryManager& memory_manager);
    
    /**
     * @brief 创建自定义XML生成器
     */
    static std::unique_ptr<TXBatchXMLGenerator> createCustomGenerator(
        TXUnifiedMemoryManager& memory_manager,
        const TXBatchXMLGenerator::XMLGeneratorConfig& config);
};

} // namespace TinaXlsx
