//
// @file TXZeroCopySerializer.hpp  
// @brief 零拷贝序列化器 - 实现极速XML生成和Excel序列化
//

#pragma once

#include "TXInMemorySheet.hpp"
#include "TXUnifiedMemoryManager.hpp"
#include "TXResult.hpp"
#include "TXXMLTemplates.hpp"
#include <vector>
#include <string>
#include <fmt/format.h>
#include <span>

namespace TinaXlsx {

/**
 * @brief XML模板 - 预编译的XML结构
 */
struct TXXMLTemplate {
    std::string header;                      // XML头部
    std::string footer;                      // XML尾部
    std::string row_start_template;          // 行开始模板
    std::string row_end_template;            // 行结束模板
    std::string cell_number_template;        // 数值单元格模板
    std::string cell_string_template;        // 字符串单元格模板
    std::string cell_formula_template;       // 公式单元格模板
    
    // 预编译优化
    bool is_compiled = false;
    size_t estimated_size_per_cell = 0;      // 每个单元格预估大小
};

/**
 * @brief 序列化选项
 */
struct TXSerializationOptions {
    bool enable_compression = true;          // 启用压缩
    bool stream_mode = true;                 // 流式模式
    bool optimize_memory = true;             // 优化内存使用
    size_t buffer_size = 1024 * 1024;      // 缓冲区大小 (1MB)
    bool enable_parallel = true;             // 启用并行处理
    size_t parallel_threshold = 10000;       // 并行处理阈值
};

/**
 * @brief 🚀 零拷贝序列化器 - 极速XML生成核心
 * 
 * 设计目标：
 * - 零拷贝：直接在内存中构建XML，无中间拷贝
 * - 模板化：预编译XML模板，运行时只填充数据
 * - 流式处理：边生成边压缩，不等待完整XML
 * - SIMD优化：批量字符串操作和数值转换
 */
class TXZeroCopySerializer {
private:
    TXUnifiedMemoryManager& memory_manager_; // 内存管理器
    std::vector<uint8_t> output_buffer_;     // 输出缓冲区
    size_t current_pos_ = 0;                 // 当前写入位置
    TXSerializationOptions options_;         // 序列化选项
    
    // XML模板缓存
    static TXXMLTemplate worksheet_template_;
    static TXXMLTemplate shared_strings_template_;
    static TXXMLTemplate workbook_template_;
    static bool templates_initialized_;
    
    // 性能统计
    struct {
        size_t total_cells_serialized = 0;
        size_t total_bytes_written = 0;
        double total_time_ms = 0.0;
        size_t template_cache_hits = 0;
        size_t compression_ratio_percent = 0;
    } stats_;

public:
    /**
     * @brief 构造函数
     * @param memory_manager 内存管理器引用
     * @param options 序列化选项
     */
    explicit TXZeroCopySerializer(
        TXUnifiedMemoryManager& memory_manager,
        const TXSerializationOptions& options = {}
    );
    
    /**
     * @brief 析构函数
     */
    ~TXZeroCopySerializer();
    
    // 禁用拷贝，支持移动
    TXZeroCopySerializer(const TXZeroCopySerializer&) = delete;
    TXZeroCopySerializer& operator=(const TXZeroCopySerializer&) = delete;
    TXZeroCopySerializer(TXZeroCopySerializer&&) noexcept;
    TXZeroCopySerializer& operator=(TXZeroCopySerializer&&) noexcept;
    
    // ==================== 核心序列化方法 ====================
    
    /**
     * @brief 序列化工作表 - 核心性能方法
     * @param sheet 内存中工作表
     * @return 序列化结果
     */
    TXResult<void> serializeWorksheet(const TXInMemorySheet& sheet);
    
    /**
     * @brief 序列化共享字符串表
     * @param string_pool 全局字符串池
     * @return 序列化结果
     */
    TXResult<void> serializeSharedStrings(const TXGlobalStringPool& string_pool);
    
    /**
     * @brief 序列化工作簿结构
     * @param sheets 工作表列表
     * @return 序列化结果
     */
    TXResult<void> serializeWorkbook(const std::vector<std::string>& sheet_names);
    
    /**
     * @brief 序列化样式表
     * @param styles 样式数据
     * @return 序列化结果
     */
    TXResult<void> serializeStyles(const std::vector<uint8_t>& styles);
    
    // ==================== 批量序列化 ====================
    
    /**
     * @brief 批量序列化单元格数据 - SIMD优化
     * @param buffer 单元格缓冲区
     * @param row_groups 行分组信息
     * @return 序列化的单元格数量
     */
    size_t serializeCellDataBatch(
        const TXCompactCellBuffer& buffer,
        const std::vector<TXRowGroup>& row_groups
    );
    
    /**
     * @brief 批量序列化行数据
     * @param buffer 单元格缓冲区
     * @param row_group 单行分组信息
     * @return 序列化的单元格数量
     */
    size_t serializeRowBatch(
        const TXCompactCellBuffer& buffer,
        const TXRowGroup& row_group
    );
    
    /**
     * @brief 并行序列化 - 多线程优化
     * @param buffer 单元格缓冲区
     * @param row_groups 行分组信息
     * @return 序列化结果
     */
    TXResult<void> serializeParallel(
        const TXCompactCellBuffer& buffer,
        const std::vector<TXRowGroup>& row_groups
    );
    
    // ==================== 高性能写入方法 ====================
    
    /**
     * @brief 预分配输出缓冲区
     * @param estimated_size 预估大小
     */
    void reserve(size_t estimated_size);
    
    /**
     * @brief 写入原始数据
     * @param data 数据指针
     * @param size 数据大小
     */
    void writeRaw(const void* data, size_t size);
    
    /**
     * @brief 写入字符串
     * @param str 字符串
     */
    void writeString(const std::string& str);
    
    /**
     * @brief 写入字符串视图 (零拷贝)
     * @param str 字符串视图
     */
    void writeStringView(std::string_view str);
    
    /**
     * @brief 批量写入字符串 - SIMD优化
     * @param strings 字符串数组
     */
    void writeStringsBatch(const std::vector<std::string>& strings);
    
    /**
     * @brief 应用XML模板
     * @param template_str 模板字符串
     * @param args 模板参数
     */
    template<typename... Args>
    void applyTemplate(const std::string& template_str, Args&&... args);
    
    // ==================== 快速单元格写入 ====================
    
    /**
     * @brief 写入数值单元格 - 极速版本
     * @param coord_str 坐标字符串 (如"A1")
     * @param value 数值
     */
    void writeNumberCell(const std::string& coord_str, double value);
    
    /**
     * @brief 写入字符串单元格 - 极速版本
     * @param coord_str 坐标字符串
     * @param value 字符串值
     */
    void writeStringCell(const std::string& coord_str, const std::string& value);
    
    /**
     * @brief 写入内联字符串单元格
     * @param coord_str 坐标字符串
     * @param value 字符串值
     */
    void writeInlineStringCell(const std::string& coord_str, const std::string& value);
    
    /**
     * @brief 批量写入数值单元格 - SIMD优化
     * @param coords 坐标字符串数组
     * @param values 数值数组
     * @param count 数量
     */
    void writeNumberCellsBatch(
        const std::vector<std::string>& coords,
        const double* values,
        size_t count
    );
    
    // ==================== XML结构写入 ====================
    
    /**
     * @brief 写入XML声明
     */
    void writeXMLDeclaration();
    
    /**
     * @brief 写入工作表开始标签
     */
    void writeWorksheetStart();
    
    /**
     * @brief 写入工作表结束标签
     */
    void writeWorksheetEnd();
    
    /**
     * @brief 写入行开始标签
     * @param row_index 行索引 (1-based)
     */
    void writeRowStart(uint32_t row_index);
    
    /**
     * @brief 写入行结束标签
     */
    void writeRowEnd();
    
    /**
     * @brief 写入sheetData开始标签
     */
    void writeSheetDataStart();
    
    /**
     * @brief 写入sheetData结束标签
     */
    void writeSheetDataEnd();
    
    // ==================== 性能优化 ====================
    
    /**
     * @brief 预估工作表序列化大小
     * @param sheet 工作表
     * @return 预估字节数
     */
    static size_t estimateWorksheetSize(const TXInMemorySheet& sheet);
    
    /**
     * @brief 预估单元格序列化大小
     * @param cell_count 单元格数量
     * @param avg_string_length 平均字符串长度
     * @return 预估字节数
     */
    static size_t estimateCellsSize(size_t cell_count, size_t avg_string_length = 10);
    
    /**
     * @brief 初始化XML模板 - 编译时优化
     */
    static void initializeTemplates();
    
    /**
     * @brief 优化输出缓冲区 - 内存对齐
     */
    void optimizeBuffer();
    
    /**
     * @brief 压缩输出数据
     * @return 压缩比例 (0.0-1.0)
     */
    double compressOutput();
    
    // ==================== 结果获取 ====================
    
    /**
     * @brief 获取序列化结果 (移动语义)
     * @return 序列化后的数据
     */
    std::vector<uint8_t> getResult() &&;
    
    /**
     * @brief 获取结果视图 (零拷贝)
     * @return 数据视图
     */
    std::vector<uint8_t> getResultView() const;
    
    /**
     * @brief 获取结果大小
     * @return 字节数
     */
    size_t getSize() const { return current_pos_; }
    
    /**
     * @brief 获取容量
     * @return 容量字节数
     */
    size_t getCapacity() const { return output_buffer_.capacity(); }
    
    /**
     * @brief 是否为空
     */
    bool empty() const { return current_pos_ == 0; }
    
    /**
     * @brief 清空内容
     */
    void clear();
    
    // ==================== 性能监控 ====================
    
    /**
     * @brief 序列化性能统计
     */
    struct SerializationStats {
        size_t total_cells;                  // 总单元格数
        size_t total_bytes;                  // 总字节数
        double serialization_time_ms;       // 序列化时间
        double throughput_cells_per_sec;     // 吞吐量(单元格/秒)
        double throughput_mb_per_sec;        // 吞吐量(MB/秒)
        size_t template_cache_hits;          // 模板缓存命中
        double compression_ratio;            // 压缩比
        size_t memory_usage_bytes;           // 内存使用量
    };
    
    /**
     * @brief 获取性能统计
     */
    SerializationStats getPerformanceStats() const;
    
    /**
     * @brief 重置性能统计
     */
    void resetPerformanceStats();

private:
    // 内部辅助方法
    void ensureCapacity(size_t additional_size);
    void appendToOutput(const void* data, size_t size);
    void appendToOutput(std::string_view str);
    
    // 坐标转换工具
    std::string coordToString(uint32_t coord) const;
    std::string rowColToString(uint32_t row, uint32_t col) const;
    
    // 数值转换工具 - SIMD优化
    void numberToString(double value, char* buffer, size_t buffer_size);
    void numberToStringBatch(const double* values, std::vector<std::string>& output, size_t count);
    
    // XML转义
    std::string escapeXMLString(const std::string& str);
    void escapeXMLStringInPlace(std::string& str);
    
    // 性能监控
    void updateStats(size_t cells_processed, size_t bytes_written, double time_ms);
};





} // namespace TinaXlsx
 