//
// @file TXBatchSIMDProcessor.hpp
// @brief 批量SIMD处理器 - 实现2毫秒内存中Excel操作
//

#pragma once

#include "TXVariant.hpp"
#include "TXCoordinate.hpp"
#include "TXTypes.hpp"
#include "TXResult.hpp"
#include <vector>
#include <string>
#include <unordered_map>
#include <xsimd/xsimd.hpp>

namespace TinaXlsx {

// 前向声明
class TXGlobalStringPool;
class TXInMemorySheet;

/**
 * @brief 紧凑单元格缓冲区 - SoA(结构体数组)设计，SIMD友好
 */
struct TXCompactCellBuffer {
    // 核心数据 - 连续内存布局，SIMD优化
    std::vector<double> number_values;        // 数值数据 (8字节对齐)
    std::vector<uint32_t> string_indices;    // 字符串索引 (4字节)
    std::vector<uint16_t> style_indices;     // 样式索引 (2字节)
    std::vector<uint32_t> coordinates;       // 压缩坐标 (row << 16 | col)
    std::vector<uint8_t> cell_types;         // 单元格类型 (1字节)
    
    // 元数据
    size_t capacity = 0;                      // 容量
    size_t size = 0;                         // 当前大小
    bool is_sorted = false;                   // 是否按坐标排序
    
    // 构造函数
    TXCompactCellBuffer() = default;
    explicit TXCompactCellBuffer(size_t initial_capacity);
    
    // 内存管理
    void reserve(size_t new_capacity);
    void resize(size_t new_size);
    void clear();
    void shrink_to_fit();
    
    // 实用工具
    bool empty() const { return size == 0; }
    void sort_by_coordinates();               // 按坐标排序，优化访问模式
};

/**
 * @brief 单元格统计信息
 */
struct TXCellStats {
    size_t count = 0;
    double sum = 0.0;
    double mean = 0.0;
    double min_value = 0.0;
    double max_value = 0.0;
    double std_dev = 0.0;
    size_t number_cells = 0;
    size_t string_cells = 0;
    size_t empty_cells = 0;
};

/**
 * @brief 样式应用规则
 */
struct TXStyleRule {
    TXRange range;                           // 应用范围
    uint16_t style_index;                    // 样式索引
    bool overwrite = true;                   // 是否覆盖现有样式
};

/**
 * @brief 导入选项
 */
struct TXImportOptions {
    bool auto_detect_types = true;           // 自动检测数据类型
    bool enable_simd = true;                 // 启用SIMD优化
    bool optimize_memory = true;             // 优化内存布局
    size_t batch_size = 10000;              // 批处理大小
    bool skip_empty_cells = true;            // 跳过空单元格
};

/**
 * @brief 🚀 批量SIMD处理器 - 核心性能组件
 * 
 * 实现极致性能的批量操作：
 * - 批量单元格创建 (SIMD优化)
 * - 批量数据转换 (SIMD优化)
 * - 批量坐标处理 (SIMD优化)
 * - 批量统计计算 (SIMD优化)
 */
class TXBatchSIMDProcessor {
public:
    // ==================== 批量单元格创建 ====================
    
    /**
     * @brief 批量创建数值单元格 - 核心性能方法
     * @param values 数值数组 (必须16字节对齐以支持AVX)
     * @param buffer 输出缓冲区
     * @param coordinates 坐标数组
     * @param count 数量
     */
    static void batchCreateNumberCells(
        const double* values,
        TXCompactCellBuffer& buffer,
        const uint32_t* coordinates,
        size_t count
    );
    
    /**
     * @brief 批量创建字符串单元格
     * @param strings 字符串数组
     * @param buffer 输出缓冲区
     * @param coordinates 坐标数组
     * @param string_pool 全局字符串池
     */
    static void batchCreateStringCells(
        const std::vector<std::string>& strings,
        TXCompactCellBuffer& buffer,
        const uint32_t* coordinates,
        TXGlobalStringPool& string_pool
    );
    
    /**
     * @brief 混合批量创建 - 自动类型检测
     * @param variants 变长数据数组
     * @param buffer 输出缓冲区
     * @param coordinates 坐标数组
     * @param string_pool 字符串池
     */
    static void batchCreateMixedCells(
        const std::vector<TXVariant>& variants,
        TXCompactCellBuffer& buffer,
        const uint32_t* coordinates,
        TXGlobalStringPool& string_pool
    );
    
    // ==================== 批量数据转换 ====================
    
    /**
     * @brief 批量坐标转换 - A1标记转数值坐标
     * @param cell_refs 单元格引用 ("A1", "B2", ...)
     * @param coordinates 输出坐标数组
     * @param count 数量
     * @return 转换成功的数量
     */
    static size_t batchConvertCoordinates(
        const std::vector<std::string>& cell_refs,
        uint32_t* coordinates,
        size_t count
    );
    
    /**
     * @brief 批量数值转换 - 字符串到数值
     * @param string_values 字符串数值数组
     * @param output_values 输出数值数组
     * @param count 数量
     * @return 转换成功的数量
     */
    static size_t batchConvertStringsToNumbers(
        const std::vector<std::string>& string_values,
        double* output_values,
        size_t count
    );
    
    /**
     * @brief 批量类型检测
     * @param values 字符串值数组
     * @param types 输出类型数组
     * @param count 数量
     */
    static void batchDetectTypes(
        const std::vector<std::string>& values,
        uint8_t* types,
        size_t count
    );
    
    // ==================== 批量计算和统计 ====================
    
    /**
     * @brief 批量统计计算 - SIMD优化
     * @param buffer 单元格缓冲区
     * @param range 计算范围 (可选)
     * @return 统计结果
     */
    static TXCellStats batchCalculateStats(
        const TXCompactCellBuffer& buffer,
        const TXRange* range = nullptr
    );
    
    /**
     * @brief 批量求和 - SIMD优化
     * @param buffer 单元格缓冲区
     * @param range 求和范围
     * @return 求和结果
     */
    static double batchSum(
        const TXCompactCellBuffer& buffer,
        const TXRange& range
    );
    
    /**
     * @brief 批量查找 - SIMD优化
     * @param buffer 单元格缓冲区
     * @param target_value 目标值
     * @param results 输出结果坐标
     * @return 找到的数量
     */
    static size_t batchFind(
        const TXCompactCellBuffer& buffer,
        double target_value,
        std::vector<uint32_t>& results
    );
    
    // ==================== 批量样式操作 ====================
    
    /**
     * @brief 批量应用样式
     * @param buffer 单元格缓冲区
     * @param rules 样式规则数组
     */
    static void batchApplyStyles(
        TXCompactCellBuffer& buffer,
        const std::vector<TXStyleRule>& rules
    );
    
    /**
     * @brief 批量清除样式
     * @param buffer 单元格缓冲区
     * @param range 清除范围
     */
    static void batchClearStyles(
        TXCompactCellBuffer& buffer,
        const TXRange& range
    );
    
    // ==================== 批量范围操作 ====================
    
    /**
     * @brief 填充范围 - SIMD优化
     * @param buffer 单元格缓冲区
     * @param range 填充范围
     * @param value 填充值
     */
    static void fillRange(
        TXCompactCellBuffer& buffer,
        const TXRange& range,
        double value
    );
    
    /**
     * @brief 拷贝范围 - SIMD优化
     * @param buffer 单元格缓冲区
     * @param src_range 源范围
     * @param dst_start 目标起始坐标
     */
    static void copyRange(
        TXCompactCellBuffer& buffer,
        const TXRange& src_range,
        const TXCoordinate& dst_start
    );
    
    /**
     * @brief 清除范围
     * @param buffer 单元格缓冲区
     * @param range 清除范围
     */
    static void clearRange(
        TXCompactCellBuffer& buffer,
        const TXRange& range
    );
    
    // ==================== 性能优化工具 ====================
    
    /**
     * @brief 优化内存布局 - 提高缓存命中率
     * @param buffer 单元格缓冲区
     */
    static void optimizeMemoryLayout(TXCompactCellBuffer& buffer);
    
    /**
     * @brief 压缩稀疏数据 - 移除空白单元格
     * @param buffer 单元格缓冲区
     * @return 压缩后的大小
     */
    static size_t compressSparseData(TXCompactCellBuffer& buffer);
    
    /**
     * @brief 预热SIMD缓存
     * @param warmup_size 预热数据大小
     */
    static void warmupSIMD(size_t warmup_size = 10000);
    
    // ==================== 性能监控 ====================
    
    /**
     * @brief 批处理性能统计
     */
    struct BatchPerformanceStats {
        size_t total_operations = 0;         // 总操作数
        size_t total_cells_processed = 0;    // 总处理单元格数
        double total_time_ms = 0.0;          // 总时间(毫秒)
        double avg_throughput = 0.0;         // 平均吞吐量(单元格/秒)
        double simd_utilization = 0.0;       // SIMD利用率
        size_t cache_hits = 0;               // 缓存命中数
        size_t cache_misses = 0;             // 缓存未命中数
    };
    
    /**
     * @brief 获取性能统计
     */
    static const BatchPerformanceStats& getPerformanceStats();
    
    /**
     * @brief 重置性能统计
     */
    static void resetPerformanceStats();
    
private:
    // 内部SIMD实现
    static void batchCreateNumberCellsSIMD(
        const double* values,
        TXCompactCellBuffer& buffer,
        const uint32_t* coordinates,
        size_t count
    );
    
    static void batchCreateNumberCellsScalar(
        const double* values,
        TXCompactCellBuffer& buffer,
        const uint32_t* coordinates,
        size_t count
    );
    
    // SIMD工具函数
    static bool is_memory_aligned(const void* ptr, size_t alignment = 16);
    static void ensure_simd_alignment(std::vector<double>& vec);
    static size_t round_up_to_simd_size(size_t size);
    
    // 性能监控
    static BatchPerformanceStats performance_stats_;
    static void update_performance_stats(size_t cells_processed, double time_ms);
};

/**
 * @brief 🚀 高级批量操作 - 用户友好的API
 */
class TXBatchOperations {
public:
    /**
     * @brief 批量数据导入 - 1毫秒级性能
     * @param sheet 目标工作表
     * @param data 二维数据 (行x列)
     * @param start_coord 起始坐标 (默认A1)
     * @param options 导入选项
     * @return 导入结果
     */
    static TXResult<size_t> importDataBatch(
        TXInMemorySheet& sheet,
        const std::vector<std::vector<TXVariant>>& data,
        const TXCoordinate& start_coord = TXCoordinate(0, 0),
        const TXImportOptions& options = {}
    );
    
    /**
     * @brief 批量数值导入 - 极速版本
     * @param sheet 目标工作表
     * @param numbers 数值矩阵
     * @param start_coord 起始坐标
     * @return 导入的单元格数量
     */
    static TXResult<size_t> importNumbersBatch(
        TXInMemorySheet& sheet,
        const std::vector<std::vector<double>>& numbers,
        const TXCoordinate& start_coord = TXCoordinate(0, 0)
    );
    
    /**
     * @brief 批量字符串导入
     * @param sheet 目标工作表
     * @param strings 字符串矩阵
     * @param start_coord 起始坐标
     * @return 导入的单元格数量
     */
    static TXResult<size_t> importStringsBatch(
        TXInMemorySheet& sheet,
        const std::vector<std::vector<std::string>>& strings,
        const TXCoordinate& start_coord = TXCoordinate(0, 0)
    );
    
    /**
     * @brief 从CSV批量导入 - 文件到内存直接转换
     * @param sheet 目标工作表
     * @param csv_content CSV内容
     * @param options 导入选项
     * @return 导入的单元格数量
     */
    static TXResult<size_t> importFromCSV(
        TXInMemorySheet& sheet,
        const std::string& csv_content,
        const TXImportOptions& options = {}
    );
};

} // namespace TinaXlsx 