//
// @file TXInMemorySheet.hpp
// @brief 内存优先工作表 - 完全内存中操作，极致性能
//

#pragma once

#include "TXBatchSIMDProcessor.hpp"
#include "TXUnifiedMemoryManager.hpp"
#include "TXGlobalStringPool.hpp"
#include "TXCoordinate.hpp"
#include "TXRange.hpp"
#include "TXTypes.hpp"
#include "TXResult.hpp"
#include <unordered_map>
#include <memory>
#include <string>

namespace TinaXlsx {

// 前向声明
class TXZeroCopySerializer;
class TXInMemoryWorkbook;
class TXZipArchiveWriter;

/**
 * @brief 行分组信息 - 优化序列化性能
 */
struct TXRowGroup {
    uint32_t row_index;                      // 行索引
    size_t start_cell_index;                 // 该行第一个单元格在buffer中的索引
    size_t cell_count;                       // 该行单元格数量
};

/**
 * @brief 内存布局优化器
 */
class TXMemoryLayoutOptimizer {
public:
    /**
     * @brief 重新排列单元格以提高缓存命中率
     * @param buffer 单元格缓冲区
     */
    static void optimizeForSequentialAccess(TXCompactCellBuffer& buffer);
    
    /**
     * @brief 优化Excel访问模式 (按行排列)
     * @param buffer 单元格缓冲区
     */
    static void optimizeForExcelAccess(TXCompactCellBuffer& buffer);
    
    /**
     * @brief SIMD友好的内存对齐
     * @param buffer 单元格缓冲区
     */
    static void optimizeForSIMD(TXCompactCellBuffer& buffer);
    
    /**
     * @brief 生成行分组信息
     * @param buffer 单元格缓冲区
     * @return 行分组列表
     */
    static std::vector<TXRowGroup> generateRowGroups(const TXCompactCellBuffer& buffer);
};

/**
 * @brief 🚀 内存优先工作表 - 核心高性能组件
 * 
 * 设计理念：
 * - 完全内存中操作，最后一次性序列化
 * - SIMD批量处理，极致性能
 * - 零拷贝设计，最小内存开销
 * - 智能内存布局，优化缓存命中
 */
class TXInMemorySheet {
private:
    // 核心数据存储
    TXCompactCellBuffer cell_buffer_;        // 单元格缓冲区
    TXUnifiedMemoryManager& memory_manager_; // 内存管理器
    TXGlobalStringPool& string_pool_;        // 全局字符串池
    
    // 性能优化组件
    std::unique_ptr<TXMemoryLayoutOptimizer> optimizer_;
    
    // 快速索引 (坐标 -> 缓冲区索引)
    std::unordered_map<uint32_t, size_t> coord_to_index_;
    
    // 工作表元数据
    std::string name_;                       // 工作表名称
    uint32_t max_row_ = 0;                  // 最大行号
    uint32_t max_col_ = 0;                  // 最大列号
    bool dirty_ = false;                     // 是否有未保存的更改
    bool auto_optimize_ = true;              // 自动优化
    
    // 批处理配置
    static constexpr size_t DEFAULT_BATCH_SIZE = 10000;
    static constexpr size_t OPTIMIZATION_THRESHOLD = 1000;
    
    // 性能统计
    mutable struct {
        size_t total_cells = 0;
        size_t batch_operations = 0;
        double total_operation_time = 0.0;
        size_t cache_hits = 0;
        size_t cache_misses = 0;
    } stats_;

public:
    /**
     * @brief 构造函数
     * @param name 工作表名称
     * @param memory_manager 内存管理器
     * @param string_pool 字符串池
     */
    explicit TXInMemorySheet(
        const std::string& name,
        TXUnifiedMemoryManager& memory_manager,
        TXGlobalStringPool& string_pool
    );
    
    /**
     * @brief 析构函数
     */
    ~TXInMemorySheet();
    
    // 禁用拷贝，支持移动
    TXInMemorySheet(const TXInMemorySheet&) = delete;
    TXInMemorySheet& operator=(const TXInMemorySheet&) = delete;
    TXInMemorySheet(TXInMemorySheet&&) noexcept;
    TXInMemorySheet& operator=(TXInMemorySheet&&) noexcept;
    
    // ==================== 批量操作接口 (核心性能) ====================
    
    /**
     * @brief 批量设置数值 - 核心性能方法
     * @param coords 坐标数组
     * @param values 数值数组
     * @return 设置成功的单元格数量
     */
    TXResult<size_t> setBatchNumbers(
        const std::vector<TXCoordinate>& coords, 
        const std::vector<double>& values
    );
    
    /**
     * @brief 批量设置字符串
     * @param coords 坐标数组
     * @param values 字符串数组
     * @return 设置成功的单元格数量
     */
    TXResult<size_t> setBatchStrings(
        const std::vector<TXCoordinate>& coords, 
        const std::vector<std::string>& values
    );
    
    /**
     * @brief 批量设置样式
     * @param coords 坐标数组
     * @param style_indices 样式索引数组
     * @return 设置成功的单元格数量
     */
    TXResult<size_t> setBatchStyles(
        const std::vector<TXCoordinate>& coords, 
        const std::vector<uint16_t>& style_indices
    );
    
    /**
     * @brief 批量设置混合数据 - 自动类型检测
     * @param coords 坐标数组
     * @param variants 变长数据数组
     * @return 设置成功的单元格数量
     */
    TXResult<size_t> setBatchMixed(
        const std::vector<TXCoordinate>& coords,
        const std::vector<TXVariant>& variants
    );
    
    // ==================== SIMD优化的范围操作 ====================
    
    /**
     * @brief 填充范围 - SIMD优化
     * @param range 填充范围
     * @param value 填充值
     * @return 填充的单元格数量
     */
    TXResult<size_t> fillRange(const TXRange& range, double value);
    
    /**
     * @brief 填充范围 - 字符串版本
     * @param range 填充范围
     * @param value 填充字符串
     * @return 填充的单元格数量
     */
    TXResult<size_t> fillRange(const TXRange& range, const std::string& value);
    
    /**
     * @brief 复制范围 - SIMD优化
     * @param src_range 源范围
     * @param dst_start 目标起始坐标
     * @return 复制的单元格数量
     */
    TXResult<size_t> copyRange(const TXRange& src_range, const TXCoordinate& dst_start);
    
    /**
     * @brief 清除范围
     * @param range 清除范围
     * @return 清除的单元格数量
     */
    TXResult<size_t> clearRange(const TXRange& range);
    
    /**
     * @brief 批量应用公式
     * @param range 应用范围
     * @param formula 公式模板
     * @return 应用的单元格数量
     */
    TXResult<size_t> applyFormula(const TXRange& range, const std::string& formula);
    
    // ==================== 单个单元格操作 (兼容性) ====================
    
    /**
     * @brief 设置单个数值单元格
     * @param coord 坐标
     * @param value 数值
     */
    TXResult<void> setNumber(const TXCoordinate& coord, double value);
    
    /**
     * @brief 设置单个字符串单元格
     * @param coord 坐标
     * @param value 字符串
     */
    TXResult<void> setString(const TXCoordinate& coord, const std::string& value);
    
    /**
     * @brief 获取单元格值
     * @param coord 坐标
     * @return 单元格值
     */
    TXResult<TXVariant> getValue(const TXCoordinate& coord) const;
    
    /**
     * @brief 检查单元格是否存在
     * @param coord 坐标
     * @return 是否存在
     */
    bool hasCell(const TXCoordinate& coord) const;
    
    // ==================== 高级批量数据导入 ====================
    
    /**
     * @brief 从二维数组导入 - 极速版本
     * @param data 二维数据 (行x列)
     * @param start_coord 起始坐标
     * @param options 导入选项
     * @return 导入的单元格数量
     */
    TXResult<size_t> importData(
        const std::vector<std::vector<TXVariant>>& data,
        const TXCoordinate& start_coord = TXCoordinate(row_t(1), column_t(1)),
        const TXImportOptions& options = {}
    );
    
    /**
     * @brief 从数值矩阵导入 - 纯数值优化
     * @param numbers 数值矩阵
     * @param start_coord 起始坐标
     * @return 导入的单元格数量
     */
    TXResult<size_t> importNumbers(
        const std::vector<std::vector<double>>& numbers,
        const TXCoordinate& start_coord = TXCoordinate(row_t(1), column_t(1))
    );
    
    /**
     * @brief 从CSV内容导入
     * @param csv_content CSV内容
     * @param options 导入选项
     * @return 导入的单元格数量
     */
    TXResult<size_t> importFromCSV(
        const std::string& csv_content,
        const TXImportOptions& options = {}
    );
    
    // ==================== 统计和查询 ====================
    
    /**
     * @brief 获取单元格统计信息 - SIMD优化
     * @param range 统计范围 (可选，默认全部)
     * @return 统计结果
     */
    TXCellStats getStats(const TXRange* range = nullptr) const;
    
    /**
     * @brief 范围求和 - SIMD优化
     * @param range 求和范围
     * @return 求和结果
     */
    TXResult<double> sum(const TXRange& range) const;
    
    /**
     * @brief 查找数值 - SIMD优化
     * @param target_value 目标值
     * @param range 查找范围 (可选)
     * @return 找到的坐标列表
     */
    std::vector<TXCoordinate> findValue(
        double target_value,
        const TXRange* range = nullptr
    ) const;
    
    /**
     * @brief 查找字符串
     * @param target_string 目标字符串
     * @param range 查找范围 (可选)
     * @return 找到的坐标列表
     */
    std::vector<TXCoordinate> findString(
        const std::string& target_string,
        const TXRange* range = nullptr
    ) const;
    
    // ==================== 内存和性能优化 ====================
    
    /**
     * @brief 优化内存布局 - 提高后续操作性能
     */
    void optimizeMemoryLayout();
    
    /**
     * @brief 压缩稀疏数据 - 移除空白单元格
     * @return 压缩前后的大小差
     */
    size_t compressSparseData();
    
    /**
     * @brief 预分配内存
     * @param estimated_cells 预计单元格数量
     */
    void reserve(size_t estimated_cells);
    
    /**
     * @brief 收缩内存到实际使用大小
     */
    void shrink_to_fit();
    
    /**
     * @brief 设置自动优化
     * @param enable 是否启用
     */
    void setAutoOptimize(bool enable) { auto_optimize_ = enable; }
    
    // ==================== 序列化和导出 ====================
    
    /**
     * @brief 零拷贝序列化到内存 - 核心性能方法
     * @param serializer 序列化器
     * @return 序列化结果
     */
    TXResult<void> serializeToMemory(TXZeroCopySerializer& serializer) const;
    
    /**
     * @brief 导出为CSV
     * @param range 导出范围 (可选)
     * @return CSV内容
     */
    TXResult<std::string> exportToCSV(const TXRange* range = nullptr) const;
    
    /**
     * @brief 导出为JSON
     * @param range 导出范围 (可选)
     * @return JSON内容
     */
    TXResult<std::string> exportToJSON(const TXRange* range = nullptr) const;
    
    // ==================== 元数据和属性 ====================
    
    /**
     * @brief 获取工作表名称
     */
    const std::string& getName() const { return name_; }
    
    /**
     * @brief 设置工作表名称
     */
    void setName(const std::string& name) { name_ = name; dirty_ = true; }
    
    /**
     * @brief 获取使用范围
     */
    TXRange getUsedRange() const;
    
    /**
     * @brief 获取最大行号 (0-based)
     */
    uint32_t getMaxRow() const { return max_row_; }
    
    /**
     * @brief 获取最大列号 (0-based)
     */
    uint32_t getMaxCol() const { return max_col_; }
    
    /**
     * @brief 获取单元格总数
     */
    size_t getCellCount() const { return cell_buffer_.size; }
    
    /**
     * @brief 是否为空
     */
    bool empty() const { return cell_buffer_.empty(); }
    
    /**
     * @brief 是否有未保存的更改
     */
    bool isDirty() const { return dirty_; }
    
    /**
     * @brief 标记为已保存
     */
    void markClean() { dirty_ = false; }
    
    // ==================== 性能监控 ====================
    
    /**
     * @brief 工作表性能统计
     */
    struct SheetPerformanceStats {
        size_t total_cells;                  // 总单元格数
        size_t batch_operations;             // 批量操作次数
        double avg_operation_time;           // 平均操作时间
        double cache_hit_ratio;              // 缓存命中率
        size_t memory_usage;                 // 内存使用量
        double compression_ratio;            // 压缩比
    };
    
    /**
     * @brief 获取性能统计
     */
    SheetPerformanceStats getPerformanceStats() const;
    
    /**
     * @brief 重置性能统计
     */
    void resetPerformanceStats();
    
    // ==================== 内部访问 (序列化用) ====================
    
    /**
     * @brief 获取单元格缓冲区 (只读)
     */
    const TXCompactCellBuffer& getCellBuffer() const { return cell_buffer_; }
    
    /**
     * @brief 生成行分组信息 (序列化优化)
     */
    std::vector<TXRowGroup> generateRowGroups() const;

private:
    // 内部辅助方法
    void updateBounds(const TXCoordinate& coord);
    void updateIndex(const TXCoordinate& coord, size_t buffer_index);
    void removeFromIndex(const TXCoordinate& coord);
    void maybeOptimize();                    // 条件优化
    
    // 性能统计更新
    void updateStats(size_t cells_processed, double time_ms) const;
    
    // 坐标转换
    static uint32_t coordToKey(const TXCoordinate& coord);
    static TXCoordinate keyToCoord(uint32_t key);
};

/**
 * @brief 🚀 内存优先工作簿 - 顶层容器
 */
class TXInMemoryWorkbook {
private:
    TXUnifiedMemoryManager memory_manager_;   // 内存管理器
    TXGlobalStringPool& string_pool_;        // 全局字符串池引用
    std::vector<std::unique_ptr<TXInMemorySheet>> sheets_; // 工作表列表
    std::string filename_;                   // 文件名
    bool auto_save_ = false;                 // 自动保存

public:
    /**
     * @brief 创建内存优先工作簿
     * @param filename 文件名 (可选)
     * @return 工作簿实例
     */
    static std::unique_ptr<TXInMemoryWorkbook> create(const std::string& filename = "");
    
    /**
     * @brief 构造函数
     */
    explicit TXInMemoryWorkbook(const std::string& filename = "");
    
    /**
     * @brief 创建工作表
     * @param name 工作表名称
     * @return 工作表引用
     */
    TXInMemorySheet& createSheet(const std::string& name);
    
    /**
     * @brief 获取工作表
     * @param name 工作表名称
     * @return 工作表引用
     */
    TXInMemorySheet& getSheet(const std::string& name);
    
    /**
     * @brief 获取工作表 (按索引)
     * @param index 索引
     * @return 工作表引用
     */
    TXInMemorySheet& getSheet(size_t index);
    
    /**
     * @brief 删除工作表
     * @param name 工作表名称
     */
    bool removeSheet(const std::string& name);
    
    /**
     * @brief 获取工作表数量
     */
    size_t getSheetCount() const { return sheets_.size(); }
    
    /**
     * @brief 保存到文件 - 一次性高性能写出
     * @param filename 文件名 (可选，使用构造时的文件名)
     * @return 保存结果
     */
    TXResult<void> saveToFile(const std::string& filename = "");

private:
    /**
     * @brief 添加XLSX结构文件
     */
    TXResult<void> addXLSXStructureFiles(TXZipArchiveWriter& zip_writer, size_t sheet_count);

    /**
     * @brief 生成[Content_Types].xml
     */
    std::string generateContentTypesXML(size_t sheet_count);

    /**
     * @brief 生成xl/_rels/workbook.xml.rels
     */
    std::string generateWorkbookRelsXML(size_t sheet_count);

public:
    
    /**
     * @brief 序列化到内存
     * @return 序列化后的数据
     */
    TXResult<std::vector<uint8_t>> serializeToMemory();
    
    /**
     * @brief 从文件加载
     * @param filename 文件名
     * @return 加载的工作簿
     */
    static TXResult<std::unique_ptr<TXInMemoryWorkbook>> loadFromFile(const std::string& filename);
};

} // namespace TinaXlsx
 