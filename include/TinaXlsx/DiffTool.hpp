/**
 * @file DiffTool.hpp
 * @brief 基于Myers算法的Excel差异对比工具
 */

#pragma once

#include "Types.hpp"
#include "Exception.hpp"
#include "Reader.hpp"
#include "Writer.hpp"
#include <vector>
#include <string>
#include <memory>
#include <functional>
#include <optional>

namespace TinaXlsx {

/**
 * @brief Excel差异对比工具
 * 
 * 基于Myers差分算法实现高性能的Excel文件对比功能
 */
class DiffTool {
public:
    /**
     * @brief 差异类型枚举
     */
    enum class DiffType {
        Unchanged,  // 未改变
        Added,      // 新增
        Deleted,    // 删除
        Modified    // 修改
    };
    
    /**
     * @brief 单元格差异信息
     */
    struct CellDiff {
        CellPosition position;      // 单元格位置
        CellValue oldValue;         // 原值
        CellValue newValue;         // 新值
        DiffType type;              // 差异类型
    };
    
    /**
     * @brief 行差异信息
     */
    struct RowDiff {
        RowIndex oldRowIndex;       // 原行索引（如果是新增行则为-1）
        RowIndex newRowIndex;       // 新行索引（如果是删除行则为-1）
        DiffType type;              // 差异类型
        RowData oldRowData;         // 原行数据
        RowData newRowData;         // 新行数据
        std::vector<CellDiff> cellDiffs; // 单元格级别的差异
        double similarity;          // 相似度（0.0-1.0）
    };
    
    /**
     * @brief 差异对比结果
     */
    struct DiffResult {
        std::vector<RowDiff> rowDiffs;      // 行差异列表
        size_t addedRowCount = 0;           // 新增行数
        size_t deletedRowCount = 0;         // 删除行数
        size_t modifiedRowCount = 0;        // 修改行数
        size_t unchangedRowCount = 0;       // 未改变行数
        double overallSimilarity = 0.0;     // 整体相似度
        
        /**
         * @brief 检查是否有差异
         * @return bool 是否有差异
         */
        [[nodiscard]] bool hasDifferences() const {
            return addedRowCount > 0 || deletedRowCount > 0 || modifiedRowCount > 0;
        }
        
        /**
         * @brief 获取总差异数
         * @return size_t 总差异数
         */
        [[nodiscard]] size_t getTotalDifferences() const {
            return addedRowCount + deletedRowCount + modifiedRowCount;
        }
    };
    
    /**
     * @brief 对比选项
     */
    struct CompareOptions {
        double similarityThreshold = 0.6;      // 相似度阈值
        bool ignoreCase = false;                // 是否忽略大小写
        bool ignoreWhitespace = false;          // 是否忽略空白字符
        bool ignoreEmptyRows = true;            // 是否忽略空行
        bool enableCellLevelDiff = true;        // 是否启用单元格级别差异
        ColumnIndex maxColumns = 0;             // 最大列数（0表示无限制）
        RowIndex maxRows = 0;                   // 最大行数（0表示无限制）
    };

private:
    CompareOptions options_;

public:
    /**
     * @brief 构造函数
     * @param options 对比选项
     */
    explicit DiffTool(const CompareOptions& options = CompareOptions{});
    
    /**
     * @brief 设置对比选项
     * @param options 对比选项
     */
    void setOptions(const CompareOptions& options);
    
    /**
     * @brief 获取对比选项
     * @return const CompareOptions& 对比选项
     */
    [[nodiscard]] const CompareOptions& getOptions() const { return options_; }
    
    /**
     * @brief 对比两个Excel文件
     * @param file1 第一个文件路径
     * @param file2 第二个文件路径
     * @param sheet1 第一个文件的工作表名（空表示默认）
     * @param sheet2 第二个文件的工作表名（空表示默认）
     * @return DiffResult 对比结果
     */
    [[nodiscard]] DiffResult compareFiles(
        const std::string& file1, 
        const std::string& file2,
        const std::string& sheet1 = "",
        const std::string& sheet2 = "");
    
    /**
     * @brief 对比两个表格数据
     * @param table1 第一个表格数据
     * @param table2 第二个表格数据
     * @return DiffResult 对比结果
     */
    [[nodiscard]] DiffResult compareData(
        const TableData& table1,
        const TableData& table2);
    
    /**
     * @brief 将对比结果输出到Excel文件
     * @param result 对比结果
     * @param outputFile 输出文件路径
     * @param table1 原始表格数据
     * @param table2 新表格数据
     * @param columnHeaders 列标题（可选）
     * @return bool 是否成功
     */
    bool exportResult(
        const DiffResult& result,
        const std::string& outputFile,
        const TableData& table1,
        const TableData& table2,
        const std::vector<std::string>& columnHeaders = {});
    
    /**
     * @brief 生成HTML格式的对比报告
     * @param result 对比结果
     * @param table1 原始表格数据
     * @param table2 新表格数据
     * @param outputFile 输出HTML文件路径
     * @param title 报告标题
     * @return bool 是否成功
     */
    bool exportHtmlReport(
        const DiffResult& result,
        const TableData& table1,
        const TableData& table2,
        const std::string& outputFile,
        const std::string& title = "Excel差异对比报告");
    
    /**
     * @brief 计算两行的相似度
     * @param row1 第一行数据
     * @param row2 第二行数据
     * @return double 相似度（0.0-1.0）
     */
    [[nodiscard]] double calculateRowSimilarity(const RowData& row1, const RowData& row2) const;
    
    /**
     * @brief 计算两个字符串的相似度
     * @param str1 第一个字符串
     * @param str2 第二个字符串
     * @return double 相似度（0.0-1.0）
     */
    [[nodiscard]] double calculateStringSimilarity(const std::string& str1, const std::string& str2) const;

private:
    /**
     * @brief 使用Myers算法计算序列差异
     * @param sequence1 第一个序列
     * @param sequence2 第二个序列
     * @return std::vector<RowDiff> 差异列表
     */
    std::vector<RowDiff> myersDiff(const TableData& sequence1, const TableData& sequence2) const;
    
    /**
     * @brief 查找最佳匹配
     * @param table1 第一个表格
     * @param table2 第二个表格
     * @return std::vector<std::pair<int, int>> 匹配对列表
     */
    std::vector<std::pair<int, int>> findBestMatches(const TableData& table1, const TableData& table2) const;
    
    /**
     * @brief 计算单元格级别的差异
     * @param row1 第一行数据
     * @param row2 第二行数据
     * @param rowIndex 行索引
     * @return std::vector<CellDiff> 单元格差异列表
     */
    std::vector<CellDiff> calculateCellDiffs(const RowData& row1, const RowData& row2, RowIndex rowIndex) const;
    
    /**
     * @brief 预处理单元格值
     * @param value 单元格值
     * @return CellValue 处理后的值
     */
    CellValue preprocessCellValue(const CellValue& value) const;
    
    /**
     * @brief 预处理表格数据
     * @param table 表格数据
     * @return TableData 处理后的表格数据
     */
    TableData preprocessTable(const TableData& table) const;
    
    /**
     * @brief 检查两个单元格值是否相等
     * @param value1 第一个值
     * @param value2 第二个值
     * @return bool 是否相等
     */
    bool areValuesEqual(const CellValue& value1, const CellValue& value2) const;
    
    /**
     * @brief 检查行是否为空
     * @param row 行数据
     * @return bool 是否为空
     */
    bool isEmptyRow(const RowData& row) const;
};

/**
 * @brief 差异对比工具构建器
 */
class DiffToolBuilder {
private:
    DiffTool::CompareOptions options_;

public:
    /**
     * @brief 设置相似度阈值
     * @param threshold 阈值（0.0-1.0）
     * @return DiffToolBuilder& 返回自身引用
     */
    DiffToolBuilder& setSimilarityThreshold(double threshold);
    
    /**
     * @brief 设置是否忽略大小写
     * @param ignore 是否忽略
     * @return DiffToolBuilder& 返回自身引用
     */
    DiffToolBuilder& setIgnoreCase(bool ignore = true);
    
    /**
     * @brief 设置是否忽略空白字符
     * @param ignore 是否忽略
     * @return DiffToolBuilder& 返回自身引用
     */
    DiffToolBuilder& setIgnoreWhitespace(bool ignore = true);
    
    /**
     * @brief 设置是否忽略空行
     * @param ignore 是否忽略
     * @return DiffToolBuilder& 返回自身引用
     */
    DiffToolBuilder& setIgnoreEmptyRows(bool ignore = true);
    
    /**
     * @brief 设置是否启用单元格级别差异
     * @param enable 是否启用
     * @return DiffToolBuilder& 返回自身引用
     */
    DiffToolBuilder& setEnableCellLevelDiff(bool enable = true);
    
    /**
     * @brief 设置最大列数
     * @param maxColumns 最大列数
     * @return DiffToolBuilder& 返回自身引用
     */
    DiffToolBuilder& setMaxColumns(ColumnIndex maxColumns);
    
    /**
     * @brief 设置最大行数
     * @param maxRows 最大行数
     * @return DiffToolBuilder& 返回自身引用
     */
    DiffToolBuilder& setMaxRows(RowIndex maxRows);
    
    /**
     * @brief 构建差异对比工具
     * @return std::unique_ptr<DiffTool> 差异对比工具
     */
    [[nodiscard]] std::unique_ptr<DiffTool> build() const;
};

} // namespace TinaXlsx 