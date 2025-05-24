/**
 * @file Reader.hpp
 * @brief 高性能Excel读取器
 */

#pragma once

#include "Types.hpp"
#include "Exception.hpp"
#include <memory>
#include <string>
#include <vector>
#include <functional>
#include <optional>

namespace TinaXlsx {

/**
 * @brief 高性能Excel读取器
 * 
 * 基于minizip-ng和expat的高性能读取器，替代xlsxio
 * 支持流式读取和批量读取，使用RAII管理资源，保证异常安全
 */
class Reader {
public:
    /**
     * @brief 行读取回调函数类型
     * @param rowIndex 行索引（0基于）
     * @param rowData 行数据
     * @return bool true继续读取，false停止读取
     */
    using RowCallback = std::function<bool(RowIndex rowIndex, const RowData& rowData)>;
    
    /**
     * @brief 单元格读取回调函数类型
     * @param position 单元格位置
     * @param value 单元格值
     * @return bool true继续读取，false停止读取
     */
    using CellCallback = std::function<bool(const CellPosition& position, const CellValue& value)>;

private:
    struct Impl;
    std::unique_ptr<Impl> pImpl_;

public:
    /**
     * @brief 构造函数
     * @param filePath Excel文件路径
     * @throws FileException 如果文件无法打开
     */
    explicit Reader(const std::string& filePath);
    
    /**
     * @brief 移动构造函数
     */
    Reader(Reader&& other) noexcept;
    
    /**
     * @brief 移动赋值操作符
     */
    Reader& operator=(Reader&& other) noexcept;
    
    /**
     * @brief 析构函数
     */
    ~Reader();
    
    // 禁用拷贝构造和拷贝赋值
    Reader(const Reader&) = delete;
    Reader& operator=(const Reader&) = delete;
    
    /**
     * @brief 获取工作表名称列表
     * @return std::vector<std::string> 工作表名称列表
     */
    [[nodiscard]] std::vector<std::string> getSheetNames() const;
    
    /**
     * @brief 打开指定工作表
     * @param sheetName 工作表名称
     * @return bool 是否成功打开
     */
    bool openSheet(const std::string& sheetName);
    
    /**
     * @brief 打开指定索引的工作表
     * @param sheetIndex 工作表索引（0基于）
     * @return bool 是否成功打开
     */
    bool openSheet(SheetIndex sheetIndex);
    
    /**
     * @brief 获取当前工作表的行数（估算值）
     * @return std::optional<RowIndex> 行数，如果无法确定则返回空
     */
    [[nodiscard]] std::optional<RowIndex> getRowCount() const;
    
    /**
     * @brief 获取当前工作表的列数（估算值）
     * @return std::optional<ColumnIndex> 列数，如果无法确定则返回空
     */
    [[nodiscard]] std::optional<ColumnIndex> getColumnCount() const;
    
    /**
     * @brief 流式读取下一行数据
     * @param rowData 输出的行数据
     * @param maxColumns 最大列数，0表示无限制
     * @return bool 是否成功读取到数据
     */
    bool readNextRow(RowData& rowData, ColumnIndex maxColumns = 0);
    
    /**
     * @brief 读取指定行的数据
     * @param rowIndex 行索引（0基于）
     * @param maxColumns 最大列数，0表示无限制
     * @return std::optional<RowData> 行数据，如果行不存在则返回空
     */
    [[nodiscard]] std::optional<RowData> readRow(RowIndex rowIndex, ColumnIndex maxColumns = 0);
    
    /**
     * @brief 读取指定单元格的值
     * @param position 单元格位置
     * @return std::optional<CellValue> 单元格值，如果单元格不存在则返回空
     */
    [[nodiscard]] std::optional<CellValue> readCell(const CellPosition& position);
    
    /**
     * @brief 读取指定范围的数据
     * @param range 单元格范围
     * @return TableData 范围内的数据
     */
    [[nodiscard]] TableData readRange(const CellRange& range);
    
    /**
     * @brief 读取整个工作表的数据
     * @param maxRows 最大行数，0表示无限制
     * @param maxColumns 最大列数，0表示无限制
     * @param skipEmptyRows 是否跳过空行
     * @return TableData 工作表数据
     */
    [[nodiscard]] TableData readAll(RowIndex maxRows = 0, ColumnIndex maxColumns = 0, bool skipEmptyRows = true);
    
    /**
     * @brief 使用回调函数流式读取所有行
     * @param callback 行读取回调函数
     * @param maxColumns 最大列数，0表示无限制
     * @param skipEmptyRows 是否跳过空行
     * @return RowIndex 实际读取的行数
     */
    RowIndex readAllRows(const RowCallback& callback, ColumnIndex maxColumns = 0, bool skipEmptyRows = true);
    
    /**
     * @brief 使用回调函数流式读取所有单元格
     * @param callback 单元格读取回调函数
     * @param range 读取范围，如果为空则读取整个工作表
     * @return size_t 实际读取的单元格数
     */
    size_t readAllCells(const CellCallback& callback, const std::optional<CellRange>& range = std::nullopt);
    
    /**
     * @brief 重置读取位置到工作表开头
     */
    void reset();
    
    /**
     * @brief 检查是否已到达工作表末尾
     * @return bool 是否已到达末尾
     */
    [[nodiscard]] bool isAtEnd() const;
    
    /**
     * @brief 获取当前读取位置
     * @return RowIndex 当前行索引
     */
    [[nodiscard]] RowIndex getCurrentRowIndex() const;
    
    /**
     * @brief 判断指定行是否为空行
     * @param rowData 行数据
     * @return bool 是否为空行
     */
    [[nodiscard]] static bool isEmptyRow(const RowData& rowData);
    
    /**
     * @brief 判断指定单元格是否为空
     * @param value 单元格值
     * @return bool 是否为空
     */
    [[nodiscard]] static bool isEmptyCell(const CellValue& value);
    
    /**
     * @brief 将字符串转换为CellValue
     * @param str 字符串
     * @return CellValue 转换后的值
     */
    [[nodiscard]] static CellValue stringToCellValue(const std::string& str);
    
    /**
     * @brief 将CellValue转换为字符串
     * @param value 单元格值
     * @return std::string 转换后的字符串
     */
    [[nodiscard]] static std::string cellValueToString(const CellValue& value);
};

} // namespace TinaXlsx 