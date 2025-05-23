/**
 * @file Worksheet.hpp
 * @brief Excel工作表类
 */

#pragma once

#include "Types.hpp"
#include "Exception.hpp"
#include "Format.hpp"
#include "Cell.hpp"
#include <xlsxwriter.h>
#include <memory>
#include <string>
#include <vector>
#include <functional>
#include <optional>

namespace TinaXlsx {

/**
 * @brief Excel工作表类
 * 
 * 提供对单个Excel工作表的操作接口
 */
class Worksheet {
public:
    /**
     * @brief 行写入回调函数类型
     * @param rowIndex 行索引
     */
    using RowWriteCallback = std::function<void(RowIndex rowIndex)>;

private:
    struct Impl;
    std::unique_ptr<Impl> pImpl_;

public:
    /**
     * @brief 构造函数（仅供内部使用）
     * @param worksheet libxlsxwriter的工作表对象
     * @param workbook libxlsxwriter的工作簿对象
     * @param name 工作表名称
     */
    Worksheet(lxw_worksheet* worksheet, lxw_workbook* workbook, const std::string& name);
    
    /**
     * @brief 移动构造函数
     */
    Worksheet(Worksheet&& other) noexcept;
    
    /**
     * @brief 移动赋值操作符
     */
    Worksheet& operator=(Worksheet&& other) noexcept;
    
    /**
     * @brief 析构函数
     */
    ~Worksheet();
    
    // 禁用拷贝构造和拷贝赋值
    Worksheet(const Worksheet&) = delete;
    Worksheet& operator=(const Worksheet&) = delete;
    
    /**
     * @brief 获取工作表名称
     * @return std::string 工作表名称
     */
    [[nodiscard]] std::string getName() const;
    
    /**
     * @brief 写入单元格值
     * @param position 单元格位置
     * @param value 单元格值
     * @param format 格式（可选）
     */
    void writeCell(const CellPosition& position, const CellValue& value, Format* format = nullptr);
    
    /**
     * @brief 写入单元格值（行列索引）
     * @param row 行索引（0基于）
     * @param column 列索引（0基于）
     * @param value 单元格值
     * @param format 格式（可选）
     */
    void writeCell(RowIndex row, ColumnIndex column, const CellValue& value, Format* format = nullptr);
    
    /**
     * @brief 写入字符串
     * @param position 单元格位置
     * @param value 字符串值
     * @param format 格式（可选）
     */
    void writeString(const CellPosition& position, const std::string& value, Format* format = nullptr);
    
    /**
     * @brief 写入数字
     * @param position 单元格位置
     * @param value 数字值
     * @param format 格式（可选）
     */
    void writeNumber(const CellPosition& position, double value, Format* format = nullptr);
    
    /**
     * @brief 写入整数
     * @param position 单元格位置
     * @param value 整数值
     * @param format 格式（可选）
     */
    void writeInteger(const CellPosition& position, int64_t value, Format* format = nullptr);
    
    /**
     * @brief 写入布尔值
     * @param position 单元格位置
     * @param value 布尔值
     * @param format 格式（可选）
     */
    void writeBoolean(const CellPosition& position, bool value, Format* format = nullptr);
    
    /**
     * @brief 写入日期时间
     * @param position 单元格位置
     * @param datetime Excel日期序列号
     * @param format 格式（可选）
     */
    void writeDateTime(const CellPosition& position, double datetime, Format* format = nullptr);
    
    /**
     * @brief 写入公式
     * @param position 单元格位置
     * @param formula 公式字符串
     * @param format 格式（可选）
     */
    void writeFormula(const CellPosition& position, const std::string& formula, Format* format = nullptr);
    
    /**
     * @brief 写入URL链接
     * @param position 单元格位置
     * @param url URL地址
     * @param text 显示文本（可选）
     * @param format 格式（可选）
     */
    void writeUrl(const CellPosition& position, const std::string& url, 
                  const std::string& text = "", Format* format = nullptr);
    
    /**
     * @brief 写入一行数据
     * @param rowIndex 行索引（0基于）
     * @param rowData 行数据
     * @param startColumn 起始列索引（0基于）
     * @param format 格式（可选）
     */
    void writeRow(RowIndex rowIndex, const RowData& rowData, 
                  ColumnIndex startColumn = 0, Format* format = nullptr);
    
    /**
     * @brief 写入多行数据
     * @param startPosition 起始位置
     * @param tableData 表格数据
     * @param format 格式（可选）
     */
    void writeTable(const CellPosition& startPosition, const TableData& tableData, Format* format = nullptr);
    
    /**
     * @brief 批量写入数据（高性能）
     * @param range 写入范围
     * @param tableData 表格数据
     * @param format 格式（可选）
     */
    void writeBatch(const CellRange& range, const TableData& tableData, Format* format = nullptr);
    
    /**
     * @brief 合并单元格
     * @param range 合并范围
     * @param value 合并后显示的值
     * @param format 格式（可选）
     */
    void mergeCells(const CellRange& range, const CellValue& value = std::monostate{}, Format* format = nullptr);
    
    /**
     * @brief 设置行高
     * @param rowIndex 行索引（0基于）
     * @param height 行高（点数）
     */
    void setRowHeight(RowIndex rowIndex, double height);
    
    /**
     * @brief 设置多行行高
     * @param startRow 起始行索引
     * @param endRow 结束行索引
     * @param height 行高（点数）
     */
    void setRowHeight(RowIndex startRow, RowIndex endRow, double height);
    
    /**
     * @brief 设置列宽
     * @param columnIndex 列索引（0基于）
     * @param width 列宽（字符数）
     */
    void setColumnWidth(ColumnIndex columnIndex, double width);
    
    /**
     * @brief 设置多列列宽
     * @param startColumn 起始列索引
     * @param endColumn 结束列索引
     * @param width 列宽（字符数）
     */
    void setColumnWidth(ColumnIndex startColumn, ColumnIndex endColumn, double width);
    
    /**
     * @brief 自动调整列宽
     * @param columnIndex 列索引（0基于）
     */
    void autoFitColumn(ColumnIndex columnIndex);
    
    /**
     * @brief 隐藏行
     * @param rowIndex 行索引（0基于）
     */
    void hideRow(RowIndex rowIndex);
    
    /**
     * @brief 隐藏列
     * @param columnIndex 列索引（0基于）
     */
    void hideColumn(ColumnIndex columnIndex);
    
    /**
     * @brief 设置行格式
     * @param rowIndex 行索引（0基于）
     * @param format 格式
     */
    void setRowFormat(RowIndex rowIndex, Format* format);
    
    /**
     * @brief 设置列格式
     * @param columnIndex 列索引（0基于）
     * @param format 格式
     */
    void setColumnFormat(ColumnIndex columnIndex, Format* format);
    
    /**
     * @brief 冻结窗格
     * @param row 冻结行数（0基于）
     * @param column 冻结列数（0基于）
     */
    void freezePanes(RowIndex row, ColumnIndex column);
    
    /**
     * @brief 拆分窗格
     * @param row 拆分行位置（0基于）
     * @param column 拆分列位置（0基于）
     */
    void splitPanes(RowIndex row, ColumnIndex column);
    
    /**
     * @brief 设置打印区域
     * @param range 打印范围
     */
    void setPrintArea(const CellRange& range);
    
    /**
     * @brief 设置重复打印行
     * @param startRow 起始行
     * @param endRow 结束行
     */
    void setRepeatRows(RowIndex startRow, RowIndex endRow);
    
    /**
     * @brief 设置重复打印列
     * @param startColumn 起始列
     * @param endColumn 结束列
     */
    void setRepeatColumns(ColumnIndex startColumn, ColumnIndex endColumn);
    
    /**
     * @brief 设置页面方向
     * @param landscape 是否横向（true为横向，false为纵向）
     */
    void setLandscape(bool landscape = true);
    
    /**
     * @brief 设置纸张大小
     * @param paperSize 纸张大小代码
     */
    void setPaperSize(int paperSize);
    
    /**
     * @brief 设置页边距
     * @param left 左边距（英寸）
     * @param right 右边距（英寸）
     * @param top 上边距（英寸）
     * @param bottom 下边距（英寸）
     */
    void setMargins(double left, double right, double top, double bottom);
    
    /**
     * @brief 设置页眉
     * @param header 页眉文本
     */
    void setHeader(const std::string& header);
    
    /**
     * @brief 设置页脚
     * @param footer 页脚文本
     */
    void setFooter(const std::string& footer);
    
    /**
     * @brief 保护工作表
     * @param password 密码（可选）
     */
    void protect(const std::string& password = "");
    
    /**
     * @brief 取消保护工作表
     */
    void unprotect();
    
    /**
     * @brief 设置是否显示网格线
     * @param show 是否显示
     */
    void showGridlines(bool show = true);
    
    /**
     * @brief 设置缩放级别
     * @param scale 缩放级别（10-400）
     */
    void setZoom(int scale);
    
    /**
     * @brief 设置选择区域
     * @param range 选择范围
     */
    void setSelection(const CellRange& range);
    
    /**
     * @brief 插入图片
     * @param position 插入位置
     * @param filename 图片文件路径
     * @param xScale X轴缩放比例
     * @param yScale Y轴缩放比例
     */
    void insertImage(const CellPosition& position, const std::string& filename, 
                     double xScale = 1.0, double yScale = 1.0);
    
    /**
     * @brief 插入图表
     * @param position 插入位置
     * @param chartType 图表类型
     * @param dataRange 数据范围
     */
    void insertChart(const CellPosition& position, int chartType, const CellRange& dataRange);
    
    /**
     * @brief 添加数据验证
     * @param range 验证范围
     * @param validationType 验证类型
     * @param criteria 验证条件
     * @param value1 值1
     * @param value2 值2（可选）
     */
    void addDataValidation(const CellRange& range, int validationType, int criteria, 
                          const std::string& value1, const std::string& value2 = "");
    
    /**
     * @brief 添加条件格式
     * @param range 应用范围
     * @param type 条件类型
     * @param criteria 条件
     * @param value 阈值
     * @param format 格式
     */
    void addConditionalFormat(const CellRange& range, int type, int criteria, 
                             double value, Format* format);
    
    /**
     * @brief 添加自动筛选
     * @param range 筛选范围
     */
    void addAutoFilter(const CellRange& range);
    
    /**
     * @brief 获取内部工作表对象指针（仅供内部使用）
     * @return lxw_worksheet* 内部工作表对象指针
     */
    [[nodiscard]] lxw_worksheet* getInternalWorksheet() const;
    
    /**
     * @brief 获取关联的工作簿对象指针（仅供内部使用）
     * @return lxw_workbook* 内部工作簿对象指针
     */
    [[nodiscard]] lxw_workbook* getInternalWorkbook() const;
    
    /**
     * @brief 流式写入行数据（高性能）
     * @param rowData 行数据
     * @param format 格式（可选）
     * @return RowIndex 写入的行索引
     */
    RowIndex appendRow(const RowData& rowData, Format* format = nullptr);
    
    /**
     * @brief 使用回调函数批量写入
     * @param startRow 起始行
     * @param rowCount 行数
     * @param callback 写入回调函数
     */
    void writeBatchWithCallback(RowIndex startRow, RowIndex rowCount, const RowWriteCallback& callback);
};

} // namespace TinaXlsx 