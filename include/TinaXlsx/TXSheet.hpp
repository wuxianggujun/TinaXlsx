#pragma once

#include "TXCoordinate.hpp"
#include "TXRange.hpp"
#include "TXCell.hpp"
#include "TXTypes.hpp"
#include <string>
#include <vector>
#include <memory>
#include <variant>

namespace TinaXlsx {

// Forward declarations
class TXCell;

/**
 * @brief Excel工作表类
 * 
 * 表示Excel中的一个Sheet，负责管理Sheet内的数据操作
 */
class TXSheet {
public:
    // ==================== 类型别名 ====================
    using CellValue = cell_value_t;
    using Coordinate = TXCoordinate;
    using Range = TXRange;

public:
    explicit TXSheet(const std::string& name);
    ~TXSheet();

    // 禁用拷贝构造和赋值
    TXSheet(const TXSheet&) = delete;
    TXSheet& operator=(const TXSheet&) = delete;

    // 支持移动构造和赋值
    TXSheet(TXSheet&& other) noexcept;
    TXSheet& operator=(TXSheet&& other) noexcept;

    /**
     * @brief 获取工作表名称
     * @return 工作表名称
     */
    const std::string& getName() const;

    /**
     * @brief 设置工作表名称
     * @param name 新名称
     */
    void setName(const std::string& name);

    // ==================== 单元格值操作 ====================

    /**
     * @brief 获取单元格值
     * @param row 行号 (1-based)
     * @param col 列号 (1-based)
     * @return 单元格值
     */
    CellValue getCellValue(row_t row, column_t col) const;

    /**
     * @brief 获取单元格值
     * @param coord 单元格坐标
     * @return 单元格值
     */
    CellValue getCellValue(const Coordinate& coord) const;

    /**
     * @brief 获取单元格值（使用A1格式）
     * @param address 单元格地址，如"A1", "B2"
     * @return 单元格值
     */
    CellValue getCellValue(const std::string& address) const;

    /**
     * @brief 设置单元格值
     * @param row 行号（1开始）
     * @param col 列号（1开始）
     * @param value 单元格值
     * @return 成功返回true，失败返回false
     */
    bool setCellValue(row_t row, column_t col, const CellValue& value);

    /**
     * @brief 设置单元格值
     * @param coord 单元格坐标
     * @param value 单元格值
     * @return 成功返回true，失败返回false
     */
    bool setCellValue(const Coordinate& coord, const CellValue& value);

    /**
     * @brief 设置单元格值（使用A1格式）
     * @param address 单元格地址，如"A1", "B2"
     * @param value 单元格值
     * @return 成功返回true，失败返回false
     */
    bool setCellValue(const std::string& address, const CellValue& value);

    // ==================== 单元格对象访问 ====================

    /**
     * @brief 获取单元格
     * @param row 行号（1开始）
     * @param col 列号（1开始）
     * @return 单元格指针，如果不存在返回nullptr
     */
    TXCell* getCell(row_t row, column_t col);

    /**
     * @brief 获取单元格（const版本）
     * @param row 行号（1开始）
     * @param col 列号（1开始）
     * @return 单元格指针，如果不存在返回nullptr
     */
    const TXCell* getCell(row_t row, column_t col) const;

    /**
     * @brief 获取单元格
     * @param coord 单元格坐标
     * @return 单元格指针，如果不存在返回nullptr
     */
    TXCell* getCell(const Coordinate& coord);

    /**
     * @brief 获取单元格（const版本）
     * @param coord 单元格坐标
     * @return 单元格指针，如果不存在返回nullptr
     */
    const TXCell* getCell(const Coordinate& coord) const;

    /**
     * @brief 获取单元格（使用A1格式）
     * @param address 单元格地址，如"A1", "B2"
     * @return 单元格指针，如果不存在返回nullptr
     */
    TXCell* getCell(const std::string& address);

    /**
     * @brief 获取单元格（const版本，使用A1格式）
     * @param address 单元格地址，如"A1", "B2"
     * @return 单元格指针，如果不存在返回nullptr
     */
    const TXCell* getCell(const std::string& address) const;

    // ==================== 行列操作 ====================

    /**
     * @brief 插入行
     * @param row 插入位置（1开始）
     * @param count 插入行数
     * @return 成功返回true，失败返回false
     */
    bool insertRows(row_t row, row_t count = row_t(1));

    /**
     * @brief 删除行
     * @param row 删除位置（1开始）
     * @param count 删除行数
     * @return 成功返回true，失败返回false
     */
    bool deleteRows(row_t row, row_t count = row_t(1));

    /**
     * @brief 插入列
     * @param col 插入位置（1开始）
     * @param count 插入列数
     * @return 成功返回true，失败返回false
     */
    bool insertColumns(column_t col, column_t count = column_t(1));

    /**
     * @brief 删除列
     * @param col 删除位置（1开始）
     * @param count 删除列数
     * @return 成功返回true，失败返回false
     */
    bool deleteColumns(column_t col, column_t count = column_t(1));

    // ==================== 范围信息 ====================

    /**
     * @brief 获取使用的行数
     * @return 最大使用行号
     */
    row_t getUsedRowCount() const;

    /**
     * @brief 获取使用的列数
     * @return 最大使用列号
     */
    column_t getUsedColumnCount() const;

    /**
     * @brief 获取使用的范围
     * @return 使用的单元格范围
     */
    Range getUsedRange() const;

    /**
     * @brief 清空工作表
     */
    void clear();

    // ==================== 批量操作 ====================

    /**
     * @brief 批量设置单元格值（高性能版本）
     * @param values 坐标到值的映射
     * @return 成功设置的单元格数量
     */
    std::size_t setCellValues(const std::vector<std::pair<Coordinate, CellValue>>& values);

    /**
     * @brief 批量获取单元格值（高性能版本）
     * @param coords 坐标列表
     * @return 坐标到值的映射
     */
    std::vector<std::pair<Coordinate, CellValue>> getCellValues(const std::vector<Coordinate>& coords) const;

    /**
     * @brief 设置范围内的值
     * @param range 单元格范围
     * @param values 值的二维数组
     * @return 成功返回true，失败返回false
     */
    bool setRangeValues(const Range& range, const std::vector<std::vector<CellValue>>& values);

    /**
     * @brief 获取范围内的值
     * @param range 单元格范围
     * @return 值的二维数组
     */
    std::vector<std::vector<CellValue>> getRangeValues(const Range& range) const;

    // ==================== 合并单元格操作 ====================

    /**
     * @brief 合并单元格区域
     * @param startRow 起始行
     * @param startCol 起始列
     * @param endRow 结束行
     * @param endCol 结束列
     * @return 成功返回true，失败返回false
     */
    bool mergeCells(row_t startRow, column_t startCol,
                   row_t endRow, column_t endCol);

    /**
     * @brief 合并单元格区域
     * @param range 要合并的范围
     * @return 成功返回true，失败返回false
     */
    bool mergeCells(const Range& range);

    /**
     * @brief 合并单元格区域（使用A1格式）
     * @param rangeStr 范围字符串，如"A1:C3"
     * @return 成功返回true，失败返回false
     */
    bool mergeCells(const std::string& rangeStr);

    /**
     * @brief 拆分包含指定单元格的合并区域
     * @param row 单元格行号
     * @param col 单元格列号
     * @return 成功返回true，失败返回false
     */
    bool unmergeCells(row_t row, column_t col);

    /**
     * @brief 拆分指定范围的所有合并区域
     * @param range 范围
     * @return 拆分的区域数量
     */
    std::size_t unmergeCellsInRange(const Range& range);

    /**
     * @brief 检查单元格是否被合并
     * @param row 行号
     * @param col 列号
     * @return 被合并返回true，否则返回false
     */
    bool isCellMerged(row_t row, column_t col) const;

    /**
     * @brief 获取包含指定单元格的合并区域
     * @param row 行号
     * @param col 列号
     * @return 合并区域，如果不存在返回空Range
     */
    Range getMergeRegion(row_t row, column_t col) const;

    /**
     * @brief 获取所有合并区域
     * @return 合并区域列表
     */
    std::vector<Range> getAllMergeRegions() const;

    /**
     * @brief 获取合并区域数量
     * @return 合并区域数量
     */
    std::size_t getMergeCount() const;

    // ==================== 公式计算 ====================

    /**
     * @brief 计算所有公式
     * @return 成功计算的公式数量
     */
    std::size_t calculateAllFormulas();

    /**
     * @brief 计算指定范围内的公式
     * @param range 范围
     * @return 成功计算的公式数量
     */
    std::size_t calculateFormulasInRange(const Range& range);

    /**
     * @brief 设置单元格公式
     * @param row 行号
     * @param col 列号
     * @param formula 公式字符串
     * @return 成功返回true，失败返回false
     */
    bool setCellFormula(row_t row, column_t col, const std::string& formula);

    /**
     * @brief 获取单元格公式
     * @param row 行号
     * @param col 列号
     * @return 公式字符串，如果不是公式返回空字符串
     */
    std::string getCellFormula(row_t row, column_t col) const;

    /**
     * @brief 批量设置公式（高性能版本）
     * @param formulas 坐标到公式的映射
     * @return 成功设置的公式数量
     */
    std::size_t setCellFormulas(const std::vector<std::pair<Coordinate, std::string>>& formulas);

    // ==================== 格式化操作 ====================

    /**
     * @brief 设置单元格数字格式
     * @param row 行号
     * @param col 列号
     * @param formatType 格式类型
     * @param decimalPlaces 小数位数
     * @return 成功返回true，失败返回false
     */
    bool setCellNumberFormat(row_t row, column_t col, 
                           TXCell::NumberFormat formatType, int decimalPlaces = 2);

    /**
     * @brief 设置单元格自定义格式
     * @param row 行号
     * @param col 列号
     * @param formatString 格式字符串
     * @return 成功返回true，失败返回false
     */
    bool setCellCustomFormat(row_t row, column_t col, 
                           const std::string& formatString);

    /**
     * @brief 设置范围内单元格的格式
     * @param range 范围
     * @param formatType 格式类型
     * @param decimalPlaces 小数位数
     * @return 成功设置的单元格数量
     */
    std::size_t setRangeNumberFormat(const Range& range, TXCell::NumberFormat formatType, 
                                   int decimalPlaces = 2);

    /**
     * @brief 获取单元格格式化后的值
     * @param row 行号
     * @param col 列号
     * @return 格式化后的字符串
     */
    std::string getCellFormattedValue(row_t row, column_t col) const;

    /**
     * @brief 批量设置格式（高性能版本）
     * @param formats 坐标到格式的映射
     * @return 成功设置的格式数量
     */
    std::size_t setCellFormats(const std::vector<std::pair<Coordinate, TXCell::NumberFormat>>& formats);

    // ==================== 样式操作 ====================
    
    /**
     * @brief 设置单元格样式
     * @param row 行号
     * @param col 列号
     * @param style 单元格样式
     * @return 成功返回true，失败返回false
     */
    bool setCellStyle(row_t row, column_t col, const class TXCellStyle& style);
    
    /**
     * @brief 设置单元格样式（使用A1格式）
     * @param address 单元格地址，如"A1"
     * @param style 单元格样式
     * @return 成功返回true，失败返回false
     */
    bool setCellStyle(const std::string& address, const class TXCellStyle& style);
    
    /**
     * @brief 设置范围内单元格的样式
     * @param range 范围
     * @param style 单元格样式
     * @return 成功设置的单元格数量
     */
    std::size_t setRangeStyle(const Range& range, const class TXCellStyle& style);
    
    /**
     * @brief 批量设置样式（高性能版本）
     * @param styles 坐标到样式的映射
     * @return 成功设置的样式数量
     */
    std::size_t setCellStyles(const std::vector<std::pair<Coordinate, class TXCellStyle>>& styles);

    // ==================== 便捷方法 ====================

    /**
     * @brief 从A1格式地址创建坐标
     * @param address A1格式地址，如"A1", "B2"
     * @return 坐标对象
     */
    static Coordinate addressToCoordinate(const std::string& address) {
        return Coordinate::fromAddress(address);
    }

    /**
     * @brief 坐标转换为A1格式地址
     * @param coord 坐标
     * @return A1格式地址
     */
    static std::string coordinateToAddress(const Coordinate& coord) {
        return coord.toAddress();
    }

    /**
     * @brief 从A1格式范围地址创建范围
     * @param rangeAddress 范围地址，如"A1:B5"
     * @return 范围对象
     */
    static Range addressToRange(const std::string& rangeAddress) {
        return Range::fromAddress(rangeAddress);
    }

    /**
     * @brief 范围转换为A1格式地址
     * @param range 范围
     * @return A1格式范围地址
     */
    static std::string rangeToAddress(const Range& range) {
        return range.toAddress();
    }

    /**
     * @brief 获取最后的错误信息
     * @return 错误信息字符串
     */
    const std::string& getLastError() const;

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

} // namespace TinaXlsx 