#pragma once

#include "TXCoordinate.hpp"
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
    /**
     * @brief 单元格值类型
     */
    using CellValue = std::variant<std::monostate, std::string, double, int64_t, bool>;
    
    // 使用统一的坐标系统
    using Coordinate = TXCoordinate;  ///< 坐标类型别名
    using Range = TXRange;            ///< 范围类型别名

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
     * @param row 行号（1开始）
     * @param col 列号（1开始）
     * @return 单元格值
     */
    CellValue getCellValue(TXTypes::RowIndex row, TXTypes::ColIndex col) const;

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
    bool setCellValue(TXTypes::RowIndex row, TXTypes::ColIndex col, const CellValue& value);

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
    TXCell* getCell(TXTypes::RowIndex row, TXTypes::ColIndex col);

    /**
     * @brief 获取单元格（const版本）
     * @param row 行号（1开始）
     * @param col 列号（1开始）
     * @return 单元格指针，如果不存在返回nullptr
     */
    const TXCell* getCell(TXTypes::RowIndex row, TXTypes::ColIndex col) const;

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
    bool insertRows(TXTypes::RowIndex row, TXTypes::RowIndex count = 1);

    /**
     * @brief 删除行
     * @param row 删除位置（1开始）
     * @param count 删除行数
     * @return 成功返回true，失败返回false
     */
    bool deleteRows(TXTypes::RowIndex row, TXTypes::RowIndex count = 1);

    /**
     * @brief 插入列
     * @param col 插入位置（1开始）
     * @param count 插入列数
     * @return 成功返回true，失败返回false
     */
    bool insertColumns(TXTypes::ColIndex col, TXTypes::ColIndex count = 1);

    /**
     * @brief 删除列
     * @param col 删除位置（1开始）
     * @param count 删除列数
     * @return 成功返回true，失败返回false
     */
    bool deleteColumns(TXTypes::ColIndex col, TXTypes::ColIndex count = 1);

    // ==================== 范围信息 ====================

    /**
     * @brief 获取使用的行数
     * @return 最大使用行号
     */
    TXTypes::RowIndex getUsedRowCount() const;

    /**
     * @brief 获取使用的列数
     * @return 最大使用列号
     */
    TXTypes::ColIndex getUsedColumnCount() const;

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