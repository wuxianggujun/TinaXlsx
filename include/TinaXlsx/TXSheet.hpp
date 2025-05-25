#pragma once

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

    /**
     * @brief 单元格坐标
     */
    struct CellCoordinate {
        std::size_t row;    ///< 行号（1开始）
        std::size_t col;    ///< 列号（1开始）
        
        CellCoordinate(std::size_t r = 1, std::size_t c = 1) : row(r), col(c) {}
        
        bool operator==(const CellCoordinate& other) const {
            return row == other.row && col == other.col;
        }
    };

    /**
     * @brief 单元格范围
     */
    struct CellRange {
        CellCoordinate start;   ///< 起始坐标
        CellCoordinate end;     ///< 结束坐标
        
        CellRange() = default;
        CellRange(const CellCoordinate& s, const CellCoordinate& e) : start(s), end(e) {}
        CellRange(std::size_t start_row, std::size_t start_col, std::size_t end_row, std::size_t end_col)
            : start(start_row, start_col), end(end_row, end_col) {}
    };

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

    /**
     * @brief 获取单元格值
     * @param row 行号（1开始）
     * @param col 列号（1开始）
     * @return 单元格值
     */
    CellValue getCellValue(std::size_t row, std::size_t col) const;

    /**
     * @brief 获取单元格值
     * @param coord 单元格坐标
     * @return 单元格值
     */
    CellValue getCellValue(const CellCoordinate& coord) const;

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
    bool setCellValue(std::size_t row, std::size_t col, const CellValue& value);

    /**
     * @brief 设置单元格值
     * @param coord 单元格坐标
     * @param value 单元格值
     * @return 成功返回true，失败返回false
     */
    bool setCellValue(const CellCoordinate& coord, const CellValue& value);

    /**
     * @brief 设置单元格值（使用A1格式）
     * @param address 单元格地址，如"A1", "B2"
     * @param value 单元格值
     * @return 成功返回true，失败返回false
     */
    bool setCellValue(const std::string& address, const CellValue& value);

    /**
     * @brief 获取单元格
     * @param row 行号（1开始）
     * @param col 列号（1开始）
     * @return 单元格指针，如果不存在返回nullptr
     */
    TXCell* getCell(std::size_t row, std::size_t col);

    /**
     * @brief 获取单元格（const版本）
     * @param row 行号（1开始）
     * @param col 列号（1开始）
     * @return 单元格指针，如果不存在返回nullptr
     */
    const TXCell* getCell(std::size_t row, std::size_t col) const;

    /**
     * @brief 插入行
     * @param row 插入位置（1开始）
     * @param count 插入行数
     * @return 成功返回true，失败返回false
     */
    bool insertRows(std::size_t row, std::size_t count = 1);

    /**
     * @brief 删除行
     * @param row 删除位置（1开始）
     * @param count 删除行数
     * @return 成功返回true，失败返回false
     */
    bool deleteRows(std::size_t row, std::size_t count = 1);

    /**
     * @brief 插入列
     * @param col 插入位置（1开始）
     * @param count 插入列数
     * @return 成功返回true，失败返回false
     */
    bool insertColumns(std::size_t col, std::size_t count = 1);

    /**
     * @brief 删除列
     * @param col 删除位置（1开始）
     * @param count 删除列数
     * @return 成功返回true，失败返回false
     */
    bool deleteColumns(std::size_t col, std::size_t count = 1);

    /**
     * @brief 获取使用的行数
     * @return 最大使用行号
     */
    std::size_t getUsedRowCount() const;

    /**
     * @brief 获取使用的列数
     * @return 最大使用列号
     */
    std::size_t getUsedColumnCount() const;

    /**
     * @brief 获取使用的范围
     * @return 使用的单元格范围
     */
    CellRange getUsedRange() const;

    /**
     * @brief 清空工作表
     */
    void clear();

    /**
     * @brief 批量设置单元格值（高性能版本）
     * @param values 坐标到值的映射
     * @return 成功设置的单元格数量
     */
    std::size_t setCellValues(const std::vector<std::pair<CellCoordinate, CellValue>>& values);

    /**
     * @brief 批量获取单元格值（高性能版本）
     * @param coords 坐标列表
     * @return 坐标到值的映射
     */
    std::vector<std::pair<CellCoordinate, CellValue>> getCellValues(const std::vector<CellCoordinate>& coords) const;

    /**
     * @brief 设置范围内的值
     * @param range 单元格范围
     * @param values 值的二维数组
     * @return 成功返回true，失败返回false
     */
    bool setRangeValues(const CellRange& range, const std::vector<std::vector<CellValue>>& values);

    /**
     * @brief 获取范围内的值
     * @param range 单元格范围
     * @return 值的二维数组
     */
    std::vector<std::vector<CellValue>> getRangeValues(const CellRange& range) const;

    /**
     * @brief A1格式地址转换为坐标
     * @param address A1格式地址，如"A1", "B2"
     * @return 坐标
     */
    static CellCoordinate addressToCoordinate(const std::string& address);

    /**
     * @brief 坐标转换为A1格式地址
     * @param coord 坐标
     * @return A1格式地址
     */
    static std::string coordinateToAddress(const CellCoordinate& coord);

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