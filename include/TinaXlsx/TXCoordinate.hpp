#pragma once

#include "TXTypes.hpp"
#include "TXRow.hpp"
#include "TXColumn.hpp"
#include <string>
#include <vector>

namespace TinaXlsx {

/**
 * @brief 坐标类
 * 
 * 封装单元格坐标相关的所有操作，统一管理坐标转换
 */
class TXCoordinate {
public:
    // ==================== 构造函数 ====================
    
    /**
     * @brief 默认构造函数 - A1坐标
     */
    TXCoordinate() : row_(1), col_(1) {}
    
    /**
     * @brief 从行列坐标构造
     * @param row 行索引 (1-based)
     * @param col 列索引 (1-based)
     */
    TXCoordinate(TXTypes::RowIndex row, TXTypes::ColIndex col) : row_(row), col_(col) {}
    
    /**
     * @brief 从A1格式地址构造
     * @param address A1格式地址 (如 "A1", "B5", "AA10")
     */
    explicit TXCoordinate(const std::string& address);
    
    // ==================== 访问器 ====================
    
    /**
     * @brief 获取行索引
     * @return 行索引 (1-based)
     */
    TXTypes::RowIndex getRow() const { return row_; }
    
    /**
     * @brief 获取列索引
     * @return 列索引 (1-based)
     */
    TXTypes::ColIndex getCol() const { return col_; }
    
    /**
     * @brief 获取坐标对
     * @return {row, col}
     */
    std::pair<TXTypes::RowIndex, TXTypes::ColIndex> getPair() const { return {row_, col_}; }
    
    // ==================== 设置器 ====================
    
    /**
     * @brief 设置行索引
     * @param row 行索引 (1-based)
     * @return 自身引用，支持链式调用
     */
    TXCoordinate& setRow(TXTypes::RowIndex row);
    
    /**
     * @brief 设置列索引
     * @param col 列索引 (1-based)
     * @return 自身引用，支持链式调用
     */
    TXCoordinate& setCol(TXTypes::ColIndex col);
    
    /**
     * @brief 设置坐标
     * @param row 行索引 (1-based)
     * @param col 列索引 (1-based)
     * @return 自身引用，支持链式调用
     */
    TXCoordinate& set(TXTypes::RowIndex row, TXTypes::ColIndex col);
    
    /**
     * @brief 从A1格式地址设置坐标
     * @param address A1格式地址
     * @return 自身引用，支持链式调用
     */
    TXCoordinate& setFromAddress(const std::string& address);
    
    // ==================== 验证方法 ====================
    
    /**
     * @brief 检查坐标是否有效
     * @return 有效返回true，无效返回false
     */
    bool isValid() const;
    
    /**
     * @brief 检查行索引是否有效
     * @return 有效返回true，无效返回false
     */
    bool isValidRow() const;
    
    /**
     * @brief 检查列索引是否有效
     * @return 有效返回true，无效返回false
     */
    bool isValidCol() const;
    
    // ==================== 转换方法 ====================
    
    /**
     * @brief 转换为A1格式地址
     * @return A1格式地址 (如 "A1", "B5", "AA10")
     */
    std::string toAddress() const;
    
    /**
     * @brief 获取列名
     * @return Excel列名 (如 "A", "B", "AA")
     */
    std::string getColName() const;
    
    /**
     * @brief 转换为字符串表示
     * @return 字符串表示 (等同于toAddress())
     */
    std::string toString() const { return toAddress(); }
    
    // ==================== 偏移操作 ====================
    
    /**
     * @brief 行偏移
     * @param offset 偏移量 (可为负数)
     * @return 新的坐标对象
     */
    TXCoordinate offsetRow(int offset) const;
    
    /**
     * @brief 列偏移
     * @param offset 偏移量 (可为负数)
     * @return 新的坐标对象
     */
    TXCoordinate offsetCol(int offset) const;
    
    /**
     * @brief 坐标偏移
     * @param row_offset 行偏移量
     * @param col_offset 列偏移量
     * @return 新的坐标对象
     */
    TXCoordinate offset(int row_offset, int col_offset) const;
    
    /**
     * @brief 移动到下一行
     * @return 新的坐标对象
     */
    TXCoordinate nextRow() const { return offsetRow(1); }
    
    /**
     * @brief 移动到上一行
     * @return 新的坐标对象
     */
    TXCoordinate prevRow() const { return offsetRow(-1); }
    
    /**
     * @brief 移动到下一列
     * @return 新的坐标对象
     */
    TXCoordinate nextCol() const { return offsetCol(1); }
    
    /**
     * @brief 移动到上一列
     * @return 新的坐标对象
     */
    TXCoordinate prevCol() const { return offsetCol(-1); }
    
    // ==================== 比较操作 ====================
    
    bool operator==(const TXCoordinate& other) const;
    bool operator!=(const TXCoordinate& other) const { return !(*this == other); }
    bool operator<(const TXCoordinate& other) const;
    bool operator<=(const TXCoordinate& other) const;
    bool operator>(const TXCoordinate& other) const;
    bool operator>=(const TXCoordinate& other) const;
    
    // ==================== 运算符重载 ====================
    
    /**
     * @brief 加法运算符 (坐标偏移)
     * @param other 另一个坐标 (作为偏移量)
     * @return 偏移后的坐标
     */
    TXCoordinate operator+(const TXCoordinate& other) const;
    
    /**
     * @brief 减法运算符 (坐标偏移)
     * @param other 另一个坐标 (作为偏移量)
     * @return 偏移后的坐标
     */
    TXCoordinate operator-(const TXCoordinate& other) const;
    
    /**
     * @brief 加法赋值运算符
     * @param other 另一个坐标 (作为偏移量)
     * @return 自身引用
     */
    TXCoordinate& operator+=(const TXCoordinate& other);
    
    /**
     * @brief 减法赋值运算符
     * @param other 另一个坐标 (作为偏移量)
     * @return 自身引用
     */
    TXCoordinate& operator-=(const TXCoordinate& other);
    
    // ==================== 静态工厂方法 ====================
    
    /**
     * @brief 从A1格式地址创建坐标
     * @param address A1格式地址 (如 "A1", "B5", "AA10")
     * @return 坐标对象，解析失败返回A1坐标
     */
    static TXCoordinate fromAddress(const std::string& address);
    
    /**
     * @brief 从行列创建坐标
     * @param row 行索引 (1-based)
     * @param col 列索引 (1-based)
     * @return 坐标对象
     */
    static TXCoordinate fromRowCol(TXTypes::RowIndex row, TXTypes::ColIndex col);
    
    /**
     * @brief 从列名和行号创建坐标
     * @param col_name 列名 (如 "A", "B", "AA")
     * @param row 行号
     * @return 坐标对象
     */
    static TXCoordinate fromColNameRow(const std::string& col_name, TXTypes::RowIndex row);
    
    // ==================== 静态工具方法 ====================
    
    /**
     * @brief 列索引转换为Excel列名
     * @param col 列索引 (1-based)
     * @return Excel列名
     */
    static std::string colIndexToName(TXTypes::ColIndex col);
    
    /**
     * @brief Excel列名转换为列索引
     * @param name Excel列名
     * @return 列索引 (1-based)，无效返回INVALID_COL
     */
    static TXTypes::ColIndex colNameToIndex(const std::string& name);
    
    /**
     * @brief 坐标转换为A1格式地址
     * @param row 行索引 (1-based)
     * @param col 列索引 (1-based)
     * @return A1格式地址
     */
    static std::string coordinateToAddress(TXTypes::RowIndex row, TXTypes::ColIndex col);
    
    /**
     * @brief A1格式地址转换为坐标
     * @param address A1格式地址
     * @return 坐标对 {row, col}，无效返回 {INVALID_ROW, INVALID_COL}
     */
    static std::pair<TXTypes::RowIndex, TXTypes::ColIndex> addressToCoordinate(const std::string& address);
    
    /**
     * @brief 检查坐标是否有效
     * @param row 行索引
     * @param col 列索引
     * @return 有效返回true，无效返回false
     */
    static bool isValidCoordinate(TXTypes::RowIndex row, TXTypes::ColIndex col);

private:
    TXTypes::RowIndex row_;
    TXTypes::ColIndex col_;
};

/**
 * @brief 坐标范围类
 */
class TXRange {
public:
    // ==================== 构造函数 ====================
    
    /**
     * @brief 默认构造函数 - A1:A1范围
     */
    TXRange() : start_(1, 1), end_(1, 1) {}
    
    /**
     * @brief 从起始和结束坐标构造
     * @param start 起始坐标
     * @param end 结束坐标
     */
    TXRange(const TXCoordinate& start, const TXCoordinate& end);
    
    /**
     * @brief 从行列坐标构造
     * @param start_row 起始行
     * @param start_col 起始列
     * @param end_row 结束行
     * @param end_col 结束列
     */
    TXRange(TXTypes::RowIndex start_row, TXTypes::ColIndex start_col,
            TXTypes::RowIndex end_row, TXTypes::ColIndex end_col);
    
    /**
     * @brief 从A1格式范围地址构造
     * @param range_address 范围地址 (如 "A1:B5", "A1:A1")
     */
    explicit TXRange(const std::string& range_address);
    
    // ==================== 访问器 ====================
    
    /**
     * @brief 获取起始坐标
     * @return 起始坐标
     */
    const TXCoordinate& getStart() const { return start_; }
    
    /**
     * @brief 获取结束坐标
     * @return 结束坐标
     */
    const TXCoordinate& getEnd() const { return end_; }
    
    /**
     * @brief 获取行数
     * @return 行数
     */
    TXTypes::RowIndex getRowCount() const;
    
    /**
     * @brief 获取列数
     * @return 列数
     */
    TXTypes::ColIndex getColCount() const;
    
    /**
     * @brief 获取单元格总数
     * @return 单元格总数
     */
    uint64_t getCellCount() const;
    
    // ==================== 设置器 ====================
    
    /**
     * @brief 设置起始坐标
     * @param start 起始坐标
     * @return 自身引用，支持链式调用
     */
    TXRange& setStart(const TXCoordinate& start);
    
    /**
     * @brief 设置结束坐标
     * @param end 结束坐标
     * @return 自身引用，支持链式调用
     */
    TXRange& setEnd(const TXCoordinate& end);
    
    /**
     * @brief 设置范围
     * @param start 起始坐标
     * @param end 结束坐标
     * @return 自身引用，支持链式调用
     */
    TXRange& set(const TXCoordinate& start, const TXCoordinate& end);
    
    // ==================== 验证和操作 ====================
    
    /**
     * @brief 检查范围是否有效
     * @return 有效返回true，无效返回false
     */
    bool isValid() const;
    
    /**
     * @brief 检查坐标是否在范围内
     * @param coord 坐标
     * @return 在范围内返回true，否则返回false
     */
    bool contains(const TXCoordinate& coord) const;
    
    /**
     * @brief 检查另一个范围是否完全在当前范围内
     * @param other 另一个范围
     * @return 完全包含返回true，否则返回false
     */
    bool contains(const TXRange& other) const;
    
    /**
     * @brief 检查两个范围是否有重叠
     * @param other 另一个范围
     * @return 有重叠返回true，否则返回false
     */
    bool intersects(const TXRange& other) const;
    
    /**
     * @brief 获取与另一个范围的交集
     * @param other 另一个范围
     * @return 交集范围，无交集返回无效范围
     */
    TXRange intersection(const TXRange& other) const;
    
    /**
     * @brief 扩展范围以包含指定坐标
     * @param coord 坐标
     * @return 自身引用，支持链式调用
     */
    TXRange& expand(const TXCoordinate& coord);
    
    /**
     * @brief 扩展范围以包含另一个范围
     * @param other 另一个范围
     * @return 自身引用，支持链式调用
     */
    TXRange& expand(const TXRange& other);
    
    // ==================== 转换方法 ====================
    
    /**
     * @brief 转换为A1格式范围地址
     * @return A1格式范围地址 (如 "A1:B5")
     */
    std::string toAddress() const;
    
    /**
     * @brief 获取范围内的所有坐标
     * @return 坐标向量
     */
    std::vector<TXCoordinate> getAllCoordinates() const;
    
    // ==================== 比较操作 ====================
    
    bool operator==(const TXRange& other) const;
    bool operator!=(const TXRange& other) const { return !(*this == other); }
    
    // ==================== 静态工厂方法 ====================
    
    /**
     * @brief 从A1格式范围地址创建范围
     * @param range_address 范围地址 (如 "A1:B5")
     * @return 范围对象
     */
    static TXRange fromAddress(const std::string& range_address);
    
    /**
     * @brief 创建单个单元格范围
     * @param coord 坐标
     * @return 单元格范围
     */
    static TXRange singleCell(const TXCoordinate& coord);
    
    /**
     * @brief 创建整行范围
     * @param row 行号
     * @return 行范围
     */
    static TXRange entireRow(TXTypes::RowIndex row);
    
    /**
     * @brief 创建整列范围
     * @param col 列号
     * @return 列范围
     */
    static TXRange entireCol(TXTypes::ColIndex col);

private:
    TXCoordinate start_;
    TXCoordinate end_;
    
    void normalize(); // 确保start <= end
};

} // namespace TinaXlsx 