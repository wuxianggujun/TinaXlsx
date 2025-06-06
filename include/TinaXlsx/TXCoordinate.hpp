#pragma once

#include "TXTypes.hpp"
#include <string>

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
    TXCoordinate(const row_t& row, const column_t& col) : row_(row), col_(col) {}
    
    /**
     * @brief 从A1格式地址构造
     * @param address A1格式地址 (如 "A1", "B5", "AA10")
     */
    explicit TXCoordinate(const std::string& address);
    
    // ==================== 访问器 ====================
    
    /**
     * @brief 获取行对象
     * @return 行对象
     */
    const row_t& getRow() const { return row_; }
    
    /**
     * @brief 获取列对象
     * @return 列对象
     */
    const column_t& getCol() const { return col_; }
    
    /**
     * @brief 获取列对象（兼容性接口）
     * @return 列对象
     */
    const column_t& getColumn() const { return col_; }
    
    /**
     * @brief 获取坐标对
     * @return {row, col}
     */
    std::pair<row_t, column_t> getPair() const { return {row_, col_}; }
    

    
    // ==================== 设置器 ====================
    
    /**
     * @brief 设置行
     * @param row 行对象
     * @return 自身引用，支持链式调用
     */
    TXCoordinate& setRow(const row_t& row);
    
    /**
     * @brief 设置列
     * @param col 列对象
     * @return 自身引用，支持链式调用
     */
    TXCoordinate& setCol(const column_t& col);
    
    /**
     * @brief 设置坐标
     * @param row 行对象
     * @param col 列对象
     * @return 自身引用，支持链式调用
     */
    TXCoordinate& set(const row_t& row, const column_t& col);
    
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
     * @return 坐标对象
     */
    static TXCoordinate fromAddress(const std::string& address);
    

    
    /**
     * @brief 从列名和行创建坐标
     * @param col_name 列名 (如 "A", "B", "AA")
     * @param row 行对象
     * @return 坐标对象
     */
    static TXCoordinate fromColNameRow(const std::string& col_name, const row_t& row);
    
    // ==================== 静态工具方法 ====================
    
    /**
     * @brief 检查坐标是否有效
     * @param row 行对象
     * @param col 列对象
     * @return 有效返回true，无效返回false
     */
    static bool isValidCoordinate(const row_t& row, const column_t& col);

private:
    row_t row_;
    column_t col_;
};

} // namespace TinaXlsx

// 为TXCoordinate提供std::hash支持
namespace std {
    template<>
    struct hash<TinaXlsx::TXCoordinate> {
        std::size_t operator()(const TinaXlsx::TXCoordinate& coord) const {
            // 使用行列索引计算哈希值
            uint64_t combined = (static_cast<uint64_t>(coord.getRow().index()) << 32) |
                               coord.getCol().index();
            return std::hash<uint64_t>()(combined);
        }
    };
}