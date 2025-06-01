#pragma once

#include "TXTypes.hpp"
#include "TXCoordinate.hpp"
#include <string>
#include <vector>

namespace TinaXlsx {

/**
 * @brief 坐标范围类
 */
class TXRange {
public:
    // ==================== 构造函数 ====================
    
    /**
     * @brief 默认构造函数 - A1:A1范围
     */
    TXRange();
    
    /**
     * @brief 从两个坐标构造范围
     * @param start 起始坐标
     * @param end 结束坐标
     */
    TXRange(const TXCoordinate& start, const TXCoordinate& end);
    
    /**
     * @brief 从范围地址构造
     * @param range_address 范围地址 (如 "A1:B5", "A:A", "1:1")
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
    row_t getRowCount() const;
    
    /**
     * @brief 获取列数
     * @return 列数
     */
    column_t getColCount() const;
    
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
     * @brief 转换为绝对引用地址字符串
     * @return 绝对引用地址字符串（如"$A$1:$B$2"）
     */
    std::string toAbsoluteAddress() const;
    
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
     * @brief 从范围地址创建范围
     * @param range_address 范围地址
     * @return 范围对象
     */
    static TXRange fromAddress(const std::string& range_address);
    
    /**
     * @brief 创建单个单元格范围
     * @param coord 坐标
     * @return 范围对象
     */
    static TXRange singleCell(const TXCoordinate& coord);
    
    /**
     * @brief 创建整行范围
     * @param row 行对象
     * @return 范围对象
     */
    static TXRange entireRow(const row_t& row);
    
    /**
     * @brief 创建整列范围
     * @param col 列对象
     * @return 范围对象
     */
    static TXRange entireCol(const column_t& col);

private:
    TXCoordinate start_;
    TXCoordinate end_;
    
    void normalize(); // 确保start <= end
};

} // namespace TinaXlsx 