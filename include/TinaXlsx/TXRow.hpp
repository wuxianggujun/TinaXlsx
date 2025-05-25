#pragma once

#include "TXTypes.hpp"
#include <string>
#include <algorithm>

namespace TinaXlsx {

/**
 * @brief 行索引类
 * 
 * 封装Excel行操作的专用类，提供行号的各种操作和验证
 */
class TXRow {
public:
    /**
     * @brief 默认构造函数
     * @param index 行索引（从1开始，默认为1）
     */
    explicit TXRow(TXTypes::RowIndex index = 1) : index_(index) {}
    
    /**
     * @brief 获取行索引
     * @return 行索引
     */
    TXTypes::RowIndex getIndex() const { return index_; }
    
    /**
     * @brief 设置行索引
     * @param index 行索引
     */
    void setIndex(TXTypes::RowIndex index) { index_ = index; }
    
    /**
     * @brief 检查行索引是否有效
     * @return 有效返回true，否则返回false
     */
    bool isValid() const { return TXTypes::isValidRow(index_); }
    
    /**
     * @brief 下一行
     * @return 下一行对象
     */
    TXRow next() const { return TXRow(index_ + 1); }
    
    /**
     * @brief 上一行
     * @return 上一行对象
     */
    TXRow previous() const { return index_ > 1 ? TXRow(index_ - 1) : TXRow(1); }
    
    /**
     * @brief 偏移指定行数
     * @param offset 偏移量（可以为负数）
     * @return 偏移后的行对象
     */
    TXRow offset(int offset) const {
        int new_index = static_cast<int>(index_) + offset;
        return TXRow(std::max(1, new_index));
    }
    
    /**
     * @brief 转换为字符串
     * @return 行号字符串
     */
    std::string toString() const { return std::to_string(index_); }
    
    // ==================== 比较操作符 ====================
    bool operator==(const TXRow& other) const { return index_ == other.index_; }
    bool operator!=(const TXRow& other) const { return index_ != other.index_; }
    bool operator<(const TXRow& other) const { return index_ < other.index_; }
    bool operator<=(const TXRow& other) const { return index_ <= other.index_; }
    bool operator>(const TXRow& other) const { return index_ > other.index_; }
    bool operator>=(const TXRow& other) const { return index_ >= other.index_; }
    
    // ==================== 算术操作符 ====================
    TXRow operator+(int offset) const { return this->offset(offset); }
    TXRow operator-(int offset) const { return this->offset(-offset); }
    TXRow& operator++() { ++index_; return *this; }
    TXRow operator++(int) { TXRow temp(*this); ++index_; return temp; }
    TXRow& operator--() { if (index_ > 1) --index_; return *this; }
    TXRow operator--(int) { TXRow temp(*this); if (index_ > 1) --index_; return temp; }
    
    // ==================== 静态工厂方法 ====================
    
    /**
     * @brief 创建第一行
     * @return 第一行对象
     */
    static TXRow first() { return TXRow(1); }
    
    /**
     * @brief 创建最后一行（Excel限制内）
     * @return 最后一行对象
     */
    static TXRow last() { return TXRow(TXTypes::MAX_ROWS); }
    
private:
    TXTypes::RowIndex index_;
};

} // namespace TinaXlsx 