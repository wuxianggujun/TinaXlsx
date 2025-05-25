#pragma once

#include "TXTypes.hpp"
#include <string>
#include <algorithm>

namespace TinaXlsx {

/**
 * @brief 列索引类
 * 
 * 封装Excel列操作的专用类，提供列号和列名之间的转换以及各种操作
 */
class TXColumn {
public:
    /**
     * @brief 默认构造函数
     * @param index 列索引（从1开始，默认为1）
     */
    explicit TXColumn(TXTypes::ColIndex index = 1) : index_(index) {}
    
    /**
     * @brief 从列名构造
     * @param name 列名（如"A", "B", "AA"）
     */
    explicit TXColumn(const std::string& name) : index_(colNameToIndex(name)) {}
    
    /**
     * @brief 获取列索引
     * @return 列索引
     */
    TXTypes::ColIndex getIndex() const { return index_; }
    
    /**
     * @brief 设置列索引
     * @param index 列索引
     */
    void setIndex(TXTypes::ColIndex index) { index_ = index; }
    
    /**
     * @brief 检查列索引是否有效
     * @return 有效返回true，否则返回false
     */
    bool isValid() const { return TXTypes::isValidCol(index_); }
    
    /**
     * @brief 下一列
     * @return 下一列对象
     */
    TXColumn next() const { return TXColumn(index_ + 1); }
    
    /**
     * @brief 上一列
     * @return 上一列对象
     */
    TXColumn previous() const { return index_ > 1 ? TXColumn(index_ - 1) : TXColumn(1); }
    
    /**
     * @brief 偏移指定列数
     * @param offset 偏移量（可以为负数）
     * @return 偏移后的列对象
     */
    TXColumn offset(int offset) const {
        int new_index = static_cast<int>(index_) + offset;
        return TXColumn(std::max(1, new_index));
    }
    
    /**
     * @brief 获取列名
     * @return 列名字符串（如"A", "B", "AA"）
     */
    std::string getName() const { return colIndexToName(index_); }
    
    /**
     * @brief 转换为字符串
     * @return 列名字符串
     */
    std::string toString() const { return getName(); }
    
    // ==================== 比较操作符 ====================
    bool operator==(const TXColumn& other) const { return index_ == other.index_; }
    bool operator!=(const TXColumn& other) const { return index_ != other.index_; }
    bool operator<(const TXColumn& other) const { return index_ < other.index_; }
    bool operator<=(const TXColumn& other) const { return index_ <= other.index_; }
    bool operator>(const TXColumn& other) const { return index_ > other.index_; }
    bool operator>=(const TXColumn& other) const { return index_ >= other.index_; }
    
    // ==================== 算术操作符 ====================
    TXColumn operator+(int offset) const { return this->offset(offset); }
    TXColumn operator-(int offset) const { return this->offset(-offset); }
    TXColumn& operator++() { ++index_; return *this; }
    TXColumn operator++(int) { TXColumn temp(*this); ++index_; return temp; }
    TXColumn& operator--() { if (index_ > 1) --index_; return *this; }
    TXColumn operator--(int) { TXColumn temp(*this); if (index_ > 1) --index_; return temp; }
    
    // ==================== 静态方法 ====================
    
    /**
     * @brief 列名转索引
     * @param name 列名（如"A", "B", "AA"）
     * @return 列索引，无效时返回0
     */
    static TXTypes::ColIndex colNameToIndex(const std::string& name);
    
    /**
     * @brief 索引转列名
     * @param index 列索引
     * @return 列名，无效时返回空字符串
     */
    static std::string colIndexToName(TXTypes::ColIndex index);
    
    // ==================== 静态工厂方法 ====================
    
    /**
     * @brief 创建第一列（A列）
     * @return 第一列对象
     */
    static TXColumn first() { return TXColumn(1); }
    
    /**
     * @brief 创建最后一列（Excel限制内）
     * @return 最后一列对象
     */
    static TXColumn last() { return TXColumn(TXTypes::MAX_COLS); }
    
    /**
     * @brief 从列名创建列对象
     * @param name 列名
     * @return 列对象
     */
    static TXColumn fromName(const std::string& name) { return TXColumn(name); }
    
private:
    TXTypes::ColIndex index_;
};

} // namespace TinaXlsx 