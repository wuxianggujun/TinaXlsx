//
// @file TXGlobalStringPool.hpp
// @brief 全局字符串池 - 字符串内化优化
//

#pragma once

#include <string>
#include <unordered_set>
#include <mutex>
#include <string_view>
#include <vector>

namespace TinaXlsx {

/**
 * @brief 全局字符串池 - 实现字符串内化优化
 * 
 * 字符串内化将相同内容的字符串存储在全局池中，
 * 避免重复存储，提高内存效率和比较性能
 */
class TXGlobalStringPool {
public:
    /**
     * @brief 获取全局字符串池实例
     */
    static TXGlobalStringPool& instance();

    /**
     * @brief 字符串内化 - 将字符串加入全局池
     * @param str 要内化的字符串
     * @return 内化后的字符串引用（指向池中的字符串）
     */
    const std::string& intern(const std::string& str);
    
    /**
     * @brief 字符串内化 - string_view版本（避免临时对象）
     */
    const std::string& intern(std::string_view str);

    /**
     * @brief 检查字符串是否已经内化
     */
    bool isInterned(const std::string& str) const;

    /**
     * @brief 获取池中字符串数量
     */
    size_t size() const;

    /**
     * @brief 清空池（谨慎使用）
     */
    void clear();

    /**
     * @brief 获取所有字符串
     * @return 池中所有字符串的向量
     */
    std::vector<std::string> getAllStrings() const;

    /**
     * @brief 根据索引获取字符串
     * @param index 字符串索引
     * @return 字符串引用
     */
    const std::string& getString(size_t index) const;

    /**
     * @brief 添加字符串到池中 (addString方法，等价于intern)
     * @param str 要添加的字符串
     * @return 内化后的字符串引用
     */
    const std::string& addString(const std::string& str);

    /**
     * @brief 获取字符串的索引
     * @param str 要查找的字符串
     * @return 字符串在池中的索引，如果不存在则返回SIZE_MAX
     */
    size_t getIndex(const std::string& str) const;
    


    // 🚀 常用字符串常量 - 预内化的高频字符串
    static const std::string& EMPTY_STRING();
    static const std::string& DEFAULT_SHEET_NAME();
    static const std::string& TRUE_STRING();
    static const std::string& FALSE_STRING();
    static const std::string& ZERO_STRING();
    static const std::string& ONE_STRING();
    
    // XML相关常量
    static const std::string& XML_DECLARATION();
    static const std::string& WORKSHEET_NAMESPACE();
    static const std::string& RELATIONSHIPS_NAMESPACE();

private:
    TXGlobalStringPool() = default;
    ~TXGlobalStringPool() = default;
    
    // 禁止拷贝和移动
    TXGlobalStringPool(const TXGlobalStringPool&) = delete;
    TXGlobalStringPool& operator=(const TXGlobalStringPool&) = delete;

    mutable std::mutex mutex_;
    std::unordered_set<std::string> pool_;
    
    // 预内化常用字符串
    void initializeCommonStrings();
};

/**
 * @brief 字符串内化辅助函数
 */
inline const std::string& intern(const std::string& str) {
    return TXGlobalStringPool::instance().intern(str);
}

inline const std::string& intern(std::string_view str) {
    return TXGlobalStringPool::instance().intern(str);
}

} // namespace TinaXlsx
