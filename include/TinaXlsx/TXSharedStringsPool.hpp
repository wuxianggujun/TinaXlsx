//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include <unordered_map>
#include <vector>
#include <string>
#include "TXTypes.hpp"
#include "TXGlobalStringPool.hpp"

namespace TinaXlsx
{
    class TXSharedStringsPool {
    public:
        // 性能优化：添加字符串并返回索引
        u32 add(const std::string& str) {

            // 优化：使用单次查找避免重复哈希计算
            auto [it, inserted] = m_uniqueIndexMap.try_emplace(str, static_cast<u32>(m_strings.size()));

            if (inserted) {
                // 新字符串 - 添加到池
                m_strings.push_back(str);
                m_frequencyMap[str] = 1;
                m_dirty = true;
                return it->second;
            } else {
                // 优化：已存在的字符串，直接增加频率（避免再次查找）
                m_frequencyMap[str]++;
                return it->second;
            }
        }
    
        // 获取所有字符串（按添加顺序）
        [[nodiscard]] const std::vector<std::string>& getStrings() const { 
            return m_strings; 
        }
    
        // 检查是否需要写入XML
        [[nodiscard]] bool isDirty() const { 
            return m_dirty; 
        }
    
        // 🚀 重置状态
        void reset() {
            m_strings.clear();
            m_uniqueIndexMap.clear();
            m_frequencyMap.clear();
            m_dirty = false;
        }

    private:
        std::vector<std::string> m_strings;              // 按顺序存储的字符串
        std::unordered_map<std::string, u32> m_uniqueIndexMap; // 字符串到索引的映射
        std::unordered_map<std::string, int> m_frequencyMap;   // 字符串频率统计
        bool m_dirty = false;                            // 是否有未保存更改
    };
}
