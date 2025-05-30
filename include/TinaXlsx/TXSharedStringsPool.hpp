//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include <unordered_map>
#include <vector>
#include "TXTypes.hpp"

namespace TinaXlsx
{
    class TXSharedStringsPool {
    public:
        // 添加字符串并返回索引
        u32 add(const std::string& str) {
            // 尝试在唯一字符串表中查找
            if (auto it = m_uniqueIndexMap.find(str); it != m_uniqueIndexMap.end()) {
                m_frequencyMap[str]++;
                return it->second;
            }
        
            // 新字符串 - 添加到池
            const u32 newIndex = static_cast<u32>(m_strings.size());
            m_strings.push_back(str);
            m_uniqueIndexMap[str] = newIndex;
            m_frequencyMap[str] = 1;
            m_dirty = true;
        
            return newIndex;
        }
    
        // 获取所有字符串（按添加顺序）
        [[nodiscard]] const std::vector<std::string>& getStrings() const { 
            return m_strings; 
        }
    
        // 检查是否需要写入XML
        [[nodiscard]] bool isDirty() const { 
            return m_dirty; 
        }
    
        // 重置状态
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
