//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include <vector>
#include <memory>
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXStyleManager.hpp"
#include "TinaXlsx/TXComponentManager.hpp"
#include "TinaXlsx/TXSharedStringsPool.hpp"

namespace TinaXlsx
{
    struct TXWorkbookContext
    {
        // 工作表列表        
        std::vector<std::unique_ptr<TXSheet>>& sheets;
        // 样式管理器
        TXStyleManager& styleManager;
        // 组件管理器
        ComponentManager& componentManager;
        std::atomic_flag componentDetectionFlag = ATOMIC_FLAG_INIT;

        // 共享字符串池
        TXSharedStringsPool& sharedStringsPool;

        // 构造函数
        TXWorkbookContext(std::vector<std::unique_ptr<TXSheet>>& sheets_ref,
                         TXStyleManager& style_manager_ref,
                         ComponentManager& component_manager_ref,
                         TXSharedStringsPool& shared_strings_pool_ref)
            : sheets(sheets_ref)
            , styleManager(style_manager_ref)
            , componentManager(component_manager_ref)
            , sharedStringsPool(shared_strings_pool_ref) {
        }

        void registerComponentFast(ExcelComponent component)
        {
            if (!componentManager.hasComponent(component))
            {
                componentManager.registerComponent(component);
            }
        }
    };
}
