//
// Created by wuxianggujun on 2025/5/29.
//

#pragma once

#include <vector>
#include <memory>
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXStyleManager.hpp"
#include "TinaXlsx/TXComponentManager.hpp"

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
    };
}
