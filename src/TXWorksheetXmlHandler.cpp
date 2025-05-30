//
// Created by wuxianggujun on 2025/5/29.
//

#include "TinaXlsx/TXWorksheetXmlHandler.hpp"
#include "TinaXlsx/TXCell.hpp"
#include <variant>

#include "TinaXlsx/TXSharedStringsPool.hpp"

namespace TinaXlsx
{
    bool TXWorksheetXmlHandler::shouldUseInlineString(const std::string& str) const
    {
        // 策略1: 极短字符串（1个字符）使用内联，节省共享字符串池空间
        if (str.length() <= 1) return true;
    
        // 策略2: 包含特殊XML字符的字符串使用内联（避免XML转义复杂性）
        if (str.find_first_of("<>&\"'") != std::string::npos) return true;
    
        // 策略3: 包含控制字符的字符串使用内联（避免XML解析问题）
        if (str.find_first_of("\n\r\t") != std::string::npos) return true;
        
        // 策略4: 非常长的字符串（>100字符）使用内联（避免共享字符串池膨胀）
        if (str.length() > 100) return true;
    
        // 策略5: 2-100字符的普通字符串使用共享字符串（最大化复用效果）
        return false;
    }

    XmlNodeBuilder TXWorksheetXmlHandler::buildCellNode(const TXCell* cell, const std::string& cellRef,
                                                        const TXWorkbookContext& context) const
    {
        XmlNodeBuilder cellNode("c");
        cellNode.addAttribute("r", cellRef);

        if (!cell) return cellNode;

        // 处理样式
        if (u32 styleIndex = cell->getStyleIndex(); styleIndex != 0)
        {
            cellNode.addAttribute("s", std::to_string(styleIndex));
        }
        
        // 获取单元格值和类型
        const cell_value_t& value = cell->getValue();
        const TXCell::CellType cellType = cell->getType();
        
        if (cellType == TXCell::CellType::String)
        {
            const std::string& str = cell->getStringValue();
            
            // 根据字符串长度和内容决定使用内联还是共享
            if (shouldUseInlineString(str)) {
                // 内联字符串 - 直接嵌入XML
                cellNode.addAttribute("t", "inlineStr");
                XmlNodeBuilder isNode("is");
                XmlNodeBuilder tNode("t");
                tNode.setText(str);
                isNode.addChild(tNode);
                cellNode.addChild(isNode);
            } else {
                // 共享字符串 - 添加到池
                u32 index = context.sharedStringsPool.add(str);
            
                cellNode.addAttribute("t", "s");
                cellNode.addChild(XmlNodeBuilder("v").setText(std::to_string(index)));
            }
        }
        else if (std::holds_alternative<double>(value))
        {
            // 浮点数类型
            XmlNodeBuilder vNode("v");
            vNode.setText(std::to_string(std::get<double>(value)));
            cellNode.addChild(vNode);
        }
        else if (std::holds_alternative<int64_t>(value))
        {
            // 整数类型
            XmlNodeBuilder vNode("v");
            vNode.setText(std::to_string(std::get<int64_t>(value)));
            cellNode.addChild(vNode);
        }
        else if (std::holds_alternative<bool>(value))
        {
            // 布尔类型
            cellNode.addAttribute("t", "b");
            XmlNodeBuilder vNode("v");
            vNode.setText(std::get<bool>(value) ? "1" : "0");
            cellNode.addChild(vNode);
        }


        return cellNode;
    }
}
