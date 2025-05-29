//
// Created by wuxianggujun on 2025/5/29.
//

#include "TinaXlsx/TXWorksheetXmlHandler.hpp"
#include "TinaXlsx/TXCell.hpp"
#include <variant>

namespace TinaXlsx
{
    XmlNodeBuilder TXWorksheetXmlHandler::buildCellNode(const TXCell* cell, const std::string& cellRef) const
    {
        XmlNodeBuilder cellNode("c");
        cellNode.addAttribute("r", cellRef);

        if (cell)
        {
            u32 styleIndex = cell->getStyleIndex();
            if (styleIndex != 0)
            {
                cellNode.addAttribute("s", std::to_string(styleIndex));
            }
            
            // 根据单元格类型设置属性和值
            cell_value_t value = cell->getValue();
            
            if (std::holds_alternative<std::string>(value)) {
                // 字符串类型
                cellNode.addAttribute("t", "inlineStr");
                
                XmlNodeBuilder isNode("is");
                XmlNodeBuilder tNode("t");
                tNode.setText(std::get<std::string>(value));
                isNode.addChild(tNode);
                cellNode.addChild(isNode);
                
            } else if (std::holds_alternative<double>(value)) {                // 浮点数类型
                XmlNodeBuilder vNode("v");
                vNode.setText(std::to_string(std::get<double>(value)));
                cellNode.addChild(vNode);
                
            } else if (std::holds_alternative<int64_t>(value)) {
                // 整数类型
                XmlNodeBuilder vNode("v");
                vNode.setText(std::to_string(std::get<int64_t>(value)));
                cellNode.addChild(vNode);
                
            } else if (std::holds_alternative<bool>(value)) {
                // 布尔类型
                cellNode.addAttribute("t", "b");
                XmlNodeBuilder vNode("v");
                vNode.setText(std::get<bool>(value) ? "1" : "0");
                cellNode.addChild(vNode);
            }
        }
        
        return cellNode;
    }
}