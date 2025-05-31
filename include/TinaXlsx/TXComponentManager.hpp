#pragma once

#include "TinaXlsx/TXTypes.hpp"
#include <unordered_set>
#include <vector>
#include <string>

namespace TinaXlsx {

class TXSheet;

/**
 * @brief Excel功能组件枚举
 */
enum class ExcelComponent {
    BasicWorkbook,      // 基础工作簿（必需）
    SharedStrings,      // 共享字符串
    Styles,            // 样式
    MergedCells,       // 合并单元格
    NumberFormats,     // 数字格式
    Formulas,          // 公式
    Charts,            // 图表
    PivotTables,       // 数据透视表
    Comments,          // 批注
    Hyperlinks,        // 超链接
    Images,            // 图片
    DataValidation,    // 数据验证
    ConditionalFormat, // 条件格式
    DocumentProperties // 文档属性
};

/**
 * @brief 组件管理器 - 负责管理Excel文件的各个功能组件
 * 
 * 简化版本，删除了ComponentGenerator，XML生成由Handler负责
 */
class ComponentManager {
public:
    /**
     * @brief 注册需要的组件
     * @param component 组件类型
     */
    void registerComponent(ExcelComponent component);
    
    /**
     * @brief 检查是否包含某个组件
     * @param component 组件类型
     * @return 包含返回true
     */
    bool hasComponent(ExcelComponent component) const;
    
    /**
     * @brief 获取所有已注册的组件
     * @return 组件集合
     */
    const std::unordered_set<ExcelComponent>& getComponents() const;
    
    /**
     * @brief 重置组件注册
     */
    void reset();

private:
    std::unordered_set<ExcelComponent> registered_components_;
};

} // namespace TinaXlsx 
