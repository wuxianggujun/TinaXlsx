#pragma once

#include "TXTypes.hpp"
#include <unordered_set>
#include <vector>
#include <string>
#include <utility>

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
     * @brief 自动检测Sheet中使用的功能并注册相应组件
     * @param sheet 工作表指针
     */
    void autoDetectComponents(const TXSheet* sheet);
    
    /**
     * @brief 重置组件注册
     */
    void reset();

private:
    std::unordered_set<ExcelComponent> registered_components_;
};

/**
 * @brief 组件生成器 - 负责生成各个功能组件的XML内容
 */
class ComponentGenerator {
public:
    /**
     * @brief 生成Content-Types.xml内容
     * @param components 需要的组件集合
     * @param sheet_count 工作表数量
     * @return XML字符串
     */
    static std::string generateContentTypes(const std::unordered_set<ExcelComponent>& components, size_t sheet_count);
    
    /**
     * @brief 生成主关系文件内容
     * @param components 需要的组件集合
     * @return XML字符串
     */
    static std::string generateMainRelationships(const std::unordered_set<ExcelComponent>& components);
    
    /**
     * @brief 生成工作簿关系文件内容
     * @param components 需要的组件集合
     * @param sheet_count 工作表数量
     * @return XML字符串
     */
    static std::string generateWorkbookRelationships(const std::unordered_set<ExcelComponent>& components, size_t sheet_count);
    
    /**
     * @brief 生成样式文件内容
     * @return XML字符串
     */
    static std::string generateStyles();
    
    /**
     * @brief 生成共享字符串文件内容
     * @param strings 字符串集合
     * @return XML字符串
     */
    static std::string generateSharedStrings(const std::vector<std::string>& strings);
    
    /**
     * @brief 生成文档属性文件内容
     * @return 包含core.xml和app.xml的pair
     */
    static std::pair<std::string, std::string> generateDocumentProperties();
};

} // namespace TinaXlsx 