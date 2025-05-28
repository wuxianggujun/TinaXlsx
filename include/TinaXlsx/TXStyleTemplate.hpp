#pragma once

#include "TXTypes.hpp"
#include "TXStyle.hpp"
#include <string>
#include <vector>
#include <memory>
#include <unordered_map>

namespace TinaXlsx {

/**
 * @brief 样式模板类别枚举
 */
enum class StyleTemplateCategory : uint8_t {
    Header = 0,          ///< 标题样式
    Data = 1,            ///< 数据样式
    Number = 2,          ///< 数字样式
    Currency = 3,        ///< 货币样式
    Date = 4,            ///< 日期样式
    Highlight = 5,       ///< 高亮样式
    Table = 6,           ///< 表格样式
    Chart = 7,           ///< 图表样式
    Custom = 8           ///< 自定义样式
};

/**
 * @brief 样式模板信息
 */
struct StyleTemplateInfo {
    std::string id;               ///< 模板ID
    std::string name;             ///< 模板名称
    std::string description;      ///< 模板描述
    StyleTemplateCategory category; ///< 模板类别
    std::string author;           ///< 创建者
    std::string version;          ///< 版本
    bool isBuiltIn;              ///< 是否为内置模板
    
    StyleTemplateInfo() 
        : category(StyleTemplateCategory::Custom)
        , isBuiltIn(false) {}
        
    StyleTemplateInfo(const std::string& templateId, const std::string& templateName, StyleTemplateCategory cat = StyleTemplateCategory::Custom)
        : id(templateId), name(templateName), category(cat), isBuiltIn(false) {}
};

/**
 * @brief 样式模板类
 */
class TXStyleTemplate {
public:
    TXStyleTemplate();
    explicit TXStyleTemplate(const StyleTemplateInfo& info);
    ~TXStyleTemplate();
    
    // 支持拷贝和移动
    TXStyleTemplate(const TXStyleTemplate& other);
    TXStyleTemplate& operator=(const TXStyleTemplate& other);
    TXStyleTemplate(TXStyleTemplate&& other) noexcept;
    TXStyleTemplate& operator=(TXStyleTemplate&& other) noexcept;
    
    // ==================== 模板信息管理 ====================
    
    /**
     * @brief 获取模板信息
     * @return 模板信息
     */
    const StyleTemplateInfo& getInfo() const;
    
    /**
     * @brief 设置模板信息
     * @param info 模板信息
     */
    void setInfo(const StyleTemplateInfo& info);
    
    /**
     * @brief 获取模板ID
     * @return 模板ID
     */
    std::string getId() const;
    
    /**
     * @brief 获取模板名称
     * @return 模板名称
     */
    std::string getName() const;
    
    /**
     * @brief 设置模板名称
     * @param name 模板名称
     */
    void setName(const std::string& name);
    
    /**
     * @brief 获取模板描述
     * @return 模板描述
     */
    std::string getDescription() const;
    
    /**
     * @brief 设置模板描述
     * @param description 模板描述
     */
    void setDescription(const std::string& description);
    
    // ==================== 样式管理 ====================
    
    /**
     * @brief 设置基础样式
     * @param style 基础样式
     */
    void setBaseStyle(const TXCellStyle& style);
    
    /**
     * @brief 获取基础样式
     * @return 基础样式
     */
    const TXCellStyle& getBaseStyle() const;
    
    /**
     * @brief 添加命名样式
     * @param name 样式名称
     * @param style 样式
     */
    void addNamedStyle(const std::string& name, const TXCellStyle& style);
    
    /**
     * @brief 获取命名样式
     * @param name 样式名称
     * @return 样式指针，如果不存在返回nullptr
     */
    const TXCellStyle* getNamedStyle(const std::string& name) const;
    
    /**
     * @brief 移除命名样式
     * @param name 样式名称
     */
    void removeNamedStyle(const std::string& name);
    
    /**
     * @brief 获取所有命名样式的名称
     * @return 样式名称列表
     */
    std::vector<std::string> getNamedStyleNames() const;
    
    /**
     * @brief 清空所有命名样式
     */
    void clearNamedStyles();
    
    // ==================== 预设样式快捷方法 ====================
    
    /**
     * @brief 设置标题样式
     * @param level 标题级别（1-6）
     * @param style 样式
     */
    void setHeaderStyle(int level, const TXCellStyle& style);
    
    /**
     * @brief 获取标题样式
     * @param level 标题级别（1-6）
     * @return 样式指针
     */
    const TXCellStyle* getHeaderStyle(int level) const;
    
    /**
     * @brief 设置数据样式
     * @param style 样式
     */
    void setDataStyle(const TXCellStyle& style);
    
    /**
     * @brief 获取数据样式
     * @return 样式指针
     */
    const TXCellStyle* getDataStyle() const;
    
    /**
     * @brief 设置数字样式
     * @param style 样式
     */
    void setNumberStyle(const TXCellStyle& style);
    
    /**
     * @brief 获取数字样式
     * @return 样式指针
     */
    const TXCellStyle* getNumberStyle() const;
    
    // ==================== 序列化功能 ====================
    
    /**
     * @brief 序列化为JSON字符串
     * @return JSON字符串
     */
    std::string toJson() const;
    
    /**
     * @brief 从JSON字符串反序列化
     * @param json JSON字符串
     * @return 成功返回true
     */
    bool fromJson(const std::string& json);
    
    /**
     * @brief 克隆模板
     * @return 克隆的模板
     */
    TXStyleTemplate clone() const;

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

/**
 * @brief 样式模板管理器
 */
class TXStyleTemplateManager {
public:
    TXStyleTemplateManager();
    ~TXStyleTemplateManager();
    
    // 单例模式
    static TXStyleTemplateManager& getInstance();
    
    // ==================== 模板管理 ====================
    
    /**
     * @brief 注册模板
     * @param styleTemplate 样式模板
     * @return 成功返回true
     */
    bool registerTemplate(const TXStyleTemplate& styleTemplate);
    
    /**
     * @brief 注销模板
     * @param templateId 模板ID
     * @return 成功返回true
     */
    bool unregisterTemplate(const std::string& templateId);
    
    /**
     * @brief 获取模板
     * @param templateId 模板ID
     * @return 模板指针，如果不存在返回nullptr
     */
    const TXStyleTemplate* getTemplate(const std::string& templateId) const;
    
    /**
     * @brief 获取模板（可修改）
     * @param templateId 模板ID
     * @return 模板指针，如果不存在返回nullptr
     */
    TXStyleTemplate* getTemplate(const std::string& templateId);
    
    /**
     * @brief 检查模板是否存在
     * @param templateId 模板ID
     * @return 存在返回true
     */
    bool hasTemplate(const std::string& templateId) const;
    
    /**
     * @brief 获取所有模板ID
     * @return 模板ID列表
     */
    std::vector<std::string> getAllTemplateIds() const;
    
    /**
     * @brief 按类别获取模板ID
     * @param category 模板类别
     * @return 模板ID列表
     */
    std::vector<std::string> getTemplateIdsByCategory(StyleTemplateCategory category) const;
    
    /**
     * @brief 清空所有非内置模板
     */
    void clearUserTemplates();
    
    /**
     * @brief 重置为默认模板
     */
    void resetToDefaults();
    
    // ==================== 内置模板加载 ====================
    
    /**
     * @brief 加载内置模板
     */
    void loadBuiltInTemplates();
    
    /**
     * @brief 创建现代简约模板
     * @return 模板
     */
    static TXStyleTemplate createModernTemplate();
    
    /**
     * @brief 创建经典商务模板
     * @return 模板
     */
    static TXStyleTemplate createBusinessTemplate();
    
    /**
     * @brief 创建彩色主题模板
     * @return 模板
     */
    static TXStyleTemplate createColorfulTemplate();
    
    /**
     * @brief 创建简单数据表模板
     * @return 模板
     */
    static TXStyleTemplate createDataTableTemplate();
    
    /**
     * @brief 创建报告模板
     * @return 模板
     */
    static TXStyleTemplate createReportTemplate();
    
    // ==================== 文件操作 ====================
    
    /**
     * @brief 从文件加载模板
     * @param filePath 文件路径
     * @return 成功返回true
     */
    bool loadTemplateFromFile(const std::string& filePath);
    
    /**
     * @brief 保存模板到文件
     * @param templateId 模板ID
     * @param filePath 文件路径
     * @return 成功返回true
     */
    bool saveTemplateToFile(const std::string& templateId, const std::string& filePath) const;
    
    /**
     * @brief 导出所有用户模板
     * @param directoryPath 目录路径
     * @return 成功返回true
     */
    bool exportUserTemplates(const std::string& directoryPath) const;
    
    /**
     * @brief 导入模板目录
     * @param directoryPath 目录路径
     * @return 成功导入的模板数量
     */
    int importTemplatesFromDirectory(const std::string& directoryPath);

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
    
    // 禁用拷贝
    TXStyleTemplateManager(const TXStyleTemplateManager&) = delete;
    TXStyleTemplateManager& operator=(const TXStyleTemplateManager&) = delete;
};

} // namespace TinaXlsx