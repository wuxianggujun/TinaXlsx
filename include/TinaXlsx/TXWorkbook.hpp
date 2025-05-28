#pragma once

#include "TXTypes.hpp"
#include "TXComponentManager.hpp"
#include <string>
#include <vector>
#include <memory>
#include <unordered_map>
#include <unordered_set>

namespace TinaXlsx {

// Forward declarations
class TXSheet;
class TXWorksheetReader;
class TXWorksheetWriter;

/**
 * @brief Excel工作簿类
 * 支持创建、读取和写入Excel文件(.xlsx格式)
 * 采用组件化设计，按需生成所需功能
 */
class TXWorkbook {
public:
    TXWorkbook();
    ~TXWorkbook();

    // 禁用拷贝构造和赋值
    TXWorkbook(const TXWorkbook&) = delete;
    TXWorkbook& operator=(const TXWorkbook&) = delete;

    // 支持移动构造和赋值
    TXWorkbook(TXWorkbook&& other) noexcept;
    TXWorkbook& operator=(TXWorkbook&& other) noexcept;

    /**
     * @brief 从文件加载工作簿
     * @param filename XLSX文件路径
     * @return 成功返回true，失败返回false
     */
    bool loadFromFile(const std::string& filename);

    /**
     * @brief 保存工作簿到文件
     * @param filename 输出文件路径
     * @return 成功返回true，失败返回false
     */
    bool saveToFile(const std::string& filename);

    /**
     * @brief 创建新的工作表
     * @param name 工作表名称
     * @return 成功返回工作表指针，失败返回nullptr
     */
    TXSheet* addSheet(const std::string& name);

    /**
     * @brief 获取工作表
     * @param name 工作表名称
     * @return 工作表指针，如果不存在返回nullptr
     */
    TXSheet* getSheet(const std::string& name);

    /**
     * @brief 获取工作表
     * @param index 工作表索引
     * @return 工作表指针，如果索引无效返回nullptr
     */
    TXSheet* getSheet(std::size_t index);

    /**
     * @brief 删除工作表
     * @param name 工作表名称
     * @return 成功返回true，失败返回false
     */
    bool removeSheet(const std::string& name);

    /**
     * @brief 获取工作表数量
     * @return 工作表数量
     */
    std::size_t getSheetCount() const;

    /**
     * @brief 获取所有工作表名称
     * @return 工作表名称列表
     */
    std::vector<std::string> getSheetNames() const;

    /**
     * @brief 检查是否有指定名称的工作表
     * @param name 工作表名称
     * @return 存在返回true，否则返回false
     */
    bool hasSheet(const std::string& name) const;

    /**
     * @brief 重命名工作表
     * @param oldName 原名称
     * @param newName 新名称
     * @return 成功返回true，失败返回false
     */
    bool renameSheet(const std::string& oldName, const std::string& newName);

    /**
     * @brief 获取活动工作表
     * @return 活动工作表指针，如果没有则返回nullptr
     */
    TXSheet* getActiveSheet();

    /**
     * @brief 设置活动工作表
     * @param name 工作表名称
     * @return 成功返回true，失败返回false
     */
    bool setActiveSheet(const std::string& name);

    /**
     * @brief 获取最后的错误信息
     * @return 错误信息字符串
     */
    const std::string& getLastError() const;

    /**
     * @brief 清空工作簿
     */
    void clear();

    /**
     * @brief 检查工作簿是否为空
     * @return 为空返回true
     */
    bool isEmpty() const;

    /**
     * @brief 获取组件管理器
     * @return 组件管理器引用
     */
    ComponentManager& getComponentManager();
    
    /**
     * @brief 获取组件管理器（常量版本）
     * @return 组件管理器常量引用
     */
    const ComponentManager& getComponentManager() const;
    
    /**
     * @brief 启用智能组件检测（默认启用）
     * 开启后会自动检测使用的功能并注册相应组件
     * @param enable 是否启用
     */
    void setAutoComponentDetection(bool enable);
    
    /**
     * @brief 手动注册组件
     * @param component 要注册的组件
     */
    void registerComponent(ExcelComponent component);

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

} // namespace TinaXlsx 