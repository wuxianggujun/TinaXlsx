#pragma once

#include <string>
#include <vector>
#include <memory>
#include <unordered_map>

namespace TinaXlsx {

// Forward declarations
class TXSheet;
class TXZipHandler;

/**
 * @brief Excel工作簿类
 * 
 * 管理整个Excel工作簿的生命周期，包括打开、保存、Sheet管理
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
     * @return 为空返回true，否则返回false
     */
    bool isEmpty() const;

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

} // namespace TinaXlsx 