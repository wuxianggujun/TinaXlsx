#pragma once

#include <string>
#include <cstdint>

namespace TinaXlsx {

/**
 * @brief 工作簿保护管理器
 * 
 * 负责管理Excel工作簿的保护功能，包括：
 * - 工作簿结构保护（防止添加/删除/重命名工作表）
 * - 窗口保护（防止调整窗口大小和位置）
 * - 修订保护（防止跟踪更改）
 * 
 * 使用现代Excel的SHA-512密码哈希算法，完全兼容：
 * - Microsoft Excel
 * - WPS Office
 * - LibreOffice Calc
 */
class TXWorkbookProtectionManager {
public:
    /**
     * @brief 工作簿保护配置结构
     */
    struct WorkbookProtection {
        bool isProtected = false;                    ///< 是否受保护
        std::string passwordHash;                    ///< 保护密码（哈希值）
        std::string algorithmName = "SHA-512";      ///< 哈希算法名称
        std::string saltValue;                       ///< 盐值（Base64编码）
        uint32_t spinCount = 100000;                ///< 迭代次数
        bool lockStructure = true;                   ///< 锁定工作簿结构
        bool lockWindows = false;                    ///< 锁定窗口
        bool lockRevision = false;                   ///< 锁定修订
    };

    /**
     * @brief 构造函数
     */
    TXWorkbookProtectionManager() = default;

    /**
     * @brief 析构函数
     */
    ~TXWorkbookProtectionManager() = default;

    /**
     * @brief 保护工作簿
     * @param password 保护密码
     * @param protection 保护配置
     * @return 成功返回true
     */
    bool protectWorkbook(const std::string& password, const WorkbookProtection& protection = WorkbookProtection{});

    /**
     * @brief 解除工作簿保护
     * @param password 保护密码
     * @return 成功返回true
     */
    bool unprotectWorkbook(const std::string& password);

    /**
     * @brief 验证密码
     * @param password 待验证的密码
     * @return 密码正确返回true
     */
    bool verifyPassword(const std::string& password) const;

    /**
     * @brief 检查工作簿是否受保护
     * @return 受保护返回true
     */
    bool isWorkbookProtected() const;

    /**
     * @brief 获取工作簿保护配置
     * @return 保护配置的常量引用
     */
    const WorkbookProtection& getWorkbookProtection() const;

    /**
     * @brief 设置工作簿保护配置（不包含密码）
     * @param protection 保护配置
     */
    void setWorkbookProtection(const WorkbookProtection& protection);

    /**
     * @brief 生成密码哈希
     * @param password 原始密码
     * @return 密码哈希字符串
     */
    std::string generatePasswordHash(const std::string& password) const;

    // 便捷方法

    /**
     * @brief 启用结构保护（防止添加/删除/重命名工作表）
     * @param password 保护密码
     * @return 成功返回true
     */
    bool protectStructure(const std::string& password);

    /**
     * @brief 启用窗口保护（防止调整窗口大小和位置）
     * @param password 保护密码
     * @return 成功返回true
     */
    bool protectWindows(const std::string& password);

    /**
     * @brief 启用修订保护（防止跟踪更改）
     * @param password 保护密码
     * @return 成功返回true
     */
    bool protectRevision(const std::string& password);

    /**
     * @brief 检查结构是否受保护
     * @return 受保护返回true
     */
    bool isStructureProtected() const;

    /**
     * @brief 检查窗口是否受保护
     * @return 受保护返回true
     */
    bool isWindowsProtected() const;

    /**
     * @brief 检查修订是否受保护
     * @return 受保护返回true
     */
    bool isRevisionProtected() const;

private:
    WorkbookProtection protection_;  ///< 工作簿保护配置
};

} // namespace TinaXlsx
