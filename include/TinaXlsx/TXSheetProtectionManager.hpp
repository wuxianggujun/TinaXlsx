#pragma once

#include <string>
#include <unordered_set>
#include "TXTypes.hpp"
#include "TXCoordinate.hpp"
#include "TXRange.hpp"

namespace TinaXlsx {

// 前向声明
class TXCellManager;

/**
 * @brief 工作表保护管理器
 * 
 * 专门负责工作表的保护功能
 * 职责：
 * - 工作表密码保护
 * - 权限控制
 * - 单元格锁定状态管理
 * - 操作权限验证
 */
class TXSheetProtectionManager {
public:
    /**
     * @brief 工作表保护选项
     */
    struct SheetProtection {
        bool isProtected = false;                    ///< 是否受保护
        std::string passwordHash;                    ///< 保护密码（哈希值）
        std::string algorithmName = "SHA-512";      ///< 哈希算法名称
        std::string saltValue;                       ///< 盐值（Base64编码）
        uint32_t spinCount = 100000;                ///< 迭代次数
        bool selectLockedCells = true;               ///< 允许选择锁定的单元格
        bool selectUnlockedCells = true;             ///< 允许选择未锁定的单元格
        bool formatCells = false;                    ///< 允许格式化单元格
        bool formatColumns = false;                  ///< 允许格式化列
        bool formatRows = false;                     ///< 允许格式化行
        bool insertColumns = false;                  ///< 允许插入列
        bool insertRows = false;                     ///< 允许插入行
        bool deleteColumns = false;                  ///< 允许删除列
        bool deleteRows = false;                     ///< 允许删除行
        bool insertHyperlinks = false;               ///< 允许插入超链接
        bool sort = false;                           ///< 允许排序
        bool autoFilter = false;                     ///< 允许自动筛选
        bool pivotTables = false;                    ///< 允许数据透视表操作
        bool objects = false;                        ///< 允许编辑对象
        bool scenarios = false;                      ///< 允许编辑方案

        /**
         * @brief 默认构造函数
         */
        SheetProtection() = default;

        /**
         * @brief 创建严格保护配置
         * @return 严格保护的配置
         */
        static SheetProtection createStrictProtection() {
            SheetProtection protection;
            protection.isProtected = true;
            // 所有操作都被禁止（除了选择单元格）
            return protection;
        }

        /**
         * @brief 创建宽松保护配置
         * @return 宽松保护的配置
         */
        static SheetProtection createLooseProtection() {
            SheetProtection protection;
            protection.isProtected = true;
            protection.formatCells = true;
            protection.formatColumns = true;
            protection.formatRows = true;
            protection.sort = true;
            protection.autoFilter = true;
            return protection;
        }
    };

    /**
     * @brief 操作类型枚举
     */
    enum class OperationType {
        SelectLockedCells,
        SelectUnlockedCells,
        FormatCells,
        FormatColumns,
        FormatRows,
        InsertColumns,
        InsertRows,
        DeleteColumns,
        DeleteRows,
        InsertHyperlinks,
        Sort,
        AutoFilter,
        PivotTables,
        Objects,
        Scenarios
    };

    TXSheetProtectionManager() = default;
    ~TXSheetProtectionManager() = default;

    // 禁用拷贝，支持移动
    TXSheetProtectionManager(const TXSheetProtectionManager&) = delete;
    TXSheetProtectionManager& operator=(const TXSheetProtectionManager&) = delete;
    TXSheetProtectionManager(TXSheetProtectionManager&&) = default;
    TXSheetProtectionManager& operator=(TXSheetProtectionManager&&) = default;

    // ==================== 工作表保护 ====================

    /**
     * @brief 保护工作表
     * @param password 保护密码（空字符串表示无密码）
     * @param protection 保护选项
     * @return 成功返回true
     */
    bool protectSheet(const std::string& password = "", const SheetProtection& protection = SheetProtection{});

    /**
     * @brief 取消工作表保护
     * @param password 解除保护的密码
     * @return 成功返回true
     */
    bool unprotectSheet(const std::string& password = "");

    /**
     * @brief 检查工作表是否受保护
     * @return 受保护返回true
     */
    bool isSheetProtected() const { return protection_.isProtected; }

    /**
     * @brief 获取工作表保护设置
     * @return 保护设置
     */
    const SheetProtection& getSheetProtection() const { return protection_; }

    /**
     * @brief 验证密码
     * @param password 输入的密码
     * @return 验证成功返回true
     */
    bool verifyPassword(const std::string& password) const;

    // ==================== 单元格锁定 ====================

    /**
     * @brief 设置单元格锁定状态
     * @param coord 坐标
     * @param locked 是否锁定
     * @param cellManager 单元格管理器
     * @return 成功返回true
     */
    bool setCellLocked(const TXCoordinate& coord, bool locked, TXCellManager& cellManager);

    /**
     * @brief 检查单元格是否锁定
     * @param coord 坐标
     * @param cellManager 单元格管理器
     * @return 锁定返回true
     */
    bool isCellLocked(const TXCoordinate& coord, const TXCellManager& cellManager) const;

    /**
     * @brief 设置范围内单元格的锁定状态
     * @param range 范围
     * @param locked 是否锁定
     * @param cellManager 单元格管理器
     * @return 成功设置的单元格数量
     */
    std::size_t setRangeLocked(const TXRange& range, bool locked, TXCellManager& cellManager);

    /**
     * @brief 批量设置单元格锁定状态
     * @param coords 坐标列表
     * @param locked 是否锁定
     * @param cellManager 单元格管理器
     * @return 成功设置的数量
     */
    std::size_t setCellsLocked(const std::vector<TXCoordinate>& coords, bool locked, TXCellManager& cellManager);

    // ==================== 权限验证 ====================

    /**
     * @brief 检查操作是否被允许
     * @param operation 操作类型
     * @return 允许返回true
     */
    bool isOperationAllowed(OperationType operation) const;

    /**
     * @brief 检查操作是否被允许（字符串版本）
     * @param operationName 操作名称
     * @return 允许返回true
     */
    bool isOperationAllowed(const std::string& operationName) const;

    /**
     * @brief 检查单元格是否可以编辑
     * @param coord 坐标
     * @param cellManager 单元格管理器
     * @return 可编辑返回true
     */
    bool isCellEditable(const TXCoordinate& coord, const TXCellManager& cellManager) const;

    /**
     * @brief 检查范围内的单元格是否都可以编辑
     * @param range 范围
     * @param cellManager 单元格管理器
     * @return 都可编辑返回true
     */
    bool isRangeEditable(const TXRange& range, const TXCellManager& cellManager) const;

    // ==================== 保护状态查询 ====================

    /**
     * @brief 获取所有被锁定的单元格坐标
     * @param cellManager 单元格管理器
     * @return 锁定单元格坐标列表
     */
    std::vector<TXCoordinate> getLockedCells(const TXCellManager& cellManager) const;

    /**
     * @brief 获取所有未锁定的单元格坐标
     * @param cellManager 单元格管理器
     * @return 未锁定单元格坐标列表
     */
    std::vector<TXCoordinate> getUnlockedCells(const TXCellManager& cellManager) const;

    /**
     * @brief 获取保护统计信息
     */
    struct ProtectionStats {
        bool isProtected = false;
        bool hasPassword = false;
        std::size_t lockedCellCount = 0;
        std::size_t unlockedCellCount = 0;
        std::size_t allowedOperationCount = 0;
    };

    /**
     * @brief 获取保护统计信息
     * @param cellManager 单元格管理器
     * @return 统计信息
     */
    ProtectionStats getProtectionStats(const TXCellManager& cellManager) const;

    // ==================== 工具方法 ====================

    /**
     * @brief 清空保护设置
     */
    void clear();

    /**
     * @brief 重置为默认保护设置
     */
    void reset();

private:
    SheetProtection protection_;

    /**
     * @brief 生成密码哈希
     * @param password 原始密码
     * @return 哈希值
     */
    std::string generatePasswordHash(const std::string& password) const;

    /**
     * @brief 操作类型转换为字符串
     * @param operation 操作类型
     * @return 操作名称
     */
    std::string operationTypeToString(OperationType operation) const;

    /**
     * @brief 字符串转换为操作类型
     * @param operationName 操作名称
     * @return 操作类型
     */
    OperationType stringToOperationType(const std::string& operationName) const;
};

} // namespace TinaXlsx
