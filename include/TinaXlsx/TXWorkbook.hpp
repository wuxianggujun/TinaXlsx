#pragma once

#include <string>
#include <vector>
#include <memory>
#include <atomic>
#include <unordered_map>
#include "TXTypes.hpp"
#include "TXComponentManager.hpp"
#include "TXStyleManager.hpp"
#include "TXSharedStringsPool.hpp"
#include "TXWorkbookProtectionManager.hpp"
#include "TXUnifiedMemoryManager.hpp"
#include "TXBatchWorksheetWriter.hpp"

namespace TinaXlsx
{
    struct TXWorkbookContext;
    class TXCellStyle;
    class TXSheet;
    class TXXmlHandler;

    /**
     * @brief Excel工作簿类
     * 支持创建、读取和写入Excel文件(.xlsx格式)
     * 采用组件化设计，按需生成所需功能
     */
    class TXWorkbook
    {
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
         * @brief 🚀 高性能批量保存工作簿到文件
         * @param filename 输出文件路径
         * @param config 批量保存配置（可选）
         * @return 成功返回true，失败返回false
         */
        bool saveToFileBatch(const std::string& filename,
                           const TXBatchWorksheetWriter::BatchConfig& config = TXBatchWorksheetWriter::BatchConfig{});

        /**
         * @brief 获取最后一次批量保存的统计信息
         * @return 批量保存统计信息
         */
        const TXBatchWorksheetWriter::BatchStats& getLastBatchStats() const;

        /**
         * @brief 创建新的工作表
         * @param name 工作表名称
         * @return 成功返回工作表指针，失败返回nullptr
         */
        TXSheet* addSheet(const std::string& name);

        /**
         * @brief 添加工作表
         * @param sheet 工作表指针
         * @return 成功返回工作表指针，失败返回nullptr
         */
        TXSheet* addSheet(std::unique_ptr<TXSheet> sheet);

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
        TXSheet* getSheet(u64 index) const;

        /**
         * @brief 删除工作表
         * @param name 工作表名称
         * @return 成功返回true，失败返回false
         */
        bool removeSheet(const std::string& name);

        /**
         * @brief 注册或获取样式ID
         * @param style 单元格样式
         * @return 样式ID
         */
        u32 registerOrGetStyleFId(const TXCellStyle& style);

        /**
         * @brief 获取工作簿上下文
         * @return 工作簿上下文指针
         */
        TXWorkbookContext* getContext();
        
        /**
         * @brief 获取工作表数量
         * @return 工作表数量
         */
        u64 getSheetCount() const;

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
         * @brief 设置活动工作表
         * @param index 工作表索引
         * @return 成功返回true，失败返回false
         */
        bool setActiveSheet(u64 index);

        /**
         * @brief 获取活动工作表索引
         * @return 活动工作表索引
         */
        u64 getActiveSheetIndex() const;

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

        /**
         * @brief 获取工作表列表
         * @return 工作表列表引用
         */
        std::vector<std::unique_ptr<TXSheet>>& getSheets();

        /**
         * @brief 获取工作表列表（常量版本）
         * @return 工作表列表常量引用
         */
        const std::vector<std::unique_ptr<TXSheet>>& getSheets() const;

        /**
         * @brief 获取样式管理器
         * @return 样式管理器引用
         */
        TXStyleManager& getStyleManager();

        /**
         * @brief 获取样式管理器（常量版本）
         * @return 样式管理器常量引用
         */
        const TXStyleManager& getStyleManager() const;

        /**
         * @brief 为保存做准备
         */
        void prepareForSaving();

        /**
         * @brief 存储工作表
         * @param sheet_uptr 工作表智能指针
         * @return 工作表指针
         */
        TXSheet* storeSheet(std::unique_ptr<TXSheet> sheet_uptr);

        // ==================== 工作簿保护功能 ====================

        /**
         * @brief 获取工作簿保护管理器
         * @return 工作簿保护管理器引用
         */
        TXWorkbookProtectionManager& getWorkbookProtectionManager();

        /**
         * @brief 获取工作簿保护管理器（常量版本）
         * @return 工作簿保护管理器常量引用
         */
        const TXWorkbookProtectionManager& getWorkbookProtectionManager() const;

        /**
         * @brief 保护工作簿（便捷方法）
         * @param password 保护密码
         * @param protection 保护配置（可选）
         * @return 成功返回true
         */
        bool protectWorkbook(const std::string& password,
                           const TXWorkbookProtectionManager::WorkbookProtection& protection = TXWorkbookProtectionManager::WorkbookProtection{});

        /**
         * @brief 解除工作簿保护（便捷方法）
         * @param password 保护密码
         * @return 成功返回true
         */
        bool unprotectWorkbook(const std::string& password);

        /**
         * @brief 检查工作簿是否受保护（便捷方法）
         * @return 受保护返回true
         */
        bool isWorkbookProtected() const;

        /**
         * @brief 保护工作簿结构（便捷方法）
         * @param password 保护密码
         * @return 成功返回true
         */
        bool protectStructure(const std::string& password);

        /**
         * @brief 保护工作簿窗口（便捷方法）
         * @param password 保护密码
         * @return 成功返回true
         */
        bool protectWindows(const std::string& password);

        // ==================== 透视表功能 ====================

        /**
         * @brief 为工作表添加透视表
         * @param sheetName 工作表名称
         * @param pivotTable 透视表对象
         * @return 成功返回true
         */
        bool addPivotTable(const std::string& sheetName, std::shared_ptr<class TXPivotTable> pivotTable);

        /**
         * @brief 获取工作表的透视表列表
         * @param sheetName 工作表名称
         * @return 透视表列表
         */
        std::vector<std::shared_ptr<class TXPivotTable>> getPivotTables(const std::string& sheetName) const;

        /**
         * @brief 移除工作表的所有透视表
         * @param sheetName 工作表名称
         * @return 成功返回true
         */
        bool removePivotTables(const std::string& sheetName);

    private:
        std::vector<std::unique_ptr<TXSheet>> sheets_;
        std::size_t active_sheet_index_;
        std::string last_error_;
        ComponentManager component_manager_;
        bool auto_component_detection_;
        TXStyleManager style_manager_;
        TXSharedStringsPool shared_strings_pool_;
        TXWorkbookProtectionManager workbook_protection_manager_;  ///< 工作簿保护管理器
        std::unique_ptr<TXWorkbookContext> context_;  // 必须在workbook_protection_manager_之后声明

        // 透视表管理
        std::unordered_map<std::string, std::vector<std::shared_ptr<class TXPivotTable>>> pivot_tables_;  ///< 工作表名称到透视表列表的映射

        // 🚀 批量保存相关
        mutable TXBatchWorksheetWriter::BatchStats lastBatchStats_;  ///< 最后一次批量保存统计

        // 透视表辅助方法
        std::string generatePivotCacheRecordsXml(const TXPivotTable* pivotTable, const std::string& sheetName) const;
    };
} // namespace TinaXlsx 
