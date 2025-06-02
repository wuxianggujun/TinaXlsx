//
// @file TXUnifiedXmlHandler.hpp
// @brief 统一的简单XML处理器 - 合并多个简单Handler
//

#pragma once

#include "TXXmlHandler.hpp"
#include "TXXmlWriter.hpp"
#include "TXTypes.hpp"
#include <memory>

namespace TinaXlsx {

// 前向声明
class TXPivotTable;

/**
 * @brief 统一的简单XML处理器
 * 
 * 合并多个功能简单的XML处理器，减少代码重复和维护成本
 */
class TXUnifiedXmlHandler : public TXXmlHandler {
public:
    /**
     * @brief 处理器类型枚举
     */
    enum class HandlerType {
        MainRels,           // _rels/.rels
        WorkbookRels,       // xl/_rels/workbook.xml.rels
        WorksheetRels,      // xl/worksheets/_rels/sheetN.xml.rels
        DocumentProperties, // docProps/core.xml 和 docProps/app.xml (同时生成两个文件)
        PivotTableRels,     // xl/pivotTables/_rels/pivotTableN.xml.rels
        PivotCacheRels      // xl/pivotCache/_rels/pivotCacheDefinitionN.xml.rels
    };

public:
    /**
     * @brief 构造函数
     * @param type 处理器类型
     * @param index 索引（用于需要编号的处理器，如工作表、透视表等）
     */
    explicit TXUnifiedXmlHandler(HandlerType type, u32 index = 0);

    /**
     * @brief 析构函数
     */
    ~TXUnifiedXmlHandler() override = default;

    /**
     * @brief 加载XML（大部分简单处理器不需要加载）
     */
    TXResult<void> load(TXZipArchiveReader& zipReader, TXWorkbookContext& context) override;

    /**
     * @brief 保存XML
     */
    TXResult<void> save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) override;

    /**
     * @brief 获取XML文件路径
     */
    std::string partName() const override;

    /**
     * @brief 设置透视表信息（用于透视表相关的处理器）
     * @param pivotTables 透视表列表
     */
    void setPivotTables(const std::vector<std::shared_ptr<TXPivotTable>>& pivotTables);

    /**
     * @brief 设置所有透视表信息（用于工作簿级别的处理器）
     * @param allPivotTables 所有工作表的透视表映射
     */
    void setAllPivotTables(const std::unordered_map<std::string, std::vector<std::shared_ptr<TXPivotTable>>>& allPivotTables);

private:
    HandlerType type_;
    u32 index_;
    
    // 透视表相关数据
    std::vector<std::shared_ptr<TXPivotTable>> pivot_tables_;
    std::unordered_map<std::string, std::vector<std::shared_ptr<TXPivotTable>>> all_pivot_tables_;



    // ==================== 流式生成方法 ====================

    /**
     * @brief 流式生成主关系文件
     */
    TXResult<void> generateMainRelsStream(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) const;

    /**
     * @brief 流式生成工作簿关系文件
     */
    TXResult<void> generateWorkbookRelsStream(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) const;

    /**
     * @brief 流式生成工作表关系文件
     */
    TXResult<void> generateWorksheetRelsStream(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) const;

    /**
     * @brief 流式生成文档属性文件
     */
    TXResult<void> generateDocumentPropertiesStream(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) const;

    /**
     * @brief 流式生成透视表关系文件
     */
    TXResult<void> generatePivotTableRelsStream(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) const;

    /**
     * @brief 流式生成透视表缓存关系文件
     */
    TXResult<void> generatePivotCacheRelsStream(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context) const;

    // ==================== 辅助方法 ====================

    /**
     * @brief 获取文件路径
     */
    std::string getPartName() const;



    /**
     * @brief 检查是否需要处理透视表
     */
    bool needsPivotTableProcessing() const;
};

/**
 * @brief 统一XML处理器工厂
 * 
 * 提供便捷的创建方法
 */
class TXUnifiedXmlHandlerFactory {
public:
    /**
     * @brief 创建主关系处理器
     */
    static std::unique_ptr<TXUnifiedXmlHandler> createMainRelsHandler();
    
    /**
     * @brief 创建工作簿关系处理器
     */
    static std::unique_ptr<TXUnifiedXmlHandler> createWorkbookRelsHandler();
    
    /**
     * @brief 创建工作表关系处理器
     * @param sheetIndex 工作表索引
     */
    static std::unique_ptr<TXUnifiedXmlHandler> createWorksheetRelsHandler(u32 sheetIndex);
    
    /**
     * @brief 创建文档属性处理器 (同时处理core.xml和app.xml)
     */
    static std::unique_ptr<TXUnifiedXmlHandler> createDocumentPropertiesHandler();
    
    /**
     * @brief 创建透视表关系处理器
     * @param pivotTableId 透视表ID
     */
    static std::unique_ptr<TXUnifiedXmlHandler> createPivotTableRelsHandler(u32 pivotTableId);
    
    /**
     * @brief 创建透视表缓存关系处理器
     * @param cacheId 缓存ID
     */
    static std::unique_ptr<TXUnifiedXmlHandler> createPivotCacheRelsHandler(u32 cacheId);
};

} // namespace TinaXlsx
