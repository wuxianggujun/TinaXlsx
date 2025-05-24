/**
 * @file WorksheetDataParser.hpp
 * @brief 工作表数据解析器 - 专门负责工作表数据的解析
 */

#pragma once

#include "Types.hpp"
#include "ExcelStructureManager.hpp"
#include "DataCache.hpp"
#include "XmlParser.hpp"
#include <functional>
#include <optional>

namespace TinaXlsx {

/**
 * @brief 工作表数据解析器
 * 专门负责工作表数据的解析，支持流式解析和回调
 */
class WorksheetDataParser {
public:
    using CellCallback = std::function<bool(const CellPosition& position, const CellValue& value)>;
    using RowCallback = std::function<bool(RowIndex rowIndex, const RowData& rowData)>;

private:
    ExcelStructureManager* structureManager_;
    DataCache* cache_;

public:
    /**
     * @brief 构造函数
     * @param structureManager Excel结构管理器指针
     * @param cache 数据缓存指针
     */
    WorksheetDataParser(ExcelStructureManager* structureManager, DataCache* cache);
    
    /**
     * @brief 析构函数
     */
    ~WorksheetDataParser() = default;
    
    // 禁用拷贝，启用移动
    WorksheetDataParser(const WorksheetDataParser&) = delete;
    WorksheetDataParser& operator=(const WorksheetDataParser&) = delete;
    WorksheetDataParser(WorksheetDataParser&&) = default;
    WorksheetDataParser& operator=(WorksheetDataParser&&) = default;
    
    /**
     * @brief 解析工作表数据
     * @param sheetInfo 工作表信息
     * @return TableData 解析得到的表格数据
     */
    TableData parseWorksheetData(const ExcelStructureManager::SheetInfo& sheetInfo);
    
    /**
     * @brief 解析工作表数据（使用回调）
     * @param sheetInfo 工作表信息
     * @param cellCallback 单元格回调函数
     * @param rowCallback 行回调函数
     * @return size_t 处理的行数
     */
    size_t parseWorksheetDataWithCallback(
        const ExcelStructureManager::SheetInfo& sheetInfo,
        CellCallback cellCallback = nullptr,
        RowCallback rowCallback = nullptr
    );
    
    /**
     * @brief 解析单个单元格
     * @param sheetInfo 工作表信息
     * @param position 单元格位置
     * @return std::optional<CellValue> 单元格值
     */
    std::optional<CellValue> parseCell(
        const ExcelStructureManager::SheetInfo& sheetInfo, 
        const CellPosition& position
    );
    
    /**
     * @brief 解析单行数据
     * @param sheetInfo 工作表信息
     * @param rowIndex 行索引
     * @param maxColumns 最大列数（0表示无限制）
     * @return std::optional<RowData> 行数据
     */
    std::optional<RowData> parseRow(
        const ExcelStructureManager::SheetInfo& sheetInfo,
        RowIndex rowIndex,
        ColumnIndex maxColumns = 0
    );
    
    /**
     * @brief 解析指定范围的数据
     * @param sheetInfo 工作表信息
     * @param range 数据范围
     * @return TableData 范围内的数据
     */
    TableData parseRange(
        const ExcelStructureManager::SheetInfo& sheetInfo,
        const CellRange& range
    );
    
    /**
     * @brief 获取工作表维度信息
     * @param sheetInfo 工作表信息
     * @return std::pair<RowIndex, ColumnIndex> 行数和列数
     */
    std::pair<RowIndex, ColumnIndex> getSheetDimensions(
        const ExcelStructureManager::SheetInfo& sheetInfo
    );
    
    /**
     * @brief 检查是否有缓存的数据
     * @param sheetInfo 工作表信息
     * @return bool 是否有缓存
     */
    bool hasCachedData(const ExcelStructureManager::SheetInfo& sheetInfo) const;
    
    /**
     * @brief 清除指定工作表的缓存
     * @param sheetInfo 工作表信息
     */
    void clearCacheForSheet(const ExcelStructureManager::SheetInfo& sheetInfo);

private:
    /**
     * @brief 从缓存或解析获取完整数据
     * @param sheetInfo 工作表信息
     * @return TableData 完整的工作表数据
     */
    TableData getOrParseFullData(const ExcelStructureManager::SheetInfo& sheetInfo);
    
    /**
     * @brief 实际解析工作表XML内容
     * @param sheetInfo 工作表信息
     * @param xmlContent XML内容
     * @return TableData 解析得到的数据
     */
    TableData parseWorksheetXml(
        const ExcelStructureManager::SheetInfo& sheetInfo,
        const std::string& xmlContent
    );
    
    /**
     * @brief 计算工作表的实际维度
     * @param data 工作表数据
     * @return std::pair<RowIndex, ColumnIndex> 实际的行数和列数
     */
    std::pair<RowIndex, ColumnIndex> calculateDimensions(const TableData& data);
};

} // namespace TinaXlsx 