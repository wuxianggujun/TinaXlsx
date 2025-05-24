/**
 * @file WorksheetDataParser.cpp
 * @brief 工作表数据解析器实现
 */

#include "TinaXlsx/WorksheetDataParser.hpp"
#include "TinaXlsx/Exception.hpp"

namespace TinaXlsx {

WorksheetDataParser::WorksheetDataParser(ExcelStructureManager* structureManager, DataCache* cache)
    : structureManager_(structureManager), cache_(cache) {
    
    if (!structureManager_) {
        throw InvalidArgumentException("Structure manager cannot be null");
    }
    
    if (!cache_) {
        throw InvalidArgumentException("Cache cannot be null");
    }
}

TableData WorksheetDataParser::parseWorksheetData(const ExcelStructureManager::SheetInfo& sheetInfo) {
    return getOrParseFullData(sheetInfo);
}

size_t WorksheetDataParser::parseWorksheetDataWithCallback(
    const ExcelStructureManager::SheetInfo& sheetInfo,
    CellCallback cellCallback,
    RowCallback rowCallback) {
    
    auto zipReader = structureManager_->getZipReader();
    if (!zipReader) {
        throw ParseException("ZIP reader not available");
    }
    
    auto xmlContent = zipReader->readWorksheet(sheetInfo.filePath);
    if (!xmlContent) {
        throw ParseException("Failed to read worksheet: " + sheetInfo.filePath);
    }
    
    WorksheetXmlParser parser;
    
    // 设置共享字符串
    auto& sharedStrings = const_cast<std::vector<std::string>&>(structureManager_->getSharedStrings());
    parser.setSharedStrings(&sharedStrings);
    
    size_t processedRows = 0;
    
    // 设置回调
    if (cellCallback) {
        parser.setCellCallback([&](const CellPosition& position, const CellValue& value) {
            cellCallback(position, value);
        });
    }
    
    if (rowCallback) {
        parser.setRowCallback([&](RowIndex rowIndex, const RowData& rowData) {
            processedRows++;
            rowCallback(rowIndex, rowData);
        });
    }
    
    // 解析工作表
    parser.parseWorksheet(*xmlContent);
    
    return processedRows;
}

std::optional<CellValue> WorksheetDataParser::parseCell(
    const ExcelStructureManager::SheetInfo& sheetInfo, 
    const CellPosition& position) {
    
    // 首先尝试从缓存获取
    auto cachedData = cache_->getCachedTableData(sheetInfo.getCacheKey());
    if (cachedData) {
        if (position.row < cachedData->size() && 
            position.column < (*cachedData)[position.row].size()) {
            return (*cachedData)[position.row][position.column];
        }
        return std::nullopt;
    }
    
    // 如果没有缓存，解析整个工作表（效率可能不高，但保证正确性）
    auto fullData = getOrParseFullData(sheetInfo);
    if (position.row < fullData.size() && 
        position.column < fullData[position.row].size()) {
        return fullData[position.row][position.column];
    }
    
    return std::nullopt;
}

std::optional<RowData> WorksheetDataParser::parseRow(
    const ExcelStructureManager::SheetInfo& sheetInfo,
    RowIndex rowIndex,
    ColumnIndex maxColumns) {
    
    // 首先尝试从缓存获取
    auto cachedData = cache_->getCachedTableData(sheetInfo.getCacheKey());
    if (cachedData) {
        if (rowIndex < cachedData->size()) {
            RowData row = (*cachedData)[rowIndex];
            if (maxColumns > 0 && row.size() > maxColumns) {
                row.resize(maxColumns);
            }
            return row;
        }
        return std::nullopt;
    }
    
    // 如果没有缓存，解析整个工作表
    auto fullData = getOrParseFullData(sheetInfo);
    if (rowIndex < fullData.size()) {
        RowData row = fullData[rowIndex];
        if (maxColumns > 0 && row.size() > maxColumns) {
            row.resize(maxColumns);
        }
        return row;
    }
    
    return std::nullopt;
}

TableData WorksheetDataParser::parseRange(
    const ExcelStructureManager::SheetInfo& sheetInfo,
    const CellRange& range) {
    
    auto fullData = getOrParseFullData(sheetInfo);
    
    TableData result;
    for (RowIndex row = range.start.row; row <= range.end.row && row < fullData.size(); ++row) {
        RowData rowData;
        for (ColumnIndex col = range.start.column; 
             col <= range.end.column && col < fullData[row].size(); ++col) {
            rowData.push_back(fullData[row][col]);
        }
        result.push_back(rowData);
    }
    
    return result;
}

std::pair<RowIndex, ColumnIndex> WorksheetDataParser::getSheetDimensions(
    const ExcelStructureManager::SheetInfo& sheetInfo) {
    
    // 首先检查缓存
    auto cachedDimensions = cache_->getCachedDimensions(sheetInfo.getCacheKey());
    if (cachedDimensions) {
        return *cachedDimensions;
    }
    
    // 获取数据并计算维度
    auto data = getOrParseFullData(sheetInfo);
    auto dimensions = calculateDimensions(data);
    
    // 缓存维度信息
    cache_->cacheDimensions(sheetInfo.getCacheKey(), dimensions.first, dimensions.second);
    
    return dimensions;
}

bool WorksheetDataParser::hasCachedData(const ExcelStructureManager::SheetInfo& sheetInfo) const {
    return cache_->hasTableData(sheetInfo.getCacheKey());
}

void WorksheetDataParser::clearCacheForSheet(const ExcelStructureManager::SheetInfo& sheetInfo) {
    cache_->clearCache(sheetInfo.getCacheKey());
}

TableData WorksheetDataParser::getOrParseFullData(const ExcelStructureManager::SheetInfo& sheetInfo) {
    // 检查缓存
    auto cacheKey = sheetInfo.getCacheKey();
    auto cachedData = cache_->getCachedTableData(cacheKey);
    if (cachedData) {
        return *cachedData;
    }
    
    // 从ZIP读取内容
    auto zipReader = structureManager_->getZipReader();
    if (!zipReader) {
        throw ParseException("ZIP reader not available");
    }
    
    auto xmlContent = zipReader->readWorksheet(sheetInfo.filePath);
    if (!xmlContent) {
        throw ParseException("Failed to read worksheet: " + sheetInfo.filePath);
    }
    
    // 解析XML内容
    auto data = parseWorksheetXml(sheetInfo, *xmlContent);
    
    // 缓存数据
    cache_->cacheTableData(cacheKey, data);
    
    return data;
}

TableData WorksheetDataParser::parseWorksheetXml(
    const ExcelStructureManager::SheetInfo& sheetInfo,
    const std::string& xmlContent) {
    
    WorksheetXmlParser parser;
    
    // 设置共享字符串
    auto& sharedStrings = const_cast<std::vector<std::string>&>(structureManager_->getSharedStrings());
    parser.setSharedStrings(&sharedStrings);
    
    TableData result;
    RowIndex maxRowIndex = 0;
    
    // 设置行回调来收集数据
    parser.setRowCallback([&](RowIndex rowIndex, const RowData& rowData) {
        // 确保result有足够的行
        while (result.size() <= rowIndex) {
            result.emplace_back();
        }
        
        result[rowIndex] = rowData;
        maxRowIndex = std::max(maxRowIndex, static_cast<RowIndex>(rowIndex));
    });
    
    // 解析工作表
    parser.parseWorksheet(xmlContent);
    
    // 清理空行
    while (!result.empty() && result.back().empty()) {
        result.pop_back();
    }
    
    return result;
}

std::pair<RowIndex, ColumnIndex> WorksheetDataParser::calculateDimensions(const TableData& data) {
    if (data.empty()) {
        return {0, 0};
    }
    
    RowIndex rows = data.size();
    ColumnIndex maxCols = 0;
    
    for (const auto& row : data) {
        maxCols = std::max(maxCols, static_cast<ColumnIndex>(row.size()));
    }
    
    return {rows, maxCols};
}

} // namespace TinaXlsx 