//
// @file TXStreamXmlReader.hpp
// @brief 流式XML读取器 - 高性能Excel文件读取
//

#pragma once

#include <string>
#include <functional>
#include <memory>
#include "TXResult.hpp"
#include "TXTypes.hpp"
#include "TXCoordinate.hpp"

// 前向声明
namespace pugi {
class xml_document;
class xml_node;
}

namespace TinaXlsx {

// 前向声明
class TXZipArchiveReader;
class TXSheet;

/**
 * @brief 单元格数据结构（轻量级）
 */
struct CellData {
    std::string ref;        // 单元格引用 (如 "A1")
    std::string value;      // 单元格值
    std::string type;       // 单元格类型 ("s", "n", "b", "inlineStr")
    u32 styleIndex = 0;     // 样式索引
    
    // 解析后的坐标（缓存）
    mutable row_t row = row_t(row_t::index_t(0));
    mutable column_t col = column_t(column_t::index_t(0));
    mutable bool coordinatesParsed = false;
    
    void parseCoordinates() const;
};

/**
 * @brief 行数据结构
 */
struct RowData {
    u32 rowIndex;
    std::vector<CellData> cells;
    
    // 行属性
    double height = 0.0;
    bool customHeight = false;
    bool hidden = false;
};

/**
 * @brief 流式XML读取器回调接口
 */
class IStreamXmlCallback {
public:
    virtual ~IStreamXmlCallback() = default;
    
    /**
     * @brief 处理单个行数据
     * @param rowData 行数据
     * @return true继续处理，false停止
     */
    virtual bool onRowData(const RowData& rowData) = 0;
    
    /**
     * @brief 处理共享字符串
     * @param index 字符串索引
     * @param text 字符串内容
     */
    virtual void onSharedString(u32 index, const std::string& text) = 0;
    
    /**
     * @brief 处理样式信息
     * @param styleIndex 样式索引
     * @param styleData 样式数据（简化）
     */
    virtual void onStyleData(u32 styleIndex, const std::string& styleData) = 0;
};

/**
 * @brief 🚀 高性能流式XML读取器
 * 
 * 专门针对Excel文件优化的流式解析器，避免构建完整DOM树
 */
class TXStreamXmlReader {
public:
    TXStreamXmlReader();
    ~TXStreamXmlReader();

    /**
     * @brief 流式解析工作表XML
     * @param zipReader ZIP读取器
     * @param worksheetPath 工作表路径
     * @param callback 回调接口
     * @return 解析结果
     */
    TXResult<void> parseWorksheet(TXZipArchiveReader& zipReader, 
                                 const std::string& worksheetPath,
                                 IStreamXmlCallback& callback);

    /**
     * @brief 流式解析共享字符串XML
     * @param zipReader ZIP读取器
     * @param callback 回调接口
     * @return 解析结果
     */
    TXResult<void> parseSharedStrings(TXZipArchiveReader& zipReader,
                                     IStreamXmlCallback& callback);

    /**
     * @brief 流式解析样式XML
     * @param zipReader ZIP读取器
     * @param callback 回调接口
     * @return 解析结果
     */
    TXResult<void> parseStyles(TXZipArchiveReader& zipReader,
                              IStreamXmlCallback& callback);

    /**
     * @brief 设置解析选项
     */
    struct ParseOptions {
        bool skipEmptyCells = true;      // 跳过空单元格
        bool parseFormulas = true;       // 解析公式
        bool parseStyles = true;         // 解析样式
        size_t batchSize = 1000;         // 批处理大小
    };
    
    void setOptions(const ParseOptions& options) { options_ = options; }

private:
    ParseOptions options_;
    std::unique_ptr<pugi::xml_document> doc_;
    
    // 解析工作表的具体实现
    TXResult<void> parseWorksheetImpl(const std::string& xmlContent, IStreamXmlCallback& callback);
    
    // 解析单个行节点
    bool parseRowNode(const pugi::xml_node& rowNode, IStreamXmlCallback& callback);
    
    // 解析单个单元格节点
    CellData parseCellNode(const pugi::xml_node& cellNode);
    
    // 批量处理行数据
    void processBatch(std::vector<RowData>& batch, IStreamXmlCallback& callback);
};

/**
 * @brief 🚀 高性能工作表加载器
 * 
 * 结合流式解析和批量操作的工作表加载器
 */
class TXFastWorksheetLoader : public IStreamXmlCallback {
public:
    explicit TXFastWorksheetLoader(TXSheet* sheet);
    
    /**
     * @brief 加载工作表数据
     * @param zipReader ZIP读取器
     * @param worksheetPath 工作表路径
     * @return 加载结果
     */
    TXResult<void> load(TXZipArchiveReader& zipReader, const std::string& worksheetPath);
    
    // IStreamXmlCallback 实现
    bool onRowData(const RowData& rowData) override;
    void onSharedString(u32 index, const std::string& text) override;
    void onStyleData(u32 styleIndex, const std::string& styleData) override;
    
    /**
     * @brief 获取加载统计信息
     */
    struct LoadStats {
        size_t totalRows = 0;
        size_t totalCells = 0;
        size_t emptySkipped = 0;
        double loadTimeMs = 0.0;
    };
    
    const LoadStats& getStats() const { return stats_; }

private:
    TXSheet* sheet_;
    TXStreamXmlReader reader_;
    LoadStats stats_;

    // 🚀 高性能批量设置方法
    void batchSetCells(const std::vector<std::pair<TXCoordinate, std::string>>& cells);

    // 批量单元格数据（已弃用）
    std::vector<std::pair<std::string, std::string>> cellBatch_;
    static constexpr size_t BATCH_SIZE = 1000;

    void flushBatch();
};

} // namespace TinaXlsx
