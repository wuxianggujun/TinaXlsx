#pragma once

#include "TXTypes.hpp"
#include "TXRange.hpp"
#include "TXXmlWriter.hpp"
#include <string>
#include <memory>

namespace TinaXlsx {

// Forward declarations
class TXSheet;
class TXCell;

/**
 * @brief 工作表写入器 - 专门生成工作表XML并写入到XLSX文件
 * 
 * 负责将TXSheet对象转换为标准的Excel工作表XML格式。
 * 这个类封装了ZIP写入操作，TXWorkbook不需要直接访问ZIP。
 */
class TXWorksheetWriter {
public:
    TXWorksheetWriter();
    ~TXWorksheetWriter();

    /**
     * @brief 将工作表写入到XLSX文件中
     * @param xlsxFilePath XLSX文件路径
     * @param sheet 工作表对象
     * @param sheetIndex 工作表索引（从1开始）
     * @param appendMode 是否为追加模式（打开现有文件追加工作表）
     * @return 成功返回true
     */
    bool writeWorksheetToFile(const std::string& xlsxFilePath, const TXSheet* sheet, 
                             std::size_t sheetIndex, bool appendMode = false);

    /**
     * @brief 生成工作表XML字符串
     * @param sheet 工作表对象
     * @return XML字符串
     */
    std::string generateXml(const TXSheet* sheet);

    /**
     * @brief 获取最后的错误信息
     * @return 错误信息
     */
    const std::string& getLastError() const;

private:
    /**
     * @brief 构建工作表XML结构
     * @param sheet 工作表对象
     * @return 构建的根节点
     */
    XmlNodeBuilder buildWorksheetXml(const TXSheet* sheet);

    /**
     * @brief 构建维度信息节点
     * @param sheet 工作表对象
     * @return 维度节点
     */
    XmlNodeBuilder buildDimensionNode(const TXSheet* sheet);

    /**
     * @brief 构建工作表数据节点
     * @param sheet 工作表对象
     * @return 数据节点
     */
    XmlNodeBuilder buildSheetDataNode(const TXSheet* sheet);

    /**
     * @brief 构建单行数据节点
     * @param sheet 工作表对象
     * @param row 行号
     * @param usedRange 使用范围
     * @return 行节点
     */
    XmlNodeBuilder buildRowNode(const TXSheet* sheet, row_t row, const TXRange& usedRange);

    /**
     * @brief 构建单个单元格节点
     * @param cell 单元格对象
     * @param cellRef 单元格引用（如A1）
     * @return 单元格节点
     */
    XmlNodeBuilder buildCellNode(const TXCell* cell, const std::string& cellRef);

    /**
     * @brief 构建合并单元格信息节点
     * @param sheet 工作表对象
     * @return 合并单元格节点
     */
    XmlNodeBuilder buildMergeCellsNode(const TXSheet* sheet);

    /**
     * @brief 获取工作表的XML路径
     * @param sheetIndex 工作表索引
     * @return XML文件路径
     */
    std::string getWorksheetXmlPath(std::size_t sheetIndex) const;

private:
    mutable std::string lastError_;
    std::unique_ptr<TXXmlWriter> xmlWriter_;
};

} // namespace TinaXlsx 