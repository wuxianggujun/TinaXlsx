#pragma once

#include "TXTypes.hpp"
#include "TXRange.hpp"
#include "TXXmlWriter.hpp"
#include <string>
#include <memory>

namespace TinaXlsx {

// Forward declarations
class TXSheet;
class TXZipArchiveWriter;
class TXCell;

/**
 * @brief 工作表写入器 - 专门生成工作表XML或直接流式写入ZIP
 * 
 * 负责将TXSheet对象转换为标准的Excel工作表XML格式
 */
class TXWorksheetWriter {
public:
    TXWorksheetWriter();
    ~TXWorksheetWriter();

    /**
     * @brief 将工作表写入到ZIP文件中
     * @param zip ZIP处理器
     * @param sheet 工作表对象
     * @param sheetIndex 工作表索引（从1开始）
     * @return 成功返回true
     */
    bool writeToZip(TXZipArchiveWriter& zip, const TXSheet* sheet, std::size_t sheetIndex);

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
     * @brief 写入工作表头部信息
     * @param writer XML写入器
     * @param sheet 工作表对象
     */
    void writeWorksheetHeader(TXXmlWriter& writer, const TXSheet* sheet);

    /**
     * @brief 写入维度信息
     * @param writer XML写入器
     * @param sheet 工作表对象
     */
    void writeDimension(TXXmlWriter& writer, const TXSheet* sheet);

    /**
     * @brief 写入工作表数据
     * @param writer XML写入器
     * @param sheet 工作表对象
     */
    void writeSheetData(TXXmlWriter& writer, const TXSheet* sheet);

    /**
     * @brief 写入单行数据
     * @param writer XML写入器
     * @param sheet 工作表对象
     * @param row 行号
     * @param usedRange 使用范围
     */
    bool writeRow(TXXmlWriter& writer, const TXSheet* sheet, row_t row, const TXRange& usedRange);

    /**
     * @brief 写入单个单元格
     * @param writer XML写入器
     * @param cell 单元格对象
     * @param cellRef 单元格引用（如A1）
     */
    void writeCell(TXXmlWriter& writer, const TXCell* cell, const std::string& cellRef);

    /**
     * @brief 写入合并单元格信息
     * @param writer XML写入器
     * @param sheet 工作表对象
     */
    void writeMergeCells(TXXmlWriter& writer, const TXSheet* sheet);

private:
    mutable std::string lastError_;
};

} // namespace TinaXlsx 