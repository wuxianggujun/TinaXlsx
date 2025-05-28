#pragma once

#include "TXTypes.hpp"
#include "TXCoordinate.hpp"
#include "TXXmlReader.hpp"
#include <string>
#include <memory>

namespace TinaXlsx {

// Forward declarations
class TXSheet;

/**
 * @brief 工作表读取器 - 专门从Excel工作表XML解析数据到TXSheet对象
 * 
 * 负责解析Excel工作表XML格式并将数据加载到TXSheet对象中。
 * 这个类封装了ZIP读取操作，TXWorkbook不需要直接访问ZIP。
 */
class TXWorksheetReader {
public:
    TXWorksheetReader();
    ~TXWorksheetReader();

    /**
     * @brief 从XLSX文件中读取指定工作表
     * @param xlsxFilePath XLSX文件路径
     * @param sheet 目标工作表对象
     * @param sheetIndex 工作表索引（从1开始）
     * @return 成功返回true
     */
    bool readWorksheetFromFile(const std::string& xlsxFilePath, TXSheet* sheet, std::size_t sheetIndex);

    /**
     * @brief 从XML字符串解析工作表
     * @param sheet 目标工作表对象
     * @param xmlContent XML内容
     * @return 成功返回true
     */
    bool parseFromXml(TXSheet* sheet, const std::string& xmlContent);

    /**
     * @brief 获取最后的错误信息
     * @return 错误信息
     */
    const std::string& getLastError() const;

private:
    /**
     * @brief 解析工作表数据
     * @param sheet 目标工作表
     * @param xmlReader XML读取器
     * @return 成功返回true
     */
    bool parseWorksheetData(TXSheet* sheet, const TXXmlReader& xmlReader);

    /**
     * @brief 解析单行数据
     * @param sheet 目标工作表
     * @param rowNode 行XML节点信息
     * @return 成功返回true
     */
    bool parseRowData(TXSheet* sheet, const XmlNodeInfo& rowNode);

    /**
     * @brief 解析单元格数据
     * @param sheet 目标工作表
     * @param cellNode 单元格XML节点信息
     * @return 成功返回true
     */
    bool parseCellData(TXSheet* sheet, const XmlNodeInfo& cellNode);

    /**
     * @brief 解析合并单元格
     * @param sheet 目标工作表
     * @param mergeCellsNode 合并单元格XML节点信息
     * @return 成功返回true
     */
    bool parseMergeCells(TXSheet* sheet, const XmlNodeInfo& mergeCellsNode);

    /**
     * @brief 解析单元格值
     * @param cellNode 单元格XML节点信息
     * @param cellType 单元格类型
     * @return 解析的值
     */
    cell_value_t parseCellValue(const XmlNodeInfo& cellNode, const std::string& cellType);

    /**
     * @brief 从地址字符串解析坐标
     * @param address 地址字符串（如"A1"）
     * @return 坐标对象
     */
    TXCoordinate parseAddress(const std::string& address);

    /**
     * @brief 获取工作表的XML路径
     * @param sheetIndex 工作表索引
     * @return XML文件路径
     */
    std::string getWorksheetXmlPath(std::size_t sheetIndex) const;

private:
    std::string lastError_;
    std::unique_ptr<TXXmlReader> xmlReader_;
};

} // namespace TinaXlsx 