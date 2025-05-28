#pragma once

#include "TXTypes.hpp"
#include "TXCoordinate.hpp"
#include "TXXmlReader.hpp"
#include <string>
#include <memory>

namespace TinaXlsx {

// Forward declarations
class TXSheet;
class TXZipArchiveReader;

/**
 * @brief 工作表读取器 - 专门从Excel工作表XML解析数据到TXSheet对象
 * 
 * 负责解析Excel工作表XML格式并将数据加载到TXSheet对象中
 */
class TXWorksheetReader {
public:
    TXWorksheetReader();
    ~TXWorksheetReader();

    /**
     * @brief 从ZIP文件中读取工作表
     * @param zip ZIP处理器
     * @param sheet 目标工作表对象
     * @param sheetIndex 工作表索引（从1开始）
     * @return 成功返回true
     */
    bool readFromZip(TXZipArchiveReader& zip, TXSheet* sheet, std::size_t sheetIndex);

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
     * @param rootNode 根XML节点
     * @return 成功返回true
     */
    bool parseWorksheetData(TXSheet* sheet, const TXXmlNode* rootNode);

    /**
     * @brief 解析单行数据
     * @param sheet 目标工作表
     * @param rowNode 行XML节点
     * @return 成功返回true
     */
    bool parseRowData(TXSheet* sheet, const TXXmlNode* rowNode);

    /**
     * @brief 解析单元格数据
     * @param sheet 目标工作表
     * @param cellNode 单元格XML节点
     * @return 成功返回true
     */
    bool parseCellData(TXSheet* sheet, const TXXmlNode* cellNode);

    /**
     * @brief 解析合并单元格
     * @param sheet 目标工作表
     * @param mergeCellsNode 合并单元格XML节点
     * @return 成功返回true
     */
    bool parseMergeCells(TXSheet* sheet, const TXXmlNode* mergeCellsNode);

    /**
     * @brief 解析单元格值
     * @param cellNode 单元格XML节点
     * @param cellType 单元格类型
     * @return 解析的值
     */
    cell_value_t parseCellValue(const TXXmlNode* cellNode, const std::string& cellType);

    /**
     * @brief 从地址字符串解析坐标
     * @param address 地址字符串（如"A1"）
     * @return 坐标对象
     */
    TXCoordinate parseAddress(const std::string& address);

private:
    std::string lastError_;
};

} // namespace TinaXlsx 