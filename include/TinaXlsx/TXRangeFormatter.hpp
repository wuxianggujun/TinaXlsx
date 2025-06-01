#pragma once

#include "TXTypes.hpp"
#include "TXRange.hpp"
#include <string>

namespace TinaXlsx {

/**
 * @brief 数据范围格式化工具类
 * 
 * 专门用于将TXRange对象格式化为Excel图表所需的引用字符串
 */
class TXRangeFormatter {
public:
    /**
     * @brief 格式化类别轴范围（X轴标签）
     * @param range 数据范围
     * @param sheetName 工作表名称
     * @return 格式化的范围字符串，如 "Sheet1!$A$7:$A$9"
     */
    static std::string formatCategoryRange(const TXRange& range, const std::string& sheetName);

    /**
     * @brief 格式化数值范围（Y轴数据）
     * @param range 数据范围
     * @param sheetName 工作表名称
     * @return 格式化的范围字符串，如 "Sheet1!$B$7:$B$9"
     */
    static std::string formatValueRange(const TXRange& range, const std::string& sheetName);

    /**
     * @brief 格式化散点图X值范围
     * @param range 数据范围
     * @param sheetName 工作表名称
     * @return 格式化的范围字符串
     */
    static std::string formatScatterXRange(const TXRange& range, const std::string& sheetName);

    /**
     * @brief 格式化散点图Y值范围
     * @param range 数据范围
     * @param sheetName 工作表名称
     * @return 格式化的范围字符串
     */
    static std::string formatScatterYRange(const TXRange& range, const std::string& sheetName);

    /**
     * @brief 格式化饼图标签范围
     * @param range 数据范围
     * @param sheetName 工作表名称
     * @return 格式化的范围字符串
     */
    static std::string formatPieLabelRange(const TXRange& range, const std::string& sheetName);

    /**
     * @brief 格式化饼图数值范围
     * @param range 数据范围
     * @param sheetName 工作表名称
     * @return 格式化的范围字符串
     */
    static std::string formatPieValueRange(const TXRange& range, const std::string& sheetName);

private:
    /**
     * @brief 构建Excel范围引用字符串
     * @param sheetName 工作表名称
     * @param startCol 起始列
     * @param startRow 起始行
     * @param endCol 结束列
     * @param endRow 结束行
     * @return 格式化的范围字符串
     */
    static std::string buildRangeString(const std::string& sheetName,
                                      const column_t& startCol, const row_t& startRow,
                                      const column_t& endCol, const row_t& endRow);

    /**
     * @brief 转义工作表名称（如果包含特殊字符）
     * @param sheetName 原始工作表名称
     * @return 转义后的工作表名称
     */
    static std::string escapeSheetName(const std::string& sheetName);
};

} // namespace TinaXlsx
