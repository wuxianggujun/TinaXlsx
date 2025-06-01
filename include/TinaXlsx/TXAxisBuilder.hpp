#pragma once

#include "TXTypes.hpp"
#include "TXXmlWriter.hpp"

namespace TinaXlsx {

/**
 * @brief 坐标轴构建器
 * 
 * 负责生成图表的坐标轴XML
 */
class TXAxisBuilder {
public:
    /**
     * @brief 构建类别轴（X轴）XML
     * @param axisId 轴ID
     * @param crossAxisId 交叉轴ID
     * @return 类别轴XML节点
     */
    static XmlNodeBuilder buildCategoryAxis(u32 axisId = 1, u32 crossAxisId = 2);

    /**
     * @brief 构建数值轴（Y轴）XML
     * @param axisId 轴ID
     * @param crossAxisId 交叉轴ID
     * @param showGridlines 是否显示网格线
     * @return 数值轴XML节点
     */
    static XmlNodeBuilder buildValueAxis(u32 axisId = 2, u32 crossAxisId = 1, bool showGridlines = true);

    /**
     * @brief 构建日期轴XML
     * @param axisId 轴ID
     * @param crossAxisId 交叉轴ID
     * @return 日期轴XML节点
     */
    static XmlNodeBuilder buildDateAxis(u32 axisId = 1, u32 crossAxisId = 2);

private:
    /**
     * @brief 添加通用轴属性
     * @param axis 轴节点
     * @param axisId 轴ID
     * @param position 轴位置（"b"=底部, "l"=左侧, "t"=顶部, "r"=右侧）
     * @param crossAxisId 交叉轴ID
     */
    static void addCommonAxisProperties(XmlNodeBuilder& axis, u32 axisId, 
                                      const std::string& position, u32 crossAxisId);

    /**
     * @brief 添加轴缩放属性
     * @param axis 轴节点
     * @param orientation 方向（"minMax" 或 "maxMin"）
     */
    static void addScaling(XmlNodeBuilder& axis, const std::string& orientation = "minMax");

    /**
     * @brief 添加刻度标签位置
     * @param axis 轴节点
     * @param position 位置（"nextTo", "low", "high", "none"）
     */
    static void addTickLabelPosition(XmlNodeBuilder& axis, const std::string& position = "nextTo");

    /**
     * @brief 添加交叉设置
     * @param axis 轴节点
     * @param crosses 交叉类型（"autoZero", "max", "min"）
     */
    static void addCrosses(XmlNodeBuilder& axis, const std::string& crosses = "autoZero");
};

} // namespace TinaXlsx
