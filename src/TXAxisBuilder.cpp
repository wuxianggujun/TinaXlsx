#include "TinaXlsx/TXAxisBuilder.hpp"

namespace TinaXlsx {

XmlNodeBuilder TXAxisBuilder::buildCategoryAxis(u32 axisId, u32 crossAxisId) {
    XmlNodeBuilder catAx("c:catAx");

    // 添加通用属性
    addCommonAxisProperties(catAx, axisId, "b", crossAxisId);

    // 添加缩放
    addScaling(catAx);

    // 添加刻度标签位置
    addTickLabelPosition(catAx);

    // 添加交叉设置
    addCrosses(catAx);

    // 类别轴特有属性
    XmlNodeBuilder auto1("c:auto");
    auto1.addAttribute("val", "1");
    catAx.addChild(auto1);

    XmlNodeBuilder lblAlgn("c:lblAlgn");
    lblAlgn.addAttribute("val", "ctr");
    catAx.addChild(lblAlgn);

    XmlNodeBuilder lblOffset("c:lblOffset");
    lblOffset.addAttribute("val", "100");
    catAx.addChild(lblOffset);

    return catAx;
}

XmlNodeBuilder TXAxisBuilder::buildValueAxis(u32 axisId, u32 crossAxisId, bool showGridlines) {
    XmlNodeBuilder valAx("c:valAx");

    // 添加通用属性
    addCommonAxisProperties(valAx, axisId, "l", crossAxisId);

    // 添加缩放
    addScaling(valAx);

    // 添加主网格线（如果需要）
    if (showGridlines) {
        XmlNodeBuilder majorGridlines("c:majorGridlines");
        valAx.addChild(majorGridlines);
    }

    // 添加刻度标签位置
    addTickLabelPosition(valAx);

    // 添加交叉设置
    addCrosses(valAx);

    // 数值轴特有属性
    XmlNodeBuilder crossBetween("c:crossBetween");
    crossBetween.addAttribute("val", "between");
    valAx.addChild(crossBetween);

    return valAx;
}

XmlNodeBuilder TXAxisBuilder::buildDateAxis(u32 axisId, u32 crossAxisId) {
    XmlNodeBuilder dateAx("c:dateAx");

    // 添加通用属性
    addCommonAxisProperties(dateAx, axisId, "b", crossAxisId);

    // 添加缩放
    addScaling(dateAx);

    // 添加刻度标签位置
    addTickLabelPosition(dateAx);

    // 添加交叉设置
    addCrosses(dateAx);

    // 日期轴特有属性
    XmlNodeBuilder auto1("c:auto");
    auto1.addAttribute("val", "1");
    dateAx.addChild(auto1);

    XmlNodeBuilder lblOffset("c:lblOffset");
    lblOffset.addAttribute("val", "100");
    dateAx.addChild(lblOffset);

    // 基本时间单位
    XmlNodeBuilder baseTimeUnit("c:baseTimeUnit");
    baseTimeUnit.addAttribute("val", "days");
    dateAx.addChild(baseTimeUnit);

    return dateAx;
}

void TXAxisBuilder::addCommonAxisProperties(XmlNodeBuilder& axis, u32 axisId, 
                                          const std::string& position, u32 crossAxisId) {
    // 轴ID
    XmlNodeBuilder axIdNode("c:axId");
    axIdNode.addAttribute("val", std::to_string(axisId));
    axis.addChild(axIdNode);

    // 轴位置
    XmlNodeBuilder axPos("c:axPos");
    axPos.addAttribute("val", position);
    axis.addChild(axPos);

    // 交叉轴ID
    XmlNodeBuilder crossAx("c:crossAx");
    crossAx.addAttribute("val", std::to_string(crossAxisId));
    axis.addChild(crossAx);
}

void TXAxisBuilder::addScaling(XmlNodeBuilder& axis, const std::string& orientation) {
    XmlNodeBuilder scaling("c:scaling");
    XmlNodeBuilder orientationNode("c:orientation");
    orientationNode.addAttribute("val", orientation);
    scaling.addChild(orientationNode);
    axis.addChild(scaling);
}

void TXAxisBuilder::addTickLabelPosition(XmlNodeBuilder& axis, const std::string& position) {
    XmlNodeBuilder tickLblPos("c:tickLblPos");
    tickLblPos.addAttribute("val", position);
    axis.addChild(tickLblPos);
}

void TXAxisBuilder::addCrosses(XmlNodeBuilder& axis, const std::string& crosses) {
    XmlNodeBuilder crossesNode("c:crosses");
    crossesNode.addAttribute("val", crosses);
    axis.addChild(crossesNode);
}

} // namespace TinaXlsx
