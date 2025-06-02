//
// Created by wuxianggujun on 2025/1/15.
//

#include "TinaXlsx/TXChartXmlHandler.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXError.hpp"
#include "TinaXlsx/TXXmlWriter.hpp"
#include "TinaXlsx/TXChartSeriesBuilder.hpp"
#include "TinaXlsx/TXAxisBuilder.hpp"
#include "TinaXlsx/TXPugiStreamWriter.hpp"  // 流式写入器
#include <sstream>
#include <cstring>  // for strlen

namespace TinaXlsx
{
    // ==================== TXChartXmlHandler 实现 ====================

    TXChartXmlHandler::TXChartXmlHandler(const TXChart* chart, u32 chartIndex)
        : m_chart(chart), m_chartIndex(chartIndex)
    {
    }

    TXResult<void> TXChartXmlHandler::load(TXZipArchiveReader& /*zipReader*/, TXWorkbookContext& /*context*/)
    {
        // 暂不实现图表加载功能
        return Err<void>(TXErrorCode::InvalidArgument, "Chart loading not implemented");
    }

    TXResult<void> TXChartXmlHandler::save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context)
    {
        if (!m_chart) {
            return Err<void>(TXErrorCode::InvalidArgument, "Chart is null");
        }

        XmlNodeBuilder chartSpace = generateChartXml();

        TXXmlWriter writer;
        auto setRootResult = writer.setRootNode(chartSpace);
        if (setRootResult.isError()) {
            return Err<void>(setRootResult.error().getCode(), 
                           "Failed to set root node: " + setRootResult.error().getMessage());
        }

        auto xmlContentResult = writer.generateXmlString();
        if (xmlContentResult.isError()) {
            return Err<void>(xmlContentResult.error().getCode(), 
                           "Failed to generate XML: " + xmlContentResult.error().getMessage());
        }

        std::vector<uint8_t> xmlData(xmlContentResult.value().begin(), xmlContentResult.value().end());
        auto writeResult = zipWriter.write(partName(), xmlData);
        if (writeResult.isError()) {
            return Err<void>(writeResult.error().getCode(), 
                           "Failed to write " + partName() + ": " + writeResult.error().getMessage());
        }

        return Ok();
    }

    std::string TXChartXmlHandler::partName() const
    {
        return "xl/charts/chart" + std::to_string(m_chartIndex + 1) + ".xml";
    }

    XmlNodeBuilder TXChartXmlHandler::generateChartXml() const
    {
        XmlNodeBuilder chartSpace("c:chartSpace");
        chartSpace.addAttribute("xmlns:c", "http://schemas.openxmlformats.org/drawingml/2006/chart")
                  .addAttribute("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main")
                  .addAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        // 添加日期1904设置
        chartSpace.addChild(XmlNodeBuilder("c:date1904").addAttribute("val", "0"));

        // 添加语言设置
        chartSpace.addChild(XmlNodeBuilder("c:lang").addAttribute("val", "zh-CN"));

        // 添加图表
        XmlNodeBuilder chart("c:chart");
        
        // 添加标题
        if (!m_chart->getName().empty()) {
            chart.addChild(generateTitleXml());
        }

        // 添加绘图区域
        chart.addChild(generatePlotAreaXml());

        // 添加图例
        chart.addChild(generateLegendXml());

        chartSpace.addChild(chart);

        return chartSpace;
    }

    XmlNodeBuilder TXChartXmlHandler::generateTitleXml() const
    {
        XmlNodeBuilder title("c:title");
        
        XmlNodeBuilder tx("c:tx");
        XmlNodeBuilder rich("c:rich");
        
        // 添加文本属性
        XmlNodeBuilder bodyPr("a:bodyPr");
        rich.addChild(bodyPr);
        
        XmlNodeBuilder lstStyle("a:lstStyle");
        rich.addChild(lstStyle);
        
        // 添加段落
        XmlNodeBuilder p("a:p");
        XmlNodeBuilder pPr("a:pPr");
        XmlNodeBuilder defRPr("a:defRPr");
        pPr.addChild(defRPr);
        p.addChild(pPr);
        
        // 添加文本内容
        XmlNodeBuilder r("a:r");
        XmlNodeBuilder t("a:t");
        t.setText(m_chart->getName());
        r.addChild(t);
        p.addChild(r);
        
        rich.addChild(p);
        tx.addChild(rich);
        title.addChild(tx);
        
        return title;
    }

    XmlNodeBuilder TXChartXmlHandler::generateLegendXml() const
    {
        XmlNodeBuilder legend("c:legend");
        
        // 图例位置
        XmlNodeBuilder legendPos("c:legendPos");
        legendPos.addAttribute("val", "r"); // 右侧
        legend.addChild(legendPos);
        
        // 图例布局
        XmlNodeBuilder layout("c:layout");
        legend.addChild(layout);
        
        return legend;
    }

    XmlNodeBuilder TXChartXmlHandler::generatePlotAreaXml() const
    {
        XmlNodeBuilder plotArea("c:plotArea");
        
        // 添加布局
        XmlNodeBuilder layout("c:layout");
        plotArea.addChild(layout);
        
        // 根据图表类型生成相应的图表XML
        switch (m_chart->getType()) {
            case ChartType::Column:
                plotArea.addChild(generateColumnChartXml());
                break;
            case ChartType::Line:
                plotArea.addChild(generateLineChartXml());
                break;
            case ChartType::Pie:
                plotArea.addChild(generatePieChartXml());
                break;
            case ChartType::Scatter:
                plotArea.addChild(generateScatterChartXml());
                break;
            default:
                plotArea.addChild(generateColumnChartXml()); // 默认为柱状图
                break;
        }
        
        // 添加坐标轴（饼图除外）
        if (m_chart->getType() != ChartType::Pie) {
            plotArea.addChild(generateCategoryAxisXml());
            plotArea.addChild(generateValueAxisXml());
        }
        
        return plotArea;
    }

    XmlNodeBuilder TXChartXmlHandler::generateColumnChartXml() const
    {
        XmlNodeBuilder barChart("c:barChart");
        
        // 柱状图方向
        XmlNodeBuilder barDir("c:barDir");
        barDir.addAttribute("val", "col");
        barChart.addChild(barDir);
        
        // 分组类型
        XmlNodeBuilder grouping("c:grouping");
        grouping.addAttribute("val", "clustered");
        barChart.addChild(grouping);
        
        // 添加数据系列
        barChart.addChild(generateSeriesXml());
        
        // 坐标轴ID
        XmlNodeBuilder axId1("c:axId");
        axId1.addAttribute("val", "1");
        barChart.addChild(axId1);
        
        XmlNodeBuilder axId2("c:axId");
        axId2.addAttribute("val", "2");
        barChart.addChild(axId2);
        
        return barChart;
    }

    XmlNodeBuilder TXChartXmlHandler::generateLineChartXml() const
    {
        XmlNodeBuilder lineChart("c:lineChart");
        
        // 分组类型
        XmlNodeBuilder grouping("c:grouping");
        grouping.addAttribute("val", "standard");
        lineChart.addChild(grouping);
        
        // 添加数据系列
        lineChart.addChild(generateSeriesXml());
        
        // 坐标轴ID
        XmlNodeBuilder axId1("c:axId");
        axId1.addAttribute("val", "1");
        lineChart.addChild(axId1);
        
        XmlNodeBuilder axId2("c:axId");
        axId2.addAttribute("val", "2");
        lineChart.addChild(axId2);
        
        return lineChart;
    }

    XmlNodeBuilder TXChartXmlHandler::generatePieChartXml() const
    {
        XmlNodeBuilder pieChart("c:pieChart");
        
        // 添加数据系列
        pieChart.addChild(generateSeriesXml());
        
        return pieChart;
    }

    XmlNodeBuilder TXChartXmlHandler::generateScatterChartXml() const
    {
        XmlNodeBuilder scatterChart("c:scatterChart");
        
        // 散点样式
        XmlNodeBuilder scatterStyle("c:scatterStyle");
        scatterStyle.addAttribute("val", "lineMarker");
        scatterChart.addChild(scatterStyle);
        
        // 添加数据系列
        scatterChart.addChild(generateSeriesXml());
        
        // 坐标轴ID
        XmlNodeBuilder axId1("c:axId");
        axId1.addAttribute("val", "1");
        scatterChart.addChild(axId1);
        
        XmlNodeBuilder axId2("c:axId");
        axId2.addAttribute("val", "2");
        scatterChart.addChild(axId2);
        
        return scatterChart;
    }

    XmlNodeBuilder TXChartXmlHandler::generateSeriesXml() const
    {
        // 使用工厂创建相应的系列构建器
        auto seriesBuilder = TXSeriesBuilderFactory::createBuilder(m_chart->getType());
        return seriesBuilder->buildSeries(m_chart, 0);
    }

    XmlNodeBuilder TXChartXmlHandler::generateCategoryAxisXml() const
    {
        return TXAxisBuilder::buildCategoryAxis(1, 2);
    }

    XmlNodeBuilder TXChartXmlHandler::generateValueAxisXml() const
    {
        return TXAxisBuilder::buildValueAxis(2, 1, true);
    }

    std::string TXChartXmlHandler::getChartTypeString() const
    {
        switch (m_chart->getType()) {
            case ChartType::Column: return "column";
            case ChartType::Bar: return "bar";
            case ChartType::Line: return "line";
            case ChartType::Pie: return "pie";
            case ChartType::Scatter: return "scatter";
            case ChartType::Area: return "area";
            case ChartType::Radar: return "radar";
            case ChartType::Combo: return "combo";
            default: return "column";
        }
    }

    // ==================== TXChartRelsXmlHandler 实现 ====================

    TXChartRelsXmlHandler::TXChartRelsXmlHandler(u32 chartIndex)
        : m_chartIndex(chartIndex)
    {
    }

    TXResult<void> TXChartRelsXmlHandler::load(TXZipArchiveReader& /*zipReader*/, TXWorkbookContext& /*context*/)
    {
        // 暂不实现
        return Err<void>(TXErrorCode::InvalidArgument, "Chart rels loading not implemented");
    }

    TXResult<void> TXChartRelsXmlHandler::save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context)
    {
        XmlNodeBuilder relationships("Relationships");
        relationships.addAttribute("xmlns", "http://schemas.openxmlformats.org/package/2006/relationships");

        // 添加样式关系（如果需要）
        // 这里可以添加图表特定的关系

        TXXmlWriter writer;
        auto setRootResult = writer.setRootNode(relationships);
        if (setRootResult.isError()) {
            return Err<void>(setRootResult.error().getCode(),
                           "Failed to set root node: " + setRootResult.error().getMessage());
        }

        auto xmlContentResult = writer.generateXmlString();
        if (xmlContentResult.isError()) {
            return Err<void>(xmlContentResult.error().getCode(),
                           "Failed to generate XML: " + xmlContentResult.error().getMessage());
        }

        std::vector<uint8_t> xmlData(xmlContentResult.value().begin(), xmlContentResult.value().end());
        auto writeResult = zipWriter.write(partName(), xmlData);
        if (writeResult.isError()) {
            return Err<void>(writeResult.error().getCode(),
                           "Failed to write " + partName() + ": " + writeResult.error().getMessage());
        }

        return Ok();
    }

    std::string TXChartRelsXmlHandler::partName() const
    {
        return "xl/charts/_rels/chart" + std::to_string(m_chartIndex + 1) + ".xml.rels";
    }

    // ==================== TXDrawingXmlHandler 实现 ====================

    TXDrawingXmlHandler::TXDrawingXmlHandler(u32 sheetIndex)
        : m_sheetIndex(sheetIndex)
    {
    }

    TXResult<void> TXDrawingXmlHandler::load(TXZipArchiveReader& /*zipReader*/, TXWorkbookContext& /*context*/)
    {
        // 暂不实现
        return Err<void>(TXErrorCode::InvalidArgument, "Drawing loading not implemented");
    }

    TXResult<void> TXDrawingXmlHandler::save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context)
    {
        if (m_sheetIndex >= context.sheets.size()) {
            return Err<void>(TXErrorCode::InvalidArgument, "Invalid sheet index");
        }

        const TXSheet* sheet = context.sheets[m_sheetIndex].get();
        auto charts = sheet->getAllCharts();

        if (charts.empty()) {
            // 如果没有图表，不需要生成drawing.xml
            return Ok();
        }

        XmlNodeBuilder drawing("xdr:wsDr");
        drawing.addAttribute("xmlns:xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
               .addAttribute("xmlns:a", "http://schemas.openxmlformats.org/drawingml/2006/main");

        // 为每个图表生成绘图对象
        for (size_t i = 0; i < charts.size(); ++i) {
            // const TXChart* chart = charts[i]; // 暂时不需要使用chart对象

            XmlNodeBuilder twoCellAnchor("xdr:twoCellAnchor");
            twoCellAnchor.addAttribute("editAs", "oneCell");

            // 起始位置
            XmlNodeBuilder from("xdr:from");
            from.addChild(XmlNodeBuilder("xdr:col").setText("1"));
            from.addChild(XmlNodeBuilder("xdr:colOff").setText("0"));
            from.addChild(XmlNodeBuilder("xdr:row").setText("1"));
            from.addChild(XmlNodeBuilder("xdr:rowOff").setText("0"));
            twoCellAnchor.addChild(from);

            // 结束位置
            XmlNodeBuilder to("xdr:to");
            to.addChild(XmlNodeBuilder("xdr:col").setText("8"));
            to.addChild(XmlNodeBuilder("xdr:colOff").setText("0"));
            to.addChild(XmlNodeBuilder("xdr:row").setText("15"));
            to.addChild(XmlNodeBuilder("xdr:rowOff").setText("0"));
            twoCellAnchor.addChild(to);

            // 图表对象
            XmlNodeBuilder graphicFrame("xdr:graphicFrame");
            graphicFrame.addAttribute("macro", "");

            XmlNodeBuilder nvGraphicFramePr("xdr:nvGraphicFramePr");
            XmlNodeBuilder cNvPr("xdr:cNvPr");
            cNvPr.addAttribute("id", std::to_string(i + 2));
            cNvPr.addAttribute("name", "Chart " + std::to_string(i + 1));
            nvGraphicFramePr.addChild(cNvPr);

            XmlNodeBuilder cNvGraphicFramePr("xdr:cNvGraphicFramePr");
            nvGraphicFramePr.addChild(cNvGraphicFramePr);

            graphicFrame.addChild(nvGraphicFramePr);

            // 变换
            XmlNodeBuilder xfrm("xdr:xfrm");
            XmlNodeBuilder off("a:off");
            off.addAttribute("x", "0").addAttribute("y", "0");
            xfrm.addChild(off);
            XmlNodeBuilder ext("a:ext");
            ext.addAttribute("cx", "0").addAttribute("cy", "0");
            xfrm.addChild(ext);
            graphicFrame.addChild(xfrm);

            // 图形
            XmlNodeBuilder graphic("a:graphic");
            XmlNodeBuilder graphicData("a:graphicData");
            graphicData.addAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            XmlNodeBuilder chartRef("c:chart");
            chartRef.addAttribute("xmlns:c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            chartRef.addAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            chartRef.addAttribute("r:id", "rId" + std::to_string(i + 1));

            graphicData.addChild(chartRef);
            graphic.addChild(graphicData);
            graphicFrame.addChild(graphic);

            twoCellAnchor.addChild(graphicFrame);

            // 客户端数据
            XmlNodeBuilder clientData("xdr:clientData");
            twoCellAnchor.addChild(clientData);

            drawing.addChild(twoCellAnchor);
        }

        TXXmlWriter writer;
        auto setRootResult = writer.setRootNode(drawing);
        if (setRootResult.isError()) {
            return Err<void>(setRootResult.error().getCode(),
                           "Failed to set root node: " + setRootResult.error().getMessage());
        }

        auto xmlContentResult = writer.generateXmlString();
        if (xmlContentResult.isError()) {
            return Err<void>(xmlContentResult.error().getCode(),
                           "Failed to generate XML: " + xmlContentResult.error().getMessage());
        }

        std::vector<uint8_t> xmlData(xmlContentResult.value().begin(), xmlContentResult.value().end());
        auto writeResult = zipWriter.write(partName(), xmlData);
        if (writeResult.isError()) {
            return Err<void>(writeResult.error().getCode(),
                           "Failed to write " + partName() + ": " + writeResult.error().getMessage());
        }

        return Ok();
    }

    std::string TXDrawingXmlHandler::partName() const
    {
        return "xl/drawings/drawing" + std::to_string(m_sheetIndex + 1) + ".xml";
    }

    // ==================== TXDrawingRelsXmlHandler 实现 ====================

    TXDrawingRelsXmlHandler::TXDrawingRelsXmlHandler(u32 sheetIndex)
        : m_sheetIndex(sheetIndex)
    {
    }

    TXResult<void> TXDrawingRelsXmlHandler::load(TXZipArchiveReader& /*zipReader*/, TXWorkbookContext& /*context*/)
    {
        // 暂不实现
        return Err<void>(TXErrorCode::InvalidArgument, "Drawing rels loading not implemented");
    }

    TXResult<void> TXDrawingRelsXmlHandler::save(TXZipArchiveWriter& zipWriter, const TXWorkbookContext& context)
    {
        if (m_sheetIndex >= context.sheets.size()) {
            return Err<void>(TXErrorCode::InvalidArgument, "Invalid sheet index");
        }

        const TXSheet* sheet = context.sheets[m_sheetIndex].get();
        auto charts = sheet->getAllCharts();

        // 使用流式写入器生成XML
        TXBufferedXmlWriter writer(4096); // 小文件，4KB缓冲区足够

        // 写入XML声明
        const char* xmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
        writer.write(xmlDecl, strlen(xmlDecl));

        // 写入根元素开始标签
        const char* rootStart = "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
        writer.write(rootStart, strlen(rootStart));

        // 为每个图表写入关系
        for (size_t i = 0; i < charts.size(); ++i) {
            std::string relationship = "<Relationship Id=\"rId" + std::to_string(i + 1) +
                                     "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\"" +
                                     " Target=\"../charts/chart" + std::to_string(i + 1) + ".xml\"/>";
            writer.write(relationship.c_str(), relationship.length());
        }

        // 写入根元素结束标签
        const char* rootEnd = "</Relationships>";
        writer.write(rootEnd, strlen(rootEnd));

        // 写入到ZIP文件
        auto writeResult = zipWriter.write(partName(), writer.getBuffer());
        if (writeResult.isError()) {
            return Err<void>(writeResult.error().getCode(),
                           "Failed to write " + partName() + ": " + writeResult.error().getMessage());
        }

        return Ok();
    }

    std::string TXDrawingRelsXmlHandler::partName() const
    {
        return "xl/drawings/_rels/drawing" + std::to_string(m_sheetIndex + 1) + ".xml.rels";
    }

} // namespace TinaXlsx
