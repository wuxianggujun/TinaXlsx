//
// Created by wuxianggujun on 2025/1/15.
//

#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "test_file_generator.hpp"
#include <filesystem>

using namespace TinaXlsx;

class ChartFunctionalityTest : public TestWithFileGeneration<ChartFunctionalityTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<ChartFunctionalityTest>::SetUp();
        TinaXlsx::initialize();
    }

    void TearDown() override {
        TinaXlsx::cleanup();
        TestWithFileGeneration<ChartFunctionalityTest>::TearDown();
    }


};

TEST_F(ChartFunctionalityTest, CreateColumnChart) {
    auto workbook = createWorkbook("column_chart_test");
    TXSheet* sheet = workbook->addSheet("销售数据");
    ASSERT_NE(sheet, nullptr);

    // 添加测试信息
    addTestInfo(sheet, "CreateColumnChart", "测试柱状图创建功能");

    // 创建测试数据（从第6行开始，避免覆盖测试信息）
    sheet->setCellValue(row_t(6), column_t(1), std::string("月份"));
    sheet->setCellValue(row_t(6), column_t(2), std::string("销售额"));
    sheet->setCellValue(row_t(6), column_t(3), std::string("利润"));

    std::vector<std::string> months = {"一月", "二月", "三月", "四月", "五月", "六月"};
    std::vector<double> sales = {1000, 1200, 1100, 1300, 1500, 1400};
    std::vector<double> profits = {200, 250, 220, 280, 320, 300};

    for (size_t i = 0; i < months.size(); ++i) {
        row_t row = row_t(i + 7);
        sheet->setCellValue(row, column_t(1), months[i]);
        sheet->setCellValue(row, column_t(2), sales[i]);
        sheet->setCellValue(row, column_t(3), profits[i]);
    }

    // 创建柱状图
    TXRange dataRange = TXRange::fromAddress("A6:C12");
    auto* chart = sheet->addColumnChart("月度销售图表", dataRange, {row_t(14), column_t(1)});

    ASSERT_NE(chart, nullptr);
    EXPECT_EQ(chart->getType(), ChartType::Column);
    EXPECT_EQ(sheet->getChartCount(), 1);

    // 设置图表属性
    chart->setShowLegend(true);
    chart->setShowDataLabels(true);
    chart->setAxisTitle("月份", true);  // X轴
    chart->setAxisTitle("金额", false); // Y轴

    // 保存文件
    bool saved = saveWorkbook(workbook, "column_chart_test");
    EXPECT_TRUE(saved) << "保存失败";

    if (saved) {
        std::cout << "柱状图测试文件已生成，请用Excel/WPS打开验证图表是否正确显示" << std::endl;
    }
}

TEST_F(ChartFunctionalityTest, CreateLineChart) {
    auto workbook = createWorkbook("line_chart_test");
    TXSheet* sheet = workbook->addSheet("趋势分析");
    ASSERT_NE(sheet, nullptr);

    // 添加测试信息
    addTestInfo(sheet, "CreateLineChart", "测试折线图创建功能");

    // 创建测试数据
    sheet->setCellValue(row_t(6), column_t(1), std::string("月份"));
    sheet->setCellValue(row_t(6), column_t(2), std::string("销售额"));

    std::vector<std::string> months = {"一月", "二月", "三月", "四月", "五月", "六月"};
    std::vector<double> sales = {1000, 1200, 1100, 1300, 1500, 1400};

    for (size_t i = 0; i < months.size(); ++i) {
        row_t row = row_t(i + 7);
        sheet->setCellValue(row, column_t(1), months[i]);
        sheet->setCellValue(row, column_t(2), sales[i]);
    }

    // 创建折线图
    TXRange dataRange = TXRange::fromAddress("A6:B12");
    auto* chart = sheet->addLineChart("销售趋势图", dataRange, {row_t(14), column_t(1)});

    ASSERT_NE(chart, nullptr);
    EXPECT_EQ(chart->getType(), ChartType::Line);

    // 设置折线图特有属性
    auto* lineChart = static_cast<TXLineChart*>(chart);
    lineChart->setSmoothLines(true);
    lineChart->setShowMarkers(true);

    // 保存文件
    bool saved = saveWorkbook(workbook, "line_chart_test");
    EXPECT_TRUE(saved) << "保存失败";

    if (saved) {
        std::cout << "折线图测试文件已生成" << std::endl;
    }
}

TEST_F(ChartFunctionalityTest, CreatePieChart) {
    auto workbook = createWorkbook("pie_chart_test");
    TXSheet* sheet = workbook->addSheet("市场份额");
    ASSERT_NE(sheet, nullptr);

    // 添加测试信息
    addTestInfo(sheet, "CreatePieChart", "测试饼图创建功能");

    // 创建饼图数据
    sheet->setCellValue(row_t(6), column_t(1), std::string("产品"));
    sheet->setCellValue(row_t(6), column_t(2), std::string("份额"));

    std::vector<std::string> products = {"产品A", "产品B", "产品C", "产品D"};
    std::vector<double> shares = {35.5, 28.3, 22.1, 14.1};

    for (size_t i = 0; i < products.size(); ++i) {
        row_t row = row_t(i + 7);
        sheet->setCellValue(row, column_t(1), products[i]);
        sheet->setCellValue(row, column_t(2), shares[i]);
    }

    // 创建饼图
    TXRange dataRange = TXRange::fromAddress("A6:B10");
    auto* chart = sheet->addPieChart("市场份额分布", dataRange, {row_t(12), column_t(1)});

    ASSERT_NE(chart, nullptr);
    EXPECT_EQ(chart->getType(), ChartType::Pie);

    // 设置饼图特有属性
    auto* pieChart = static_cast<TXPieChart*>(chart);
    pieChart->setFirstSliceAngle(90);
    pieChart->setExplodeSlice(0, true); // 突出第一个扇形

    // 保存文件
    bool saved = saveWorkbook(workbook, "pie_chart_test");
    EXPECT_TRUE(saved) << "保存失败";

    if (saved) {
        std::cout << "饼图测试文件已生成" << std::endl;
    }
}

TEST_F(ChartFunctionalityTest, CreateScatterChart) {
    auto workbook = createWorkbook("scatter_chart_test");
    TXSheet* sheet = workbook->addSheet("相关性分析");
    ASSERT_NE(sheet, nullptr);

    // 添加测试信息
    addTestInfo(sheet, "CreateScatterChart", "测试散点图创建功能");

    // 创建散点图数据
    sheet->setCellValue(row_t(6), column_t(1), std::string("X值"));
    sheet->setCellValue(row_t(6), column_t(2), std::string("Y值"));

    // 生成一些相关性数据
    for (int i = 1; i <= 10; ++i) {
        row_t row = row_t(i + 6);
        double x = i * 2.0;
        double y = x * 1.5 + (i % 3 - 1) * 5; // 添加一些噪声
        sheet->setCellValue(row, column_t(1), x);
        sheet->setCellValue(row, column_t(2), y);
    }

    // 创建散点图
    TXRange dataRange = TXRange::fromAddress("A6:B16");
    auto* chart = sheet->addScatterChart("相关性分析图", dataRange, {row_t(18), column_t(1)});

    ASSERT_NE(chart, nullptr);
    EXPECT_EQ(chart->getType(), ChartType::Scatter);

    // 设置散点图特有属性
    auto* scatterChart = static_cast<TXScatterChart*>(chart);
    scatterChart->setShowTrendLine(true);
    scatterChart->setTrendLineType(TXScatterChart::TrendLineType::Linear);

    // 保存文件
    bool saved = saveWorkbook(workbook, "scatter_chart_test");
    EXPECT_TRUE(saved) << "保存失败";

    if (saved) {
        std::cout << "散点图测试文件已生成" << std::endl;
    }
}

TEST_F(ChartFunctionalityTest, MultipleChartsTest) {
    auto workbook = createWorkbook("multiple_charts_test");
    TXSheet* sheet = workbook->addSheet("多图表测试");
    ASSERT_NE(sheet, nullptr);

    // 添加测试信息
    addTestInfo(sheet, "MultipleChartsTest", "测试在一个工作表中创建多个图表");

    // 创建测试数据
    sheet->setCellValue(row_t(6), column_t(1), std::string("月份"));
    sheet->setCellValue(row_t(6), column_t(2), std::string("销售额"));

    std::vector<std::string> months = {"一月", "二月", "三月", "四月"};
    std::vector<double> sales = {1000, 1200, 1100, 1300};

    for (size_t i = 0; i < months.size(); ++i) {
        row_t row = row_t(i + 7);
        sheet->setCellValue(row, column_t(1), months[i]);
        sheet->setCellValue(row, column_t(2), sales[i]);
    }

    // 创建两个图表
    TXRange dataRange = TXRange::fromAddress("A6:B10");

    auto* columnChart = sheet->addColumnChart("销售柱状图", dataRange, {row_t(12), column_t(1)});
    auto* lineChart = sheet->addLineChart("销售趋势图", dataRange, {row_t(12), column_t(6)});

    ASSERT_NE(columnChart, nullptr);
    ASSERT_NE(lineChart, nullptr);
    EXPECT_EQ(sheet->getChartCount(), 2);

    // 保存文件
    bool saved = saveWorkbook(workbook, "multiple_charts_test");
    EXPECT_TRUE(saved) << "保存失败";

    if (saved) {
        std::cout << "多图表测试文件已生成" << std::endl;
    }
}
