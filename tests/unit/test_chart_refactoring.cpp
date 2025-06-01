#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "TinaXlsx/TXChartSeriesBuilder.hpp"
#include "TinaXlsx/TXAxisBuilder.hpp"
#include "TinaXlsx/TXRangeFormatter.hpp"
#include "test_file_generator.hpp"

using namespace TinaXlsx;

class ChartRefactoringTest : public TestWithFileGeneration<ChartRefactoringTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<ChartRefactoringTest>::SetUp();
        TinaXlsx::initialize();
    }

    void TearDown() override {
        TinaXlsx::cleanup();
        TestWithFileGeneration<ChartRefactoringTest>::TearDown();
    }
};

TEST_F(ChartRefactoringTest, RangeFormatterTest) {
    // 测试数据范围格式化工具
    TXRange range = TXRange::fromAddress("A6:B10");
    std::string sheetName = "测试工作表";

    // 测试类别轴范围格式化
    std::string catRange = TXRangeFormatter::formatCategoryRange(range, sheetName);
    EXPECT_EQ(catRange, "'测试工作表'!$A$7:$A$10");

    // 测试数值范围格式化
    std::string valRange = TXRangeFormatter::formatValueRange(range, sheetName);
    EXPECT_EQ(valRange, "'测试工作表'!$B$7:$B$10");

    // 测试散点图范围格式化
    std::string xRange = TXRangeFormatter::formatScatterXRange(range, sheetName);
    std::string yRange = TXRangeFormatter::formatScatterYRange(range, sheetName);
    EXPECT_EQ(xRange, "'测试工作表'!$A$7:$A$10");
    EXPECT_EQ(yRange, "'测试工作表'!$B$7:$B$10");

    std::cout << "范围格式化测试通过" << std::endl;
    std::cout << "类别轴范围: " << catRange << std::endl;
    std::cout << "数值范围: " << valRange << std::endl;
}

TEST_F(ChartRefactoringTest, SeriesBuilderFactoryTest) {
    // 测试系列构建器工厂
    auto columnBuilder = TXSeriesBuilderFactory::createBuilder(ChartType::Column);
    auto lineBuilder = TXSeriesBuilderFactory::createBuilder(ChartType::Line);
    auto pieBuilder = TXSeriesBuilderFactory::createBuilder(ChartType::Pie);
    auto scatterBuilder = TXSeriesBuilderFactory::createBuilder(ChartType::Scatter);

    EXPECT_NE(columnBuilder, nullptr);
    EXPECT_NE(lineBuilder, nullptr);
    EXPECT_NE(pieBuilder, nullptr);
    EXPECT_NE(scatterBuilder, nullptr);

    std::cout << "系列构建器工厂测试通过" << std::endl;
}

TEST_F(ChartRefactoringTest, AxisBuilderTest) {
    // 测试坐标轴构建器
    auto catAxis = TXAxisBuilder::buildCategoryAxis(1, 2);
    auto valAxis = TXAxisBuilder::buildValueAxis(2, 1, true);

    EXPECT_EQ(catAxis.getName(), "c:catAx");
    EXPECT_EQ(valAxis.getName(), "c:valAx");

    std::cout << "坐标轴构建器测试通过" << std::endl;
}

TEST_F(ChartRefactoringTest, RefactoredChartCreationTest) {
    // 测试重构后的图表创建功能
    auto workbook = createWorkbook("refactored_chart_test");
    TXSheet* sheet = workbook->addSheet("重构测试");
    ASSERT_NE(sheet, nullptr);

    // 添加测试信息
    addTestInfo(sheet, "RefactoredChartCreationTest", "测试重构后的图表创建功能");

    // 创建测试数据
    sheet->setCellValue(row_t(6), column_t(1), std::string("产品"));
    sheet->setCellValue(row_t(6), column_t(2), std::string("销量"));
    sheet->setCellValue(row_t(6), column_t(3), std::string("利润"));

    std::vector<std::string> products = {"产品A", "产品B", "产品C", "产品D"};
    std::vector<double> sales = {1200, 1500, 1100, 1800};
    std::vector<double> profits = {240, 300, 220, 360};

    for (size_t i = 0; i < products.size(); ++i) {
        row_t row = row_t(i + 7);
        sheet->setCellValue(row, column_t(1), products[i]);
        sheet->setCellValue(row, column_t(2), sales[i]);
        sheet->setCellValue(row, column_t(3), profits[i]);
    }

    // 创建不同类型的图表测试重构后的功能
    TXRange dataRange = TXRange::fromAddress("A6:B10");

    // 柱状图
    auto* columnChart = sheet->addColumnChart("重构后的柱状图", dataRange, {row_t(12), column_t(1)});
    ASSERT_NE(columnChart, nullptr);
    EXPECT_EQ(columnChart->getType(), ChartType::Column);

    // 折线图
    auto* lineChart = sheet->addLineChart("重构后的折线图", dataRange, {row_t(12), column_t(6)});
    ASSERT_NE(lineChart, nullptr);
    EXPECT_EQ(lineChart->getType(), ChartType::Line);

    // 饼图
    auto* pieChart = sheet->addPieChart("重构后的饼图", dataRange, {row_t(25), column_t(1)});
    ASSERT_NE(pieChart, nullptr);
    EXPECT_EQ(pieChart->getType(), ChartType::Pie);

    // 散点图
    auto* scatterChart = sheet->addScatterChart("重构后的散点图", dataRange, {row_t(25), column_t(6)});
    ASSERT_NE(scatterChart, nullptr);
    EXPECT_EQ(scatterChart->getType(), ChartType::Scatter);

    EXPECT_EQ(sheet->getChartCount(), 4);

    // 保存文件
    bool saved = saveWorkbook(workbook, "refactored_chart_test");
    EXPECT_TRUE(saved) << "保存失败";

    if (saved) {
        std::cout << "重构后的图表创建测试通过" << std::endl;
        std::cout << "生成了包含4种图表类型的测试文件" << std::endl;
        std::cout << "请用Excel/WPS打开验证重构后的图表是否正常显示" << std::endl;
    }
}

TEST_F(ChartRefactoringTest, CodeQualityTest) {
    // 测试代码质量改进
    auto workbook = createWorkbook("code_quality_test");
    TXSheet* sheet = workbook->addSheet("代码质量测试");
    ASSERT_NE(sheet, nullptr);

    // 添加测试信息
    addTestInfo(sheet, "CodeQualityTest", "验证重构后的代码质量和可维护性");

    // 创建复杂的测试数据
    sheet->setCellValue(row_t(6), column_t(1), std::string("月份"));
    sheet->setCellValue(row_t(6), column_t(2), std::string("收入"));

    std::vector<std::string> months = {"1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月"};
    std::vector<double> revenue = {5000, 5500, 4800, 6200, 7100, 6800, 7500, 8200};

    for (size_t i = 0; i < months.size(); ++i) {
        row_t row = row_t(i + 7);
        sheet->setCellValue(row, column_t(1), months[i]);
        sheet->setCellValue(row, column_t(2), revenue[i]);
    }

    // 测试更大的数据范围
    TXRange largeRange = TXRange::fromAddress("A6:B14");
    auto* chart = sheet->addLineChart("月度收入趋势", largeRange, {row_t(16), column_t(1)});

    ASSERT_NE(chart, nullptr);
    EXPECT_EQ(chart->getDataRange().toAddress(), "A6:B14");
    EXPECT_EQ(chart->getDataSheet()->getName(), "代码质量测试");

    // 保存文件
    bool saved = saveWorkbook(workbook, "code_quality_test");
    EXPECT_TRUE(saved) << "保存失败";

    if (saved) {
        std::cout << "代码质量测试通过" << std::endl;
        std::cout << "重构后的代码能够正确处理复杂数据" << std::endl;
    }
}
