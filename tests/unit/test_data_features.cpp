#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "TinaXlsx/TXFormulaBuilder.hpp"
#include "TinaXlsx/TXDataValidation.hpp"
#include "TinaXlsx/TXDataFilter.hpp"
#include "test_file_generator.hpp"

using namespace TinaXlsx;

class DataFeaturesTest : public TestWithFileGeneration<DataFeaturesTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<DataFeaturesTest>::SetUp();
        TinaXlsx::initialize();
    }

    void TearDown() override {
        TinaXlsx::cleanup();
        TestWithFileGeneration<DataFeaturesTest>::TearDown();
    }
};

TEST_F(DataFeaturesTest, FormulaBuilderTest) {
    // 测试公式构建器
    std::cout << "=== 公式构建器测试 ===" << std::endl;

    // 测试统计函数
    TXRange range = TXRange::fromAddress("B2:B10");
    
    std::string sumFormula = TXFormulaBuilder::sum(range);
    EXPECT_EQ(sumFormula, "=SUM(B2:B10)");
    std::cout << "SUM公式: " << sumFormula << std::endl;

    std::string avgFormula = TXFormulaBuilder::average(range);
    EXPECT_EQ(avgFormula, "=AVERAGE(B2:B10)");
    std::cout << "AVERAGE公式: " << avgFormula << std::endl;

    std::string maxFormula = TXFormulaBuilder::max(range);
    EXPECT_EQ(maxFormula, "=MAX(B2:B10)");
    std::cout << "MAX公式: " << maxFormula << std::endl;

    // 测试条件函数
    std::string sumIfFormula = TXFormulaBuilder::sumIf(range, ">100");
    EXPECT_EQ(sumIfFormula, "=SUMIF(B2:B10,\">100\")");
    std::cout << "SUMIF公式: " << sumIfFormula << std::endl;

    std::string countIfFormula = TXFormulaBuilder::countIf(range, ">=50");
    EXPECT_EQ(countIfFormula, "=COUNTIF(B2:B10,\">=50\")");
    std::cout << "COUNTIF公式: " << countIfFormula << std::endl;

    // 测试文本函数
    std::string concatFormula = TXFormulaBuilder::concatenate({"A1", "B1", "C1"});
    EXPECT_EQ(concatFormula, "=CONCATENATE(A1,B1,C1)");
    std::cout << "CONCATENATE公式: " << concatFormula << std::endl;

    // 测试日期函数
    std::string todayFormula = TXFormulaBuilder::today();
    EXPECT_EQ(todayFormula, "=TODAY()");
    std::cout << "TODAY公式: " << todayFormula << std::endl;

    std::cout << "✅ 公式构建器测试通过" << std::endl;
}

TEST_F(DataFeaturesTest, DataValidationTest) {
    // 测试数据验证
    std::cout << "=== 数据验证测试 ===" << std::endl;

    // 测试整数验证
    auto intValidation = TXDataValidation::createIntegerValidation(1, 100);
    EXPECT_EQ(intValidation.getType(), DataValidationType::Whole);
    EXPECT_EQ(intValidation.getOperator(), DataValidationOperator::Between);
    EXPECT_EQ(intValidation.getFormula1(), "1");
    EXPECT_EQ(intValidation.getFormula2(), "100");
    std::cout << "✅ 整数验证创建成功" << std::endl;

    // 测试列表验证
    std::vector<std::string> listItems = {"优秀", "良好", "一般", "差"};
    auto listValidation = TXDataValidation::createListValidation(listItems);
    EXPECT_EQ(listValidation.getType(), DataValidationType::List);
    EXPECT_TRUE(listValidation.getShowDropDown());
    EXPECT_EQ(listValidation.getListItems().size(), 4);
    std::cout << "✅ 列表验证创建成功" << std::endl;

    // 测试小数验证
    auto decimalValidation = TXDataValidation::createDecimalValidation(0.0, 100.0);
    EXPECT_EQ(decimalValidation.getType(), DataValidationType::Decimal);
    std::cout << "✅ 小数验证创建成功" << std::endl;

    // 测试文本长度验证
    auto textValidation = TXDataValidation::createTextLengthValidation(5, 20);
    EXPECT_EQ(textValidation.getType(), DataValidationType::TextLength);
    std::cout << "✅ 文本长度验证创建成功" << std::endl;

    std::cout << "✅ 数据验证测试通过" << std::endl;
}

TEST_F(DataFeaturesTest, DataFilterTest) {
    // 测试数据筛选
    std::cout << "=== 数据筛选测试 ===" << std::endl;

    TXRange dataRange = TXRange::fromAddress("A1:D10");
    
    // 测试自动筛选
    TXAutoFilter autoFilter(dataRange);
    EXPECT_EQ(autoFilter.getRange().toAddress(), "A1:D10");
    EXPECT_TRUE(autoFilter.getShowFilterButtons());
    std::cout << "✅ 自动筛选创建成功" << std::endl;

    // 测试筛选条件
    autoFilter.setTextFilter(0, "产品A", FilterOperator::Contains);
    autoFilter.setNumberFilter(1, 100, FilterOperator::GreaterThan);
    autoFilter.setRangeFilter(2, 50, 200);  // 这个方法会添加2个条件

    EXPECT_EQ(autoFilter.getFilterConditions().size(), 4);  // 1+1+2=4个条件
    std::cout << "✅ 筛选条件设置成功" << std::endl;

    // 测试数据排序器
    TXDataSorter sorter(dataRange);
    sorter.setHasHeader(true);
    sorter.sortByColumn(1, SortOrder::Descending);
    
    EXPECT_EQ(sorter.getSortConditions().size(), 1);
    EXPECT_EQ(sorter.getSortConditions()[0].columnIndex, 1);
    EXPECT_EQ(sorter.getSortConditions()[0].order, SortOrder::Descending);
    std::cout << "✅ 数据排序器测试成功" << std::endl;

    // 测试数据表格
    TXDataTable dataTable(dataRange, true);
    auto& filter = dataTable.enableAutoFilter();
    filter.setTextFilter(0, "测试", FilterOperator::Contains);
    
    EXPECT_TRUE(dataTable.hasAutoFilter());
    std::cout << "✅ 数据表格测试成功" << std::endl;

    std::cout << "✅ 数据筛选测试通过" << std::endl;
}

TEST_F(DataFeaturesTest, FormulaFileGenerationTest) {
    // 测试公式功能文件生成
    auto workbook = createWorkbook("formula_test");
    TXSheet* sheet = workbook->addSheet("公式测试");
    ASSERT_NE(sheet, nullptr);

    // 添加测试信息
    addTestInfo(sheet, "FormulaFileGenerationTest", "测试扩展公式函数库的文件生成");

    // 创建测试数据
    sheet->setCellValue(row_t(6), column_t(1), std::string("产品"));
    sheet->setCellValue(row_t(6), column_t(2), std::string("销量"));
    sheet->setCellValue(row_t(6), column_t(3), std::string("单价"));
    sheet->setCellValue(row_t(6), column_t(4), std::string("总额"));

    std::vector<std::string> products = {"产品A", "产品B", "产品C", "产品D", "产品E"};
    std::vector<double> sales = {120, 85, 200, 150, 95};
    std::vector<double> prices = {25.5, 30.0, 18.8, 22.3, 28.9};

    for (size_t i = 0; i < products.size(); ++i) {
        row_t row = row_t(static_cast<u32>(i + 7));
        sheet->setCellValue(row, column_t(1), products[i]);
        sheet->setCellValue(row, column_t(2), sales[i]);
        sheet->setCellValue(row, column_t(3), prices[i]);
        
        // 使用公式计算总额
        std::string formula = "=B" + std::to_string(row.index()) + "*C" + std::to_string(row.index());
        sheet->setCellFormula(row, column_t(4), formula);
    }

    // 添加统计公式
    row_t statsRow = row_t(13);
    sheet->setCellValue(statsRow, column_t(1), std::string("统计"));
    
    // 使用公式构建器创建统计公式
    TXRange salesRange = TXRange::fromAddress("B7:B11");
    TXRange priceRange = TXRange::fromAddress("C7:C11");
    TXRange totalRange = TXRange::fromAddress("D7:D11");

    sheet->setCellValue(row_t(14), column_t(1), std::string("总销量"));
    sheet->setCellFormula(row_t(14), column_t(2), TXFormulaBuilder::sum(salesRange));

    sheet->setCellValue(row_t(15), column_t(1), std::string("平均单价"));
    sheet->setCellFormula(row_t(15), column_t(2), TXFormulaBuilder::average(priceRange));

    sheet->setCellValue(row_t(16), column_t(1), std::string("最大总额"));
    sheet->setCellFormula(row_t(16), column_t(2), TXFormulaBuilder::max(totalRange));

    sheet->setCellValue(row_t(17), column_t(1), std::string("最小总额"));
    sheet->setCellFormula(row_t(17), column_t(2), TXFormulaBuilder::min(totalRange));

    sheet->setCellValue(row_t(18), column_t(1), std::string("高销量产品数"));
    sheet->setCellFormula(row_t(18), column_t(2), TXFormulaBuilder::countIf(salesRange, ">100"));

    sheet->setCellValue(row_t(19), column_t(1), std::string("高销量总额"));
    sheet->setCellFormula(row_t(19), column_t(2), TXFormulaBuilder::sumIf(salesRange, ">100", totalRange));

    // 保存文件
    bool saved = saveWorkbook(workbook, "formula_test");
    EXPECT_TRUE(saved) << "保存失败";

    if (saved) {
        std::cout << "✅ 公式功能文件生成测试通过" << std::endl;
        std::cout << "生成了包含扩展公式功能的测试文件：" << std::endl;
        std::cout << "  - 基础计算公式：B*C" << std::endl;
        std::cout << "  - 统计函数：SUM, AVERAGE, MAX, MIN" << std::endl;
        std::cout << "  - 条件函数：COUNTIF, SUMIF" << std::endl;
    }
}

TEST_F(DataFeaturesTest, DataValidationFileGenerationTest) {
    // 测试数据验证文件生成
    auto workbook = createWorkbook("data_validation_test");
    TXSheet* sheet = workbook->addSheet("数据验证测试");
    ASSERT_NE(sheet, nullptr);

    // 添加测试信息
    addTestInfo(sheet, "DataValidationFileGenerationTest", "测试数据验证功能的文件生成");

    // 创建数据验证示例
    sheet->setCellValue(row_t(6), column_t(1), std::string("验证类型"));
    sheet->setCellValue(row_t(6), column_t(2), std::string("输入区域"));
    sheet->setCellValue(row_t(6), column_t(3), std::string("说明"));

    sheet->setCellValue(row_t(7), column_t(1), std::string("整数验证"));
    sheet->setCellValue(row_t(7), column_t(3), std::string("请输入1-100之间的整数"));

    sheet->setCellValue(row_t(8), column_t(1), std::string("列表验证"));
    sheet->setCellValue(row_t(8), column_t(3), std::string("请从下拉列表选择"));

    sheet->setCellValue(row_t(9), column_t(1), std::string("小数验证"));
    sheet->setCellValue(row_t(9), column_t(3), std::string("请输入0.0-100.0之间的小数"));

    sheet->setCellValue(row_t(10), column_t(1), std::string("文本长度验证"));
    sheet->setCellValue(row_t(10), column_t(3), std::string("请输入5-50个字符"));

    // 实际应用数据验证规则到工作表
    auto ratingValidation = TXDataValidation::createIntegerValidation(1, 100);
    bool added1 = sheet->addDataValidation(TXRange::fromAddress("B7"), ratingValidation);
    EXPECT_TRUE(added1);
    std::cout << "✅ 整数验证规则添加: " << (added1 ? "成功" : "失败") << std::endl;

    // 方法1：创建数据源单元格（推荐方法）
    std::cout << "创建列表验证数据源..." << std::endl;

    // 在F列创建等级选项数据源
    sheet->setCellValue(row_t(12), column_t(6), std::string("等级选项"));
    sheet->setCellValue(row_t(13), column_t(6), std::string("Excellent"));
    sheet->setCellValue(row_t(14), column_t(6), std::string("Good"));
    sheet->setCellValue(row_t(15), column_t(6), std::string("Fair"));
    sheet->setCellValue(row_t(16), column_t(6), std::string("Poor"));

    // 在G列创建简单选项数据源
    sheet->setCellValue(row_t(12), column_t(7), std::string("简单选项"));
    sheet->setCellValue(row_t(13), column_t(7), std::string("A"));
    sheet->setCellValue(row_t(14), column_t(7), std::string("B"));
    sheet->setCellValue(row_t(15), column_t(7), std::string("C"));

    // 使用单元格范围引用创建列表验证（推荐方法）
    auto rangeValidation = TXDataValidation::createListValidationFromRange(TXRange::fromAddress("F13:F16"));
    bool added2 = sheet->addDataValidation(TXRange::fromAddress("B8"), rangeValidation);
    EXPECT_TRUE(added2);
    std::cout << "✅ 范围引用列表验证添加: " << (added2 ? "成功" : "失败") << std::endl;
    std::cout << "   范围引用公式: " << rangeValidation.getFormula1() << std::endl;

    // 测试另一个范围引用
    auto simpleRangeValidation = TXDataValidation::createListValidationFromRange(TXRange::fromAddress("G13:G15"));
    bool added2b = sheet->addDataValidation(TXRange::fromAddress("C8"), simpleRangeValidation);
    EXPECT_TRUE(added2b);
    std::cout << "✅ 简单范围引用列表验证添加: " << (added2b ? "成功" : "失败") << std::endl;
    std::cout << "   简单范围引用公式: " << simpleRangeValidation.getFormula1() << std::endl;

    auto discountValidation = TXDataValidation::createDecimalValidation(0.0, 100.0);
    bool added3 = sheet->addDataValidation(TXRange::fromAddress("B9"), discountValidation);
    EXPECT_TRUE(added3);
    std::cout << "✅ 小数验证规则添加: " << (added3 ? "成功" : "失败") << std::endl;

    auto commentValidation = TXDataValidation::createTextLengthValidation(5, 50);
    bool added4 = sheet->addDataValidation(TXRange::fromAddress("B10"), commentValidation);
    EXPECT_TRUE(added4);
    std::cout << "✅ 文本长度验证规则添加: " << (added4 ? "成功" : "失败") << std::endl;

    // 验证数据验证规则数量
    EXPECT_EQ(sheet->getDataValidationCount(), 5);  // 5个验证规则（增加了中文列表验证）
    std::cout << "数据验证规则总数: " << sheet->getDataValidationCount() << std::endl;

    // 保存文件
    bool saved = saveWorkbook(workbook, "data_validation_test");
    EXPECT_TRUE(saved) << "保存失败";

    if (saved) {
        std::cout << "✅ 数据验证文件生成测试通过" << std::endl;
        std::cout << "生成了包含实际数据验证规则的文件" << std::endl;
        std::cout << "🔍 验证方法:" << std::endl;
        std::cout << "- 解压xlsx文件，查看xl/worksheets/sheet1.xml" << std::endl;
        std::cout << "- 应该能看到<dataValidations>节点" << std::endl;
        std::cout << "- 用Excel打开测试:" << std::endl;
        std::cout << "  * B7: 整数验证 (1-100)" << std::endl;
        std::cout << "  * B8: 列表验证 (引用F13:F16) - Excellent,Good,Fair,Poor" << std::endl;
        std::cout << "  * C8: 列表验证 (引用G13:G15) - A,B,C" << std::endl;
        std::cout << "  * B9: 小数验证 (0.0-100.0)" << std::endl;
        std::cout << "  * B10: 文本长度验证 (5-50字符)" << std::endl;
        std::cout << "- 数据源位置:" << std::endl;
        std::cout << "  * F13:F16: Excellent, Good, Fair, Poor" << std::endl;
        std::cout << "  * G13:G15: A, B, C" << std::endl;
        std::cout << "- 使用推荐方法:" << std::endl;
        std::cout << "  * 单元格范围引用而非直接文本列表" << std::endl;
        std::cout << "  * 数据源在工作表中可见和可编辑" << std::endl;
        std::cout << "  * 符合Excel标准做法" << std::endl;
    }
}
