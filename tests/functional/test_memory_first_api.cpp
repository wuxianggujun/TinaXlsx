//
// @file test_memory_first_api.cpp
// @brief 内存优先架构功能测试 - GTest风格
//

#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <vector>
#include <chrono>
#include <iomanip>

using namespace TinaXlsx;

class MemoryFirstAPITest : public ::testing::Test {
protected:
    void SetUp() override {
        // 测试前准备
    }
    
    void TearDown() override {
        // 测试后清理
    }
};

/**
 * @brief 测试快速创建数值表格
 */
TEST_F(MemoryFirstAPITest, QuickNumbersCreation) {
    // 准备数据 - 1000行 x 10列 = 10,000个数值单元格
    std::vector<std::vector<double>> data;
    for (int row = 0; row < 1000; ++row) {
        std::vector<double> row_data;
        for (int col = 0; col < 10; ++col) {
            row_data.push_back(row * 10 + col);
        }
        data.push_back(std::move(row_data));
    }
    
    auto start = std::chrono::high_resolution_clock::now();
    
    // 使用快速API创建Excel文件
    auto result = QuickExcel::createFromNumbers(data, "test_quick_numbers.xlsx");
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    ASSERT_TRUE(result.isSuccess()) << "创建Excel失败: " << result.getError().getMessage();
    
    // 性能断言 - 10,000单元格应该在合理时间内完成
    EXPECT_LT(duration.count(), 100) << "10,000单元格创建耗时超过100ms";
    
    std::cout << "✓ 10,000单元格创建耗时: " << duration.count() << "ms" << std::endl;
    std::cout << "  性能: " << (10000.0 / duration.count()) << " 单元格/ms" << std::endl;
}

/**
 * @brief 测试混合数据类型处理
 */
TEST_F(MemoryFirstAPITest, MixedDataCreation) {
    // 准备混合数据
    std::vector<std::vector<TXVariant>> data = {
        {TXVariant("产品名称"), TXVariant("价格"), TXVariant("库存"), TXVariant("是否促销")},
        {TXVariant("苹果"), TXVariant(12.5), TXVariant(100), TXVariant(true)},
        {TXVariant("香蕉"), TXVariant(8.0), TXVariant(50), TXVariant(false)},
        {TXVariant("橙子"), TXVariant(15.0), TXVariant(75), TXVariant(true)}
    };
    
    auto start = std::chrono::high_resolution_clock::now();
    
    auto result = QuickExcel::createFromData(data, "test_mixed_data.xlsx");
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    ASSERT_TRUE(result.isSuccess()) << "混合数据创建失败: " << result.getError().getMessage();
    
    // 小数据集应该很快完成
    EXPECT_LT(duration.count(), 10000) << "小数据集处理耗时超过10ms";
    
    std::cout << "✓ 混合数据创建耗时: " << duration.count() << "μs" << std::endl;
}

/**
 * @brief 测试内存优先工作簿高级用法
 */
TEST_F(MemoryFirstAPITest, MemoryWorkbookAdvanced) {
    auto start = std::chrono::high_resolution_clock::now();
    
    // 创建内存优先工作簿
    auto workbook = TXInMemoryWorkbook::create("test_advanced_demo.xlsx");
    ASSERT_NE(workbook, nullptr) << "工作簿创建失败";
    
    // 创建工作表
    auto& sheet1 = workbook->createSheet("销售数据");
    auto& sheet2 = workbook->createSheet("统计汇总");
    
    // 批量设置数据到第一个工作表
    std::vector<TXCoordinate> coords;
    std::vector<double> values;
    
    // 生成销售数据
    for (int month = 1; month <= 12; ++month) {
        coords.emplace_back(month, 0);  // A列: 月份
        values.push_back(month);
        
        coords.emplace_back(month, 1);  // B列: 销售额
        values.push_back(10000 + month * 500 + (rand() % 1000));
    }
    
    // 批量设置数值
    auto result1 = sheet1.setBatchNumbers(coords, values);
    ASSERT_TRUE(result1.isSuccess()) << "批量设置数值失败";
    EXPECT_EQ(result1.getValue(), 24) << "应该设置24个数值单元格";
    
    // 设置标题行
    std::vector<TXCoordinate> title_coords = {{0, 0}, {0, 1}};
    std::vector<std::string> titles = {"月份", "销售额"};
    auto result2 = sheet1.setBatchStrings(title_coords, titles);
    ASSERT_TRUE(result2.isSuccess()) << "设置标题失败";
    EXPECT_EQ(result2.getValue(), 2) << "应该设置2个标题单元格";
    
    // 计算统计信息
    TXRange data_range(1, 1, 12, 1);  // B2:B13
    auto stats = sheet1.getStats(&data_range);
    EXPECT_EQ(stats.number_cells, 12) << "应该有12个数值单元格";
    EXPECT_GT(stats.sum, 0) << "总和应该大于0";
    EXPECT_GT(stats.mean, 0) << "平均值应该大于0";
    
    // 添加汇总数据到第二个工作表
    std::vector<TXCoordinate> summary_coords = {{0, 0}, {0, 1}, {1, 0}, {1, 1}};
    std::vector<TXVariant> summary_data = {
        TXVariant("项目"), TXVariant("数值"),
        TXVariant("年度总销售额"), TXVariant(stats.sum)
    };
    auto result3 = sheet2.setBatchMixed(summary_coords, summary_data);
    ASSERT_TRUE(result3.isSuccess()) << "设置汇总数据失败";
    
    // 保存文件
    auto save_result = workbook->saveToFile();
    ASSERT_TRUE(save_result.isSuccess()) << "保存文件失败: " << save_result.getError().getMessage();
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    // 性能要求：复杂操作应在合理时间内完成
    EXPECT_LT(duration.count(), 50) << "高级工作簿操作耗时超过50ms";
    
    std::cout << "✓ 高级工作簿操作耗时: " << duration.count() << "ms" << std::endl;
}

/**
 * @brief 测试CSV导入功能
 */
TEST_F(MemoryFirstAPITest, CSVImport) {
    // 模拟CSV数据
    std::string csv_content = 
        "姓名,年龄,部门,工资\n"
        "张三,28,技术部,8000\n"
        "李四,32,销售部,7500\n"
        "王五,25,市场部,6500\n"
        "赵六,30,人事部,7000\n";
    
    auto start = std::chrono::high_resolution_clock::now();
    
    auto result = QuickExcel::createFromCSV(csv_content, "test_employee_data.xlsx");
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    ASSERT_TRUE(result.isSuccess()) << "CSV导入失败: " << result.getError().getMessage();
    
    // CSV导入应该很快
    EXPECT_LT(duration.count(), 5000) << "CSV导入耗时超过5ms";
    
    std::cout << "✓ CSV导入耗时: " << duration.count() << "μs" << std::endl;
}

/**
 * @brief 测试2ms挑战 - 10,000单元格性能目标
 */
TEST_F(MemoryFirstAPITest, TwoMillisecondChallenge) {
    const size_t ROWS = 100;
    const size_t COLS = 100;  // 100x100 = 10,000单元格
    
    // 准备数据
    std::vector<std::vector<double>> data;
    data.reserve(ROWS);
    
    for (size_t row = 0; row < ROWS; ++row) {
        std::vector<double> row_data;
        row_data.reserve(COLS);
        for (size_t col = 0; col < COLS; ++col) {
            row_data.push_back(row * COLS + col + 3.14159);
        }
        data.push_back(std::move(row_data));
    }
    
    auto start = std::chrono::high_resolution_clock::now();
    
    auto result = QuickExcel::createFromNumbers(data, "test_2ms_challenge.xlsx");
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    double duration_ms = duration.count() / 1000.0;
    
    ASSERT_TRUE(result.isSuccess()) << "2ms挑战失败: " << result.getError().getMessage();
    
    // 严格的性能要求 - 目标是2ms以内
    EXPECT_LT(duration_ms, 5.0) << "10,000单元格处理应在5ms内完成 (目标2ms)";
    
    std::cout << "🚀 2ms挑战结果: " << std::fixed << std::setprecision(3) 
              << duration_ms << "ms" << std::endl;
    std::cout << "   性能: " << (10000.0 / duration_ms) << " 单元格/ms" << std::endl;
    
    if (duration_ms <= 2.0) {
        std::cout << "🎉 恭喜！达成2ms挑战目标！" << std::endl;
    } else if (duration_ms <= 3.0) {
        std::cout << "👏 很棒！接近2ms目标！" << std::endl;  
    } else {
        std::cout << "⚠️  还需要继续优化以达到2ms目标" << std::endl;
    }
}

/**
 * @brief 测试API易用性 - 最简单的使用场景
 */
TEST_F(MemoryFirstAPITest, SimpleUsageAPI) {
    // 最简单的使用 - 单个值
    auto result1 = QuickExcel::createSingle(42.0, "test_single.xlsx");
    EXPECT_TRUE(result1.isSuccess()) << "单值创建失败";
    
    // 简单的一维数组
    std::vector<double> simple_data = {1.0, 2.0, 3.0, 4.0, 5.0};
    auto result2 = QuickExcel::createFromVector(simple_data, "test_vector.xlsx");
    EXPECT_TRUE(result2.isSuccess()) << "一维数组创建失败";
    
    std::cout << "✓ 简单API测试通过" << std::endl;
} 