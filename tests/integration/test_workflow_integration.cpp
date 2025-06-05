//
// @file test_workflow_integration.cpp
// @brief 完整工作流集成测试 - 端到端测试
//

#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <filesystem>
#include <fstream>

using namespace TinaXlsx;
namespace fs = std::filesystem;

class WorkflowIntegrationTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 创建测试输出目录
        test_output_dir = "test_output";
        if (!fs::exists(test_output_dir)) {
            fs::create_directory(test_output_dir);
        }
    }
    
    void TearDown() override {
        // 清理测试文件（可选）
        // fs::remove_all(test_output_dir);
    }
    
    std::string test_output_dir;
};

/**
 * @brief 测试完整的Excel创建工作流
 */
TEST_F(WorkflowIntegrationTest, CompleteExcelCreation) {
    std::string output_file = test_output_dir + "/complete_workflow.xlsx";
    
    // 1. 创建工作簿
    auto workbook = TXInMemoryWorkbook::create(output_file);
    ASSERT_NE(workbook, nullptr) << "工作簿创建失败";
    
    // 2. 创建多个工作表
    auto& sales_sheet = workbook->createSheet("销售数据");
    auto& summary_sheet = workbook->createSheet("汇总统计");
    auto& chart_sheet = workbook->createSheet("图表分析");
    
    // 3. 在销售数据表中填充数据
    std::vector<TXCoordinate> coords;
    std::vector<double> values;
    std::vector<std::string> headers = {"月份", "销售额", "利润", "增长率"};
    
    // 设置标题行
    for (size_t i = 0; i < headers.size(); ++i) {
        auto result = sales_sheet.setString(TXCoordinate(0, i), headers[i]);
        EXPECT_TRUE(result.isSuccess()) << "设置标题失败: " << headers[i];
    }
    
    // 填充12个月的数据
    for (int month = 1; month <= 12; ++month) {
        coords.clear();
        values.clear();
        
        coords.emplace_back(month, 0); // 月份
        values.push_back(month);
        
        coords.emplace_back(month, 1); // 销售额
        values.push_back(10000 + month * 1000 + (rand() % 2000));
        
        coords.emplace_back(month, 2); // 利润
        values.push_back(values.back() * 0.2 + (rand() % 500));
        
        coords.emplace_back(month, 3); // 增长率
        values.push_back((rand() % 20) - 10); // -10% 到 +10%
    }
    
    auto batch_result = sales_sheet.setBatchNumbers(coords, values);
    EXPECT_TRUE(batch_result.isSuccess()) << "批量设置数据失败";
    EXPECT_EQ(batch_result.getValue(), 48) << "应该设置48个数值单元格";
    
    // 4. 计算统计信息并填充到汇总表
    TXRange sales_range(1, 1, 12, 1); // B2:B13 销售额
    TXRange profit_range(1, 2, 12, 2); // C2:C13 利润
    
    auto sales_stats = sales_sheet.getStats(&sales_range);
    auto profit_stats = sales_sheet.getStats(&profit_range);
    
    // 在汇总表中设置统计数据
    std::vector<TXCoordinate> summary_coords = {
        {0, 0}, {0, 1}, // 标题
        {1, 0}, {1, 1}, // 总销售额
        {2, 0}, {2, 1}, // 总利润
        {3, 0}, {3, 1}, // 平均销售额
        {4, 0}, {4, 1}  // 平均利润
    };
    
    std::vector<TXVariant> summary_data = {
        TXVariant("项目"), TXVariant("数值"),
        TXVariant("总销售额"), TXVariant(sales_stats.sum),
        TXVariant("总利润"), TXVariant(profit_stats.sum),
        TXVariant("平均销售额"), TXVariant(sales_stats.mean),
        TXVariant("平均利润"), TXVariant(profit_stats.mean)
    };
    
    auto summary_result = summary_sheet.setBatchMixed(summary_coords, summary_data);
    EXPECT_TRUE(summary_result.isSuccess()) << "设置汇总数据失败";
    
    // 5. 添加公式到图表表
    auto formula_result = chart_sheet.setFormula(TXCoordinate(0, 0), "=销售数据.B2:B13");
    EXPECT_TRUE(formula_result.isSuccess()) << "设置公式失败";
    
    // 6. 保存文件
    auto save_result = workbook->saveToFile();
    EXPECT_TRUE(save_result.isSuccess()) << "保存文件失败: " << save_result.getError().getMessage();
    
    // 7. 验证文件是否存在
    EXPECT_TRUE(fs::exists(output_file)) << "输出文件不存在";
    
    // 8. 验证文件大小合理
    auto file_size = fs::file_size(output_file);
    EXPECT_GT(file_size, 1000) << "文件太小，可能有问题";
    EXPECT_LT(file_size, 1000000) << "文件太大，可能有问题";
    
    std::cout << "✅ 完整工作流测试通过" << std::endl;
    std::cout << "   - 文件: " << output_file << std::endl;
    std::cout << "   - 大小: " << file_size << " bytes" << std::endl;
    std::cout << "   - 工作表数: 3" << std::endl;
    std::cout << "   - 数据单元格: " << batch_result.getValue() << std::endl;
}

/**
 * @brief 测试大数据量工作流
 */
TEST_F(WorkflowIntegrationTest, LargeDataWorkflow) {
    std::string output_file = test_output_dir + "/large_data_workflow.xlsx";
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 创建包含大量数据的Excel
    auto workbook = TXInMemoryWorkbook::create(output_file);
    auto& sheet = workbook->createSheet("大数据表");
    
    // 生成10,000个单元格的数据
    const size_t ROWS = 100;
    const size_t COLS = 100;
    std::vector<TXCoordinate> coords;
    std::vector<double> values;
    
    coords.reserve(ROWS * COLS);
    values.reserve(ROWS * COLS);
    
    for (size_t row = 0; row < ROWS; ++row) {
        for (size_t col = 0; col < COLS; ++col) {
            coords.emplace_back(row, col);
            values.push_back(row * COLS + col + 0.5);
        }
    }
    
    auto batch_result = sheet.setBatchNumbers(coords, values);
    EXPECT_TRUE(batch_result.isSuccess()) << "大数据批量设置失败";
    EXPECT_EQ(batch_result.getValue(), ROWS * COLS) << "数据量不匹配";
    
    // 计算统计信息
    TXRange full_range(TXCoordinate(0, 0), TXCoordinate(ROWS-1, COLS-1));
    auto stats = sheet.getStats(&full_range);
    
    EXPECT_EQ(stats.count, ROWS * COLS) << "统计单元格数不正确";
    EXPECT_GT(stats.sum, 0) << "总和应该大于0";
    
    // 保存文件
    auto save_result = workbook->saveToFile();
    EXPECT_TRUE(save_result.isSuccess()) << "保存大文件失败";
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
    
    // 性能验证：10,000单元格应在合理时间内完成
    EXPECT_LT(duration.count(), 1000) << "大数据工作流应在1秒内完成";
    
    // 验证文件
    EXPECT_TRUE(fs::exists(output_file)) << "大数据文件不存在";
    auto file_size = fs::file_size(output_file);
    EXPECT_GT(file_size, 10000) << "大数据文件太小";
    
    std::cout << "✅ 大数据工作流测试通过" << std::endl;
    std::cout << "   - 处理时间: " << duration.count() << "ms" << std::endl;
    std::cout << "   - 单元格数: " << ROWS * COLS << std::endl;
    std::cout << "   - 文件大小: " << file_size << " bytes" << std::endl;
    std::cout << "   - 性能: " << (ROWS * COLS * 1000.0 / duration.count()) << " 单元格/秒" << std::endl;
}

/**
 * @brief 测试错误处理工作流
 */
TEST_F(WorkflowIntegrationTest, ErrorHandlingWorkflow) {
    // 测试无效文件路径
    auto workbook1 = TXInMemoryWorkbook::create("invalid/path/test.xlsx");
    // 应该返回错误而不是崩溃
    
    // 测试有效工作簿的错误操作
    auto workbook2 = TXInMemoryWorkbook::create(test_output_dir + "/error_test.xlsx");
    ASSERT_NE(workbook2, nullptr);
    
    auto& sheet = workbook2->createSheet("测试表");
    
    // 测试无效坐标
    auto result1 = sheet.setNumber(TXCoordinate(-1, -1), 42.0);
    EXPECT_FALSE(result1.isSuccess()) << "应该拒绝无效坐标";
    
    // 测试空数据批量操作
    std::vector<TXCoordinate> empty_coords;
    std::vector<double> empty_values;
    auto result2 = sheet.setBatchNumbers(empty_coords, empty_values);
    EXPECT_TRUE(result2.isSuccess()) << "空数据应该正常处理";
    EXPECT_EQ(result2.getValue(), 0) << "空数据应该返回0";
    
    // 测试坐标和数值数量不匹配
    std::vector<TXCoordinate> coords = {{0, 0}, {0, 1}};
    std::vector<double> values = {1.0}; // 数量不匹配
    auto result3 = sheet.setBatchNumbers(coords, values);
    EXPECT_FALSE(result3.isSuccess()) << "应该检测到数量不匹配";
    
    std::cout << "✅ 错误处理工作流测试通过" << std::endl;
}

/**
 * @brief 测试并发访问工作流（如果支持）
 */
TEST_F(WorkflowIntegrationTest, ConcurrentAccessWorkflow) {
    std::string output_file = test_output_dir + "/concurrent_test.xlsx";
    
    auto workbook = TXInMemoryWorkbook::create(output_file);
    auto& sheet = workbook->createSheet("并发测试");
    
    // 模拟并发写入（在单线程中顺序执行）
    const size_t THREAD_COUNT = 4;
    const size_t CELLS_PER_THREAD = 1000;
    
    std::vector<std::future<TXResult<size_t>>> futures;
    
    for (size_t thread_id = 0; thread_id < THREAD_COUNT; ++thread_id) {
        // 为每个"线程"准备数据
        std::vector<TXCoordinate> coords;
        std::vector<double> values;
        
        for (size_t i = 0; i < CELLS_PER_THREAD; ++i) {
            size_t row = thread_id * CELLS_PER_THREAD + i;
            coords.emplace_back(row, 0);
            values.push_back(static_cast<double>(thread_id * 1000 + i));
        }
        
        // 执行批量操作
        auto result = sheet.setBatchNumbers(coords, values);
        EXPECT_TRUE(result.isSuccess()) << "并发操作失败";
    }
    
    // 验证总数据量
    auto stats = sheet.getStats();
    EXPECT_EQ(stats.count, THREAD_COUNT * CELLS_PER_THREAD) << "并发数据总量不正确";
    
    auto save_result = workbook->saveToFile();
    EXPECT_TRUE(save_result.isSuccess()) << "并发测试文件保存失败";
    
    std::cout << "✅ 并发访问工作流测试通过" << std::endl;
    std::cout << "   - 模拟线程数: " << THREAD_COUNT << std::endl;
    std::cout << "   - 总单元格数: " << THREAD_COUNT * CELLS_PER_THREAD << std::endl;
} 