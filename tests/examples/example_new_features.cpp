#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <iostream>
#include <chrono>

namespace {

/**
 * @brief 新功能演示测试类
 */
class NewFeaturesExampleTest : public ::testing::Test {
protected:
    void SetUp() override {
        TinaXlsx::initialize();
    }
    
    void TearDown() override {
        TinaXlsx::cleanup();
    }
};

/**
 * @brief 测试公式功能
 */
TEST_F(NewFeaturesExampleTest, FormulaFeatures) {
    std::cout << "\n=== 公式功能演示 ===\n";
    
    // 创建工作簿和工作表
    TinaXlsx::Workbook workbook;
    auto* sheet = workbook.addSheet("公式示例");
    ASSERT_NE(sheet, nullptr);
    
    // 设置基础数据
    sheet->setCellValue(TinaXlsx::row_t(1), TinaXlsx::column_t(1), 100.0);   // A1 = 100
    sheet->setCellValue(TinaXlsx::row_t(1), TinaXlsx::column_t(2), 200.0);   // B1 = 200
    sheet->setCellValue(TinaXlsx::row_t(1), TinaXlsx::column_t(3), 300.0);   // C1 = 300
    
    // 设置公式
    bool success = sheet->setCellFormula(TinaXlsx::row_t(2), TinaXlsx::column_t(1), "SUM(A1:C1)");  // A2 = SUM(A1:C1)
    EXPECT_TRUE(success);
    
    success = sheet->setCellFormula(TinaXlsx::row_t(2), TinaXlsx::column_t(2), "AVERAGE(A1:C1)");   // B2 = AVERAGE(A1:C1)
    EXPECT_TRUE(success);
    
    success = sheet->setCellFormula(TinaXlsx::row_t(2), TinaXlsx::column_t(3), "MAX(A1:C1)");       // C2 = MAX(A1:C1)
    EXPECT_TRUE(success);
    
    // 计算公式
    std::size_t calculatedCount = sheet->calculateAllFormulas();
    EXPECT_GT(calculatedCount, 0);
    
    // 验证公式结果
    std::string formula = sheet->getCellFormula(TinaXlsx::row_t(2), TinaXlsx::column_t(1));
    EXPECT_EQ(formula, "SUM(A1:C1)");
    
    std::cout << "公式 A2: " << formula << std::endl;
    std::cout << "计算结果数量: " << calculatedCount << std::endl;
    
    // 保存文件
    success = workbook.saveToFile("formula_example.xlsx");
    EXPECT_TRUE(success);
    
    std::cout << "公式示例文件已保存: formula_example.xlsx\n";
}

/**
 * @brief 测试合并单元格功能
 */
TEST_F(NewFeaturesExampleTest, MergedCellsFeatures) {
    std::cout << "\n=== 合并单元格功能演示 ===\n";
    
    // 创建工作簿和工作表
    TinaXlsx::Workbook workbook;
    auto* sheet = workbook.addSheet("合并示例");
    ASSERT_NE(sheet, nullptr);
    
    // 设置标题
    sheet->setCellValue(TinaXlsx::row_t(1), TinaXlsx::column_t(1), "销售报表");
    
    // 合并标题单元格 A1:D1
    bool success = sheet->mergeCells(TinaXlsx::row_t(1), TinaXlsx::column_t(1), TinaXlsx::row_t(1), TinaXlsx::column_t(4));
    EXPECT_TRUE(success);
    
    // 设置表头
    sheet->setCellValue(TinaXlsx::row_t(2), TinaXlsx::column_t(1), "产品");
    sheet->setCellValue(TinaXlsx::row_t(2), TinaXlsx::column_t(2), "Q1");
    sheet->setCellValue(TinaXlsx::row_t(2), TinaXlsx::column_t(3), "Q2");
    sheet->setCellValue(TinaXlsx::row_t(2), TinaXlsx::column_t(4), "合计");
    
    // 设置数据
    sheet->setCellValue(TinaXlsx::row_t(3), TinaXlsx::column_t(1), "产品A");
    sheet->setCellValue(TinaXlsx::row_t(3), TinaXlsx::column_t(2), 1000.0);
    sheet->setCellValue(TinaXlsx::row_t(3), TinaXlsx::column_t(3), 1200.0);
    
    sheet->setCellValue(TinaXlsx::row_t(4), TinaXlsx::column_t(1), "产品B");
    sheet->setCellValue(TinaXlsx::row_t(4), TinaXlsx::column_t(2), 800.0);
    sheet->setCellValue(TinaXlsx::row_t(4), TinaXlsx::column_t(3), 900.0);
    
    // 合并产品列的小计行 A5:B5
    sheet->setCellValue(TinaXlsx::row_t(5), TinaXlsx::column_t(1), "小计");
    success = sheet->mergeCells(TinaXlsx::row_t(5), TinaXlsx::column_t(1), TinaXlsx::row_t(5), TinaXlsx::column_t(2));
    EXPECT_TRUE(success);
    
    // 检查合并状态
    bool isMerged = sheet->isCellMerged(TinaXlsx::row_t(1), TinaXlsx::column_t(1));
    EXPECT_TRUE(isMerged);
    
    isMerged = sheet->isCellMerged(TinaXlsx::row_t(1), TinaXlsx::column_t(3));  // A1:D1范围内的C1也应该被合并
    EXPECT_TRUE(isMerged);
    
    // 获取合并区域
    auto mergeRegion = sheet->getMergeRegion(TinaXlsx::row_t(1), TinaXlsx::column_t(1));
    std::cout << "合并区域: " << mergeRegion.toAddress() << std::endl;
    
    // 获取所有合并区域
    auto allRegions = sheet->getAllMergeRegions();
    std::cout << "总合并区域数量: " << allRegions.size() << std::endl;
    
    // 拆分一个合并区域
    success = sheet->unmergeCells(TinaXlsx::row_t(5), TinaXlsx::column_t(1));
    EXPECT_TRUE(success);
    
    // 验证拆分后的状态
    isMerged = sheet->isCellMerged(TinaXlsx::row_t(5), TinaXlsx::column_t(1));
    EXPECT_FALSE(isMerged);
    
    // 保存文件
    success = workbook.saveToFile("merged_cells_example.xlsx");
    EXPECT_TRUE(success);
    
    std::cout << "合并单元格示例文件已保存: merged_cells_example.xlsx\n";
}

/**
 * @brief 测试数字格式化功能
 */
TEST_F(NewFeaturesExampleTest, NumberFormatFeatures) {
    std::cout << "\n=== 数字格式化功能演示 ===\n";
    
    // 创建工作簿和工作表
    TinaXlsx::Workbook workbook;
    auto* sheet = workbook.addSheet("格式化示例");
    ASSERT_NE(sheet, nullptr);
    
    // 设置标题
    sheet->setCellValue(TinaXlsx::row_t(1), TinaXlsx::column_t(1), "格式类型");
    sheet->setCellValue(TinaXlsx::row_t(1), TinaXlsx::column_t(2), "原始值");
    sheet->setCellValue(TinaXlsx::row_t(1), TinaXlsx::column_t(3), "格式化后");
    
    double testValue = 1234.5678;
    
    // 数字格式
    sheet->setCellValue(TinaXlsx::row_t(2), TinaXlsx::column_t(1), "数字格式");
    sheet->setCellValue(TinaXlsx::row_t(2), TinaXlsx::column_t(2), testValue);
    bool success = sheet->setCellNumberFormat(TinaXlsx::row_t(2), TinaXlsx::column_t(3), TinaXlsx::TXCell::NumberFormat::Number, 2);
    EXPECT_TRUE(success);
    sheet->setCellValue(TinaXlsx::row_t(2), TinaXlsx::column_t(3), testValue);
    
    // 货币格式
    sheet->setCellValue(TinaXlsx::row_t(3), TinaXlsx::column_t(1), "货币格式");
    sheet->setCellValue(TinaXlsx::row_t(3), TinaXlsx::column_t(2), testValue);
    success = sheet->setCellNumberFormat(TinaXlsx::row_t(3), TinaXlsx::column_t(3), TinaXlsx::TXCell::NumberFormat::Currency, 2);
    EXPECT_TRUE(success);
    sheet->setCellValue(TinaXlsx::row_t(3), TinaXlsx::column_t(3), testValue);
    
    // 百分比格式
    double percentValue = 0.1234;
    sheet->setCellValue(TinaXlsx::row_t(4), TinaXlsx::column_t(1), "百分比格式");
    sheet->setCellValue(TinaXlsx::row_t(4), TinaXlsx::column_t(2), percentValue);
    success = sheet->setCellNumberFormat(TinaXlsx::row_t(4), TinaXlsx::column_t(3), TinaXlsx::TXCell::NumberFormat::Percentage, 2);
    EXPECT_TRUE(success);
    sheet->setCellValue(TinaXlsx::row_t(4), TinaXlsx::column_t(3), percentValue);
    
    // 科学计数法
    double largeValue = 1234567890.0;
    sheet->setCellValue(TinaXlsx::row_t(5), TinaXlsx::column_t(1), "科学计数法");
    sheet->setCellValue(TinaXlsx::row_t(5), TinaXlsx::column_t(2), largeValue);
    success = sheet->setCellNumberFormat(TinaXlsx::row_t(5), TinaXlsx::column_t(3), TinaXlsx::TXCell::NumberFormat::Scientific, 2);
    EXPECT_TRUE(success);
    sheet->setCellValue(TinaXlsx::row_t(5), TinaXlsx::column_t(3), largeValue);
    
    // 日期格式
    double dateValue = TinaXlsx::TXNumberFormat::getCurrentExcelDate();
    sheet->setCellValue(TinaXlsx::row_t(6), TinaXlsx::column_t(1), "日期格式");
    sheet->setCellValue(TinaXlsx::row_t(6), TinaXlsx::column_t(2), dateValue);
    success = sheet->setCellNumberFormat(TinaXlsx::row_t(6), TinaXlsx::column_t(3), TinaXlsx::TXCell::NumberFormat::Date);
    EXPECT_TRUE(success);
    sheet->setCellValue(TinaXlsx::row_t(6), TinaXlsx::column_t(3), dateValue);
    
    // 自定义格式
    sheet->setCellValue(TinaXlsx::row_t(7), TinaXlsx::column_t(1), "自定义格式");
    sheet->setCellValue(TinaXlsx::row_t(7), TinaXlsx::column_t(2), testValue);
    success = sheet->setCellCustomFormat(TinaXlsx::row_t(7), TinaXlsx::column_t(3), "#,##0.00 \"元\"");
    EXPECT_TRUE(success);
    sheet->setCellValue(TinaXlsx::row_t(7), TinaXlsx::column_t(3), testValue);
    
    // 获取格式化后的值
    std::string formattedValue = sheet->getCellFormattedValue(TinaXlsx::row_t(2), TinaXlsx::column_t(3));
    std::cout << "数字格式化结果: " << formattedValue << std::endl;
    
    formattedValue = sheet->getCellFormattedValue(TinaXlsx::row_t(3), TinaXlsx::column_t(3));
    std::cout << "货币格式化结果: " << formattedValue << std::endl;
    
    formattedValue = sheet->getCellFormattedValue(TinaXlsx::row_t(4), TinaXlsx::column_t(3));
    std::cout << "百分比格式化结果: " << formattedValue << std::endl;
    
    // 批量设置格式
    std::vector<std::pair<TinaXlsx::TXCoordinate, TinaXlsx::TXCell::NumberFormat>> formats;
    formats.emplace_back(TinaXlsx::TXCoordinate(TinaXlsx::row_t(8), TinaXlsx::column_t(1)), TinaXlsx::TXCell::NumberFormat::Number);
    formats.emplace_back(TinaXlsx::TXCoordinate(TinaXlsx::row_t(8), TinaXlsx::column_t(2)), TinaXlsx::TXCell::NumberFormat::Currency);
    formats.emplace_back(TinaXlsx::TXCoordinate(TinaXlsx::row_t(8), TinaXlsx::column_t(3)), TinaXlsx::TXCell::NumberFormat::Percentage);
    
    std::size_t setCount = sheet->setCellFormats(formats);
    std::cout << "批量设置格式数量: " << setCount << std::endl;
    
    // 保存文件
    success = workbook.saveToFile("number_format_example.xlsx");
    EXPECT_TRUE(success);
    
    std::cout << "数字格式化示例文件已保存: number_format_example.xlsx\n";
}

/**
 * @brief 综合功能演示
 */
TEST_F(NewFeaturesExampleTest, ComprehensiveExample) {
    std::cout << "\n=== 综合功能演示 ===\n";
    
    auto start = std::chrono::high_resolution_clock::now();
    
    // 创建工作簿和工作表
    TinaXlsx::Workbook workbook;
    auto* sheet = workbook.addSheet("综合示例");
    ASSERT_NE(sheet, nullptr);
    
    // 创建一个销售报表，综合使用所有新功能
    
    // 1. 标题行（合并单元格）
    sheet->setCellValue(TinaXlsx::row_t(1), TinaXlsx::column_t(1), "2024年度销售业绩报表");
    sheet->mergeCells(TinaXlsx::row_t(1), TinaXlsx::column_t(1), TinaXlsx::row_t(1), TinaXlsx::column_t(6));
    
    // 2. 表头
    sheet->setCellValue(TinaXlsx::row_t(3), TinaXlsx::column_t(1), "产品");
    sheet->setCellValue(TinaXlsx::row_t(3), TinaXlsx::column_t(2), "单价");
    sheet->setCellValue(TinaXlsx::row_t(3), TinaXlsx::column_t(3), "数量");
    sheet->setCellValue(TinaXlsx::row_t(3), TinaXlsx::column_t(4), "金额");
    sheet->setCellValue(TinaXlsx::row_t(3), TinaXlsx::column_t(5), "税率");
    sheet->setCellValue(TinaXlsx::row_t(3), TinaXlsx::column_t(6), "含税金额");
    
    // 3. 数据行
    std::vector<std::string> products = {"笔记本电脑", "台式机", "显示器", "键盘", "鼠标"};
    std::vector<double> prices = {5999.99, 3999.99, 1299.99, 299.99, 99.99};
    std::vector<int> quantities = {50, 30, 100, 200, 300};
    
    for (size_t i = 0; i < products.size(); ++i) {
        int row = 4 + static_cast<int>(i);
        
        // 产品名称
        sheet->setCellValue(TinaXlsx::row_t(row), TinaXlsx::column_t(1), products[i]);
        
        // 单价（货币格式）
        sheet->setCellValue(TinaXlsx::row_t(row), TinaXlsx::column_t(2), prices[i]);
        sheet->setCellNumberFormat(TinaXlsx::row_t(row), TinaXlsx::column_t(2), TinaXlsx::TXCell::NumberFormat::Currency, 2);
        
        // 数量（数字格式）
        sheet->setCellValue(TinaXlsx::row_t(row), TinaXlsx::column_t(3), static_cast<double>(quantities[i]));
        sheet->setCellNumberFormat(TinaXlsx::row_t(row), TinaXlsx::column_t(3), TinaXlsx::TXCell::NumberFormat::Number, 0);
        
        // 金额（公式：单价*数量）
        std::string formula = "B" + std::to_string(row) + "*C" + std::to_string(row);
        sheet->setCellFormula(TinaXlsx::row_t(row), TinaXlsx::column_t(4), formula);
        sheet->setCellNumberFormat(TinaXlsx::row_t(row), TinaXlsx::column_t(4), TinaXlsx::TXCell::NumberFormat::Currency, 2);
        
        // 税率（百分比格式）
        sheet->setCellValue(TinaXlsx::row_t(row), TinaXlsx::column_t(5), 0.13);  // 13%税率
        sheet->setCellNumberFormat(TinaXlsx::row_t(row), TinaXlsx::column_t(5), TinaXlsx::TXCell::NumberFormat::Percentage, 1);
        
        // 含税金额（公式：金额*(1+税率)）
        formula = "D" + std::to_string(row) + "*(1+E" + std::to_string(row) + ")";
        sheet->setCellFormula(TinaXlsx::row_t(row), TinaXlsx::column_t(6), formula);
        sheet->setCellNumberFormat(TinaXlsx::row_t(row), TinaXlsx::column_t(6), TinaXlsx::TXCell::NumberFormat::Currency, 2);
    }
    
    // 4. 合计行
    int totalRow = 4 + static_cast<int>(products.size());
    sheet->setCellValue(TinaXlsx::row_t(totalRow), TinaXlsx::column_t(1), "合计");
    sheet->mergeCells(TinaXlsx::row_t(totalRow), TinaXlsx::column_t(1), TinaXlsx::row_t(totalRow), TinaXlsx::column_t(3));  // 合并"合计"跨3列
    
    // 金额合计（公式：SUM）
    std::string sumFormula = "SUM(D4:D" + std::to_string(totalRow - 1) + ")";
    sheet->setCellFormula(TinaXlsx::row_t(totalRow), TinaXlsx::column_t(4), sumFormula);
    sheet->setCellNumberFormat(TinaXlsx::row_t(totalRow), TinaXlsx::column_t(4), TinaXlsx::TXCell::NumberFormat::Currency, 2);
    
    // 含税金额合计
    sumFormula = "SUM(F4:F" + std::to_string(totalRow - 1) + ")";
    sheet->setCellFormula(TinaXlsx::row_t(totalRow), TinaXlsx::column_t(6), sumFormula);
    sheet->setCellNumberFormat(TinaXlsx::row_t(totalRow), TinaXlsx::column_t(6), TinaXlsx::TXCell::NumberFormat::Currency, 2);
    
    // 5. 计算所有公式
    std::size_t calculatedCount = sheet->calculateAllFormulas();
    std::cout << "计算的公式数量: " << calculatedCount << std::endl;
    
    // 6. 验证合并单元格数量
    std::size_t mergeCount = sheet->getMergeCount();
    std::cout << "合并区域数量: " << mergeCount << std::endl;
    
    // 7. 添加日期信息
    double currentDate = TinaXlsx::TXNumberFormat::getCurrentExcelDate();
    sheet->setCellValue(TinaXlsx::row_t(totalRow + 2), TinaXlsx::column_t(1), "报表生成日期:");
    sheet->setCellValue(TinaXlsx::row_t(totalRow + 2), TinaXlsx::column_t(2), currentDate);
    sheet->setCellNumberFormat(TinaXlsx::row_t(totalRow + 2), TinaXlsx::column_t(2), TinaXlsx::TXCell::NumberFormat::Date);
    
    // 保存文件
    bool success = workbook.saveToFile("comprehensive_example.xlsx");
    EXPECT_TRUE(success);
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    std::cout << "综合示例文件已保存: comprehensive_example.xlsx\n";
    std::cout << "处理时间: " << duration.count() << "ms\n";
    
    // 验证性能 - 应该在合理时间内完成
    EXPECT_LT(duration.count(), 1000);  // 应该在1秒内完成
}

/**
 * @brief 性能测试
 */
TEST_F(NewFeaturesExampleTest, PerformanceTest) {
    std::cout << "\n=== 性能测试 ===\n";
    
    auto start = std::chrono::high_resolution_clock::now();
    
    TinaXlsx::Workbook workbook;
    auto* sheet = workbook.addSheet("性能测试");
    ASSERT_NE(sheet, nullptr);
    
    const int ROWS = 1000;
    const int COLS = 10;
    
    std::cout << "开始创建 " << ROWS << "x" << COLS << " 的数据表...\n";
    
    // 批量设置数据
    std::vector<std::pair<TinaXlsx::TXCoordinate, TinaXlsx::TXSheet::CellValue>> values;
    values.reserve(ROWS * COLS);
    
    for (int row = 1; row <= ROWS; ++row) {
        for (int col = 1; col <= COLS; ++col) {
            double value = row * col * 1.5;
            values.emplace_back(TinaXlsx::TXCoordinate(TinaXlsx::row_t(row), TinaXlsx::column_t(col)), value);
        }
    }
    
    auto batchStart = std::chrono::high_resolution_clock::now();
    std::size_t setCount = sheet->setCellValues(values);
    auto batchEnd = std::chrono::high_resolution_clock::now();
    
    EXPECT_EQ(setCount, ROWS * COLS);
    
    auto batchDuration = std::chrono::duration_cast<std::chrono::milliseconds>(batchEnd - batchStart);
    std::cout << "批量设置 " << setCount << " 个单元格耗时: " << batchDuration.count() << "ms\n";
    
    // 批量设置格式
    std::vector<std::pair<TinaXlsx::TXCoordinate, TinaXlsx::TXCell::NumberFormat>> formats;
    formats.reserve(COLS);
    
    for (int col = 1; col <= COLS; ++col) {
        TinaXlsx::TXCell::NumberFormat fmt = (col % 2 == 0) ? 
            TinaXlsx::TXCell::NumberFormat::Currency : TinaXlsx::TXCell::NumberFormat::Number;
        formats.emplace_back(TinaXlsx::TXCoordinate(TinaXlsx::row_t(1), TinaXlsx::column_t(col)), fmt);
    }
    
    auto formatStart = std::chrono::high_resolution_clock::now();
    std::size_t formatCount = sheet->setCellFormats(formats);
    auto formatEnd = std::chrono::high_resolution_clock::now();
    
    auto formatDuration = std::chrono::duration_cast<std::chrono::milliseconds>(formatEnd - formatStart);
    std::cout << "批量设置 " << formatCount << " 个格式耗时: " << formatDuration.count() << "ms\n";
    
    // 批量合并单元格
    std::vector<TinaXlsx::TXMergedCells::MergeRegion> mergeRegions;
    for (int i = 0; i < 10; ++i) {
        int startRow = i * 100 + 1;
        int endRow = startRow + 1;
        mergeRegions.emplace_back(TinaXlsx::row_t(startRow), TinaXlsx::column_t(1), TinaXlsx::row_t(endRow), TinaXlsx::column_t(2));
    }
    
    TinaXlsx::TXMergedCells mergedCells;
    auto mergeStart = std::chrono::high_resolution_clock::now();
    std::size_t mergeCount = mergedCells.batchMergeCells(mergeRegions);
    auto mergeEnd = std::chrono::high_resolution_clock::now();
    
    auto mergeDuration = std::chrono::duration_cast<std::chrono::milliseconds>(mergeEnd - mergeStart);
    std::cout << "批量合并 " << mergeCount << " 个区域耗时: " << mergeDuration.count() << "ms\n";
    
    // 保存大文件
    auto saveStart = std::chrono::high_resolution_clock::now();
    bool success = workbook.saveToFile("performance_test.xlsx");
    auto saveEnd = std::chrono::high_resolution_clock::now();
    
    EXPECT_TRUE(success);
    
    auto saveDuration = std::chrono::duration_cast<std::chrono::milliseconds>(saveEnd - saveStart);
    std::cout << "保存大文件耗时: " << saveDuration.count() << "ms\n";
    
    auto end = std::chrono::high_resolution_clock::now();
    auto totalDuration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    std::cout << "性能测试总耗时: " << totalDuration.count() << "ms\n";
    std::cout << "平均每个单元格处理时间: " << static_cast<double>(totalDuration.count()) / (ROWS * COLS) << "ms\n";
    
    // 性能要求：处理1万个单元格应该在10秒内完成
    EXPECT_LT(totalDuration.count(), 10000);
}

} // namespace 