#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"

using namespace TinaXlsx;

class NewFeaturesTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("TestSheet");
    }

    void TearDown() override {
        workbook.reset();
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
};

// ==================== 公式功能测试 ====================

TEST_F(NewFeaturesTest, FormulaBasicOperations) {
    // 测试基本公式创建
    auto formula = std::make_unique<TXFormula>();
    EXPECT_TRUE(formula->parseFormula("SUM(A1:A3)"));
    EXPECT_EQ(formula->getFormulaString(), "SUM(A1:A3)");
    
    // 测试公式错误检测
    EXPECT_EQ(formula->getLastError(), TXFormula::FormulaError::None);
}

TEST_F(NewFeaturesTest, FormulaEvaluation) {
    // 设置测试数据
    sheet->setCellValue(row_t(1), column_t(1), 10.0);  // A1 = 10
    sheet->setCellValue(row_t(2), column_t(1), 20.0);  // A2 = 20
    sheet->setCellValue(row_t(3), column_t(1), 30.0);  // A3 = 30
    
    // 创建SUM公式
    auto formula = std::make_unique<TXFormula>();
    EXPECT_TRUE(formula->parseFormula("SUM(A1:A3)"));
    auto result = formula->evaluate(sheet, row_t(4), column_t(1));  // 在A4位置计算
    
    EXPECT_TRUE(std::holds_alternative<double>(result));
    EXPECT_DOUBLE_EQ(std::get<double>(result), 60.0);
}

TEST_F(NewFeaturesTest, CellFormulaIntegration) {
    // 设置测试数据
    sheet->setCellValue(row_t(1), column_t(1), 5.0);
    sheet->setCellValue(row_t(2), column_t(1), 15.0);
    
    // 在单元格中设置公式
    TXCell* cell = sheet->getCell(row_t(3), column_t(1));
    ASSERT_NE(cell, nullptr);
    
    cell->setFormula("SUM(A1:A2)");

    EXPECT_TRUE(cell->isFormula());
    EXPECT_EQ(cell->getFormula(), "SUM(A1:A2)");
    
    // 计算公式
    auto result = cell->evaluateFormula(sheet, 3, 1);
    EXPECT_TRUE(std::holds_alternative<double>(result));
    EXPECT_DOUBLE_EQ(std::get<double>(result), 20.0);
}

// ==================== 合并单元格功能测试 ====================

TEST_F(NewFeaturesTest, MergedCellsBasicOperations) {
    TXMergedCells mergedCells;
    
    // 测试添加合并区域
    TXRange range(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(2), column_t(2)));  // A1:B2
    EXPECT_TRUE(mergedCells.mergeCells(range));
    EXPECT_EQ(mergedCells.getMergeCount(), 1);
    
    // 测试查找合并区域
    const auto* foundRegion = mergedCells.getMergeRegion(row_t(1), column_t(1));
    EXPECT_NE(foundRegion, nullptr);
    EXPECT_EQ(foundRegion->startRow, row_t(1));
    EXPECT_EQ(foundRegion->startCol, column_t(1));
    EXPECT_EQ(foundRegion->endRow, row_t(2));
    EXPECT_EQ(foundRegion->endCol, column_t(2));
}

TEST_F(NewFeaturesTest, MergedCellsOverlapDetection) {
    TXMergedCells mergedCells;
    
    // 添加第一个合并区域
    TXRange range1(TXCoordinate(row_t(1), column_t(1)), TXCoordinate(row_t(2), column_t(2)));  // A1:B2
    EXPECT_TRUE(mergedCells.mergeCells(range1));
    
    // 尝试添加重叠的区域
    TXRange range2(TXCoordinate(row_t(2), column_t(2)), TXCoordinate(row_t(3), column_t(3)));  // B2:C3
    auto overlapping = mergedCells.getOverlappingRegions(range2);
    EXPECT_FALSE(overlapping.empty());
    EXPECT_FALSE(mergedCells.mergeCells(range2));
}

TEST_F(NewFeaturesTest, SheetMergedCellsIntegration) {
    // 测试工作表级别的合并单元格功能
    EXPECT_TRUE(sheet->mergeCells(row_t(1), column_t(1), row_t(2), column_t(2)));  // 合并A1:B2
    EXPECT_EQ(sheet->getMergeCount(), 1);
    
    // 检查单元格是否被标记为合并
    EXPECT_TRUE(sheet->isCellMerged(row_t(1), column_t(1)));
    EXPECT_TRUE(sheet->isCellMerged(row_t(2), column_t(2)));
    EXPECT_FALSE(sheet->isCellMerged(row_t(3), column_t(3)));
    
    // 获取合并区域
    auto region = sheet->getMergeRegion(row_t(1), column_t(1));
    EXPECT_TRUE(region.isValid());
    EXPECT_EQ(region.getRowCount(), row_t(2));
    EXPECT_EQ(region.getColCount(), column_t(2));
    
    // 取消合并
    EXPECT_TRUE(sheet->unmergeCells(row_t(1), column_t(1)));
    EXPECT_EQ(sheet->getMergeCount(), 0);
    EXPECT_FALSE(sheet->isCellMerged(row_t(1), column_t(1)));
}

// ==================== 数字格式化功能测试 ====================

TEST_F(NewFeaturesTest, NumberFormatBasicOperations) {
    TXNumberFormat formatter;
    
    // 测试数字格式化
    TXNumberFormat::FormatOptions options;
    options.decimalPlaces = 2;
    options.useThousandSeparator = true;
    
    formatter.setFormat(TXNumberFormat::FormatType::Number, options);
    
    std::string result = formatter.formatNumber(1234.567);
    // 结果应该是类似 "1,234.57" 的格式
    EXPECT_TRUE(result.find("1,234") != std::string::npos);
    EXPECT_TRUE(result.find(".57") != std::string::npos);
}

TEST_F(NewFeaturesTest, NumberFormatPercentage) {
    TXNumberFormat formatter;
    
    TXNumberFormat::FormatOptions options;
    options.decimalPlaces = 1;
    
    formatter.setFormat(TXNumberFormat::FormatType::Percentage, options);
    
    std::string result = formatter.formatPercentage(0.1234);
    // 结果应该是类似 "12.3%" 的格式
    EXPECT_TRUE(result.find("12.3") != std::string::npos);
    EXPECT_TRUE(result.find("%") != std::string::npos);
}

TEST_F(NewFeaturesTest, NumberFormatCurrency) {
    TXNumberFormat formatter;
    
    TXNumberFormat::FormatOptions options;
    options.decimalPlaces = 2;
    options.currencySymbol = "$";
    
    formatter.setFormat(TXNumberFormat::FormatType::Currency, options);
    
    std::string result = formatter.formatCurrency(1234.56);
    // 结果应该是类似 "$1,234.56" 的格式
    EXPECT_TRUE(result.find("$") != std::string::npos);
    EXPECT_TRUE(result.find("1,234.56") != std::string::npos);
}

TEST_F(NewFeaturesTest, CellNumberFormatIntegration) {
    // 测试单元格级别的数字格式化
    TXCell* cell = sheet->getCell(row_t(1), column_t(1));
    ASSERT_NE(cell, nullptr);
    
    cell->setNumberValue(1234.567);
    
    // 设置预定义格式
    cell->setPredefinedFormat(TXCell::NumberFormat::Number, 2, true);
    
    std::string formatted = cell->getFormattedValue();
    EXPECT_FALSE(formatted.empty());
    
    // 测试自定义格式
    cell->setCustomFormat("#,##0.00");
    formatted = cell->getFormattedValue();
    EXPECT_FALSE(formatted.empty());
}

// ==================== 综合功能测试 ====================

TEST_F(NewFeaturesTest, IntegratedFeatureTest) {
    // 创建一个包含所有新功能的综合测试
    
    // 1. 设置数据和格式
    sheet->setCellValue(row_t(1), column_t(1), 100.0);
    sheet->setCellValue(row_t(1), column_t(2), 200.0);
    sheet->setCellValue(row_t(1), column_t(3), 300.0);
    
    // 2. 设置数字格式
    sheet->setCellNumberFormat(row_t(1), column_t(1), TXCell::NumberFormat::Currency, 2);
    sheet->setCellNumberFormat(row_t(1), column_t(2), TXCell::NumberFormat::Currency, 2);
    sheet->setCellNumberFormat(row_t(1), column_t(3), TXCell::NumberFormat::Currency, 2);
    
    // 3. 创建求和公式
    sheet->setCellFormula(row_t(2), column_t(1), "SUM(A1:C1)");
    
    // 4. 合并结果单元格
    sheet->mergeCells(row_t(3), column_t(1), row_t(3), column_t(3));  // 合并A3:C3
    
    // 5. 验证结果
    EXPECT_EQ(sheet->getMergeCount(), 1);
    EXPECT_TRUE(sheet->isCellMerged(row_t(3), column_t(2)));
    
    std::string formula = sheet->getCellFormula(row_t(2), column_t(1));
    EXPECT_EQ(formula, "SUM(A1:C1)");
    
    // 6. 计算公式
    std::size_t calculatedCount = sheet->calculateAllFormulas();
    EXPECT_GT(calculatedCount, 0);
} 
