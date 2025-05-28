//
// Formulas Test
// 测试公式功能
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include <filesystem>
#include <iostream>

class FormulasTest : public ::testing::Test {
protected:
    void SetUp() override {
        std::filesystem::create_directories("output");
    }

    void TearDown() override {
        std::filesystem::remove("output/formulas_test.xlsx");
    }
};

/**
 * @brief 测试基本公式设置和获取
 */
TEST_F(FormulasTest, BasicFormulaOperations) {
    std::cout << "\n=== 基本公式操作测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("公式测试");
    ASSERT_NE(sheet, nullptr);
    
    // 设置一些基础数据
    sheet->setCellValue("A1", 10.0);
    sheet->setCellValue("A2", 20.0);
    sheet->setCellValue("A3", 30.0);
    
    // 设置求和公式
    bool formulaSet = sheet->setCellFormula(TinaXlsx::row_t(4), TinaXlsx::column_t(1), "=SUM(A1:A3)");
    EXPECT_TRUE(formulaSet);
    
    // 验证公式是否正确设置
    std::string formula = sheet->getCellFormula(TinaXlsx::row_t(4), TinaXlsx::column_t(1));
    EXPECT_EQ("=SUM(A1:A3)", formula);
    
    // 设置更多公式
    EXPECT_TRUE(sheet->setCellFormula(TinaXlsx::row_t(5), TinaXlsx::column_t(1), "=A1+A2"));
    EXPECT_TRUE(sheet->setCellFormula(TinaXlsx::row_t(6), TinaXlsx::column_t(1), "=A1*A2"));
    EXPECT_TRUE(sheet->setCellFormula(TinaXlsx::row_t(7), TinaXlsx::column_t(1), "=A2/A1"));
    
    // 验证公式
    EXPECT_EQ("=A1+A2", sheet->getCellFormula(TinaXlsx::row_t(5), TinaXlsx::column_t(1)));
    EXPECT_EQ("=A1*A2", sheet->getCellFormula(TinaXlsx::row_t(6), TinaXlsx::column_t(1)));
    EXPECT_EQ("=A2/A1", sheet->getCellFormula(TinaXlsx::row_t(7), TinaXlsx::column_t(1)));
    
    // 保存文件
    bool saved = workbook.saveToFile("output/formulas_test.xlsx");
    EXPECT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    
    std::cout << "基本公式操作测试通过！\n";
}/**
 * @brief 测试批量公式设置
 */
TEST_F(FormulasTest, BatchFormulaOperations) {
    std::cout << "\n=== 批量公式操作测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("批量公式");
    ASSERT_NE(sheet, nullptr);
    
    // 准备批量公式数据
    std::vector<std::pair<TinaXlsx::TXCoordinate, std::string>> formulas;
    
    // 设置一些基础数据
    for (int i = 1; i <= 10; ++i) {
        sheet->setCellValue(TinaXlsx::row_t(i), TinaXlsx::column_t(1), static_cast<double>(i));
        sheet->setCellValue(TinaXlsx::row_t(i), TinaXlsx::column_t(2), static_cast<double>(i * 2));
        
        // 创建求和公式
        std::string formula = "=A" + std::to_string(i) + "+B" + std::to_string(i);
        formulas.emplace_back(
            TinaXlsx::TXCoordinate(TinaXlsx::row_t(i), TinaXlsx::column_t(3)),
            formula
        );
    }
    
    // 批量设置公式
    size_t count = sheet->setCellFormulas(formulas);
    EXPECT_EQ(10, count);
    
    // 验证批量设置的公式
    for (int i = 1; i <= 10; ++i) {
        std::string expected = "=A" + std::to_string(i) + "+B" + std::to_string(i);
        std::string actual = sheet->getCellFormula(TinaXlsx::row_t(i), TinaXlsx::column_t(3));
        EXPECT_EQ(expected, actual);
    }
    
    bool saved = workbook.saveToFile("output/formulas_test.xlsx");
    EXPECT_TRUE(saved);
    
    std::cout << "批量公式操作测试通过！\n";
}

/**
 * @brief 测试复杂公式
 */
TEST_F(FormulasTest, ComplexFormulas) {
    std::cout << "\n=== 复杂公式测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("复杂公式");
    ASSERT_NE(sheet, nullptr);
    
    // 设置复杂公式
    EXPECT_TRUE(sheet->setCellFormula(TinaXlsx::row_t(1), TinaXlsx::column_t(1), "=SUM(A1:A10)"));
    EXPECT_TRUE(sheet->setCellFormula(TinaXlsx::row_t(2), TinaXlsx::column_t(1), "=AVERAGE(B1:B10)"));
    EXPECT_TRUE(sheet->setCellFormula(TinaXlsx::row_t(3), TinaXlsx::column_t(1), "=MAX(C1:C10)"));
    EXPECT_TRUE(sheet->setCellFormula(TinaXlsx::row_t(4), TinaXlsx::column_t(1), "=MIN(D1:D10)"));
    EXPECT_TRUE(sheet->setCellFormula(TinaXlsx::row_t(5), TinaXlsx::column_t(1), "=COUNT(E1:E10)"));
    
    // 设置条件公式
    EXPECT_TRUE(sheet->setCellFormula(TinaXlsx::row_t(6), TinaXlsx::column_t(1), "=IF(A1>10,\"大\",\"小\")"));
    
    // 设置嵌套公式
    EXPECT_TRUE(sheet->setCellFormula(TinaXlsx::row_t(7), TinaXlsx::column_t(1), "=SUM(A1:A5)+AVERAGE(B1:B5)"));
    
    bool saved = workbook.saveToFile("output/formulas_test.xlsx");
    EXPECT_TRUE(saved);
    
    std::cout << "复杂公式测试通过！\n";
}