// Created by wuxianggujun on 2025/5/29.
// 简化的完整XLSX文件创建和功能测试

#include <gtest/gtest.h>
#include <cstdio>
#include "TinaXlsx/TinaXlsx.hpp"

using namespace TinaXlsx;

class CompleteXlsxSimpleTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 每个测试前清理可能存在的文件
        std::remove("test_simple.xlsx");
        std::remove("test_formulas.xlsx");
        std::remove("test_mergedcells.xlsx");
    }

    void TearDown() override {
        // 测试后可选择保留或清理文件
    }
};

// 测试创建简单的XLSX文件
TEST_F(CompleteXlsxSimpleTest, CreateSimpleXlsxFile) {
    // 创建工作簿
    TXWorkbook workbook;
    
    // 创建工作表
    auto sheet = std::make_unique<TXSheet>("简单测试", &workbook);
    
    // 使用getCell获取单元格并设置值
    auto cellA1 = sheet->getCell("A1");
    if (cellA1) cellA1->setValue(std::string("姓名"));
    
    auto cellB1 = sheet->getCell("B1");
    if (cellB1) cellB1->setValue(std::string("年龄"));
    
    auto cellC1 = sheet->getCell("C1");
    if (cellC1) cellC1->setValue(std::string("薪资"));
    
    auto cellA2 = sheet->getCell("A2");
    if (cellA2) cellA2->setValue(std::string("张三"));
    
    auto cellB2 = sheet->getCell("B2");
    if (cellB2) cellB2->setValue(static_cast<int64_t>(25));
    
    auto cellC2 = sheet->getCell("C2");
    if (cellC2) cellC2->setValue(5000.50);
    
    auto cellA3 = sheet->getCell("A3");
    if (cellA3) cellA3->setValue(std::string("李四"));
    
    auto cellB3 = sheet->getCell("B3");
    if (cellB3) cellB3->setValue(static_cast<int64_t>(30));
    
    auto cellC3 = sheet->getCell("C3");
    if (cellC3) cellC3->setValue(6500.75);
    
    // 将工作表添加到工作簿
    workbook.addSheet(std::move(sheet));
    
    // 保存文件
    bool saveResult = workbook.saveToFile("test_simple.xlsx");
    EXPECT_TRUE(saveResult) << "Failed to save XLSX file";
    
    // 验证文件是否存在
    FILE* file = fopen("test_simple.xlsx", "rb");
    EXPECT_NE(file, nullptr) << "Generated XLSX file does not exist";
    if (file) {
        fclose(file);
    }
    
    // 验证工作表数量
    EXPECT_EQ(workbook.getSheetCount(), 1) << "Workbook should contain 1 sheet";
}

// 测试公式和布尔值
TEST_F(CompleteXlsxSimpleTest, CreateXlsxWithFormulas) {
    TXWorkbook workbook;
    auto sheet = std::make_unique<TXSheet>("公式测试", &workbook);
    
    // 添加数据
    auto cellA1 = sheet->getCell("A1");
    if (cellA1) cellA1->setValue(std::string("数值1"));
    
    auto cellB1 = sheet->getCell("B1");
    if (cellB1) cellB1->setValue(std::string("数值2"));
    
    auto cellC1 = sheet->getCell("C1");
    if (cellC1) cellC1->setValue(std::string("总和"));
    
    auto cellA2 = sheet->getCell("A2");
    if (cellA2) cellA2->setValue(100.0);
    
    auto cellB2 = sheet->getCell("B2");
    if (cellB2) cellB2->setValue(200.0);
    
    // 设置公式
    auto cellC2 = sheet->getCell("C2");
    if (cellC2) cellC2->setFormula("A2+B2");
    
    auto cellA3 = sheet->getCell("A3");
    if (cellA3) cellA3->setValue(150.0);
    
    auto cellB3 = sheet->getCell("B3");
    if (cellB3) cellB3->setValue(250.0);
    
    auto cellC3 = sheet->getCell("C3");
    if (cellC3) cellC3->setFormula("A3+B3");
    
    // 添加布尔值测试
    auto cellD1 = sheet->getCell("D1");
    if (cellD1) cellD1->setValue(std::string("状态"));
    
    auto cellD2 = sheet->getCell("D2");
    if (cellD2) cellD2->setValue(true);
    
    auto cellD3 = sheet->getCell("D3");
    if (cellD3) cellD3->setValue(false);
    
    workbook.addSheet(std::move(sheet));
    
    bool saveResult = workbook.saveToFile("test_formulas.xlsx");
    EXPECT_TRUE(saveResult) << "Failed to save XLSX file with formulas";    
    FILE* file = fopen("test_formulas.xlsx", "rb");
    EXPECT_NE(file, nullptr) << "Generated XLSX file with formulas does not exist";
    if (file) {
        fclose(file);
    }
}

// 测试合并单元格
TEST_F(CompleteXlsxSimpleTest, CreateXlsxWithMergedCells) {
    TXWorkbook workbook;
    auto sheet = std::make_unique<TXSheet>("合并单元格", &workbook);
    
    // 设置标题
    auto cellA1 = sheet->getCell("A1");
    if (cellA1) cellA1->setValue(std::string("季度销售报告"));
    
    // 合并标题单元格
    sheet->mergeCells("A1:D1");
    
    // 设置表头
    auto cellA3 = sheet->getCell("A3");
    if (cellA3) cellA3->setValue(std::string("季度"));
    
    auto cellB3 = sheet->getCell("B3");
    if (cellB3) cellB3->setValue(std::string("Q1"));
    
    auto cellC3 = sheet->getCell("C3");
    if (cellC3) cellC3->setValue(std::string("Q2"));
    
    auto cellD3 = sheet->getCell("D3");
    if (cellD3) cellD3->setValue(std::string("Q3"));
    
    // 设置数据
    auto cellA4 = sheet->getCell("A4");
    if (cellA4) cellA4->setValue(std::string("销售额"));
    
    auto cellB4 = sheet->getCell("B4");
    if (cellB4) cellB4->setValue(100000.0);
    
    auto cellC4 = sheet->getCell("C4");
    if (cellC4) cellC4->setValue(120000.0);
    
    auto cellD4 = sheet->getCell("D4");
    if (cellD4) cellD4->setValue(110000.0);
    
    workbook.addSheet(std::move(sheet));
    
    bool saveResult = workbook.saveToFile("test_mergedcells.xlsx");
    EXPECT_TRUE(saveResult) << "Failed to save XLSX file with merged cells";
    
    FILE* file = fopen("test_mergedcells.xlsx", "rb");
    EXPECT_NE(file, nullptr) << "Generated XLSX file with merged cells does not exist";
    if (file) {
        fclose(file);
    }
}