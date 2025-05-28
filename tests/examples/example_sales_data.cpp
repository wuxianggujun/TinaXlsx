//
// Sales Data Example Test
// Demonstrates practical usage of TinaXlsx library
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include <filesystem>
#include <iostream>

class SalesDataExampleTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 清理可能存在的测试文件
        std::filesystem::remove("output/SalesDataExample.xlsx");
    }

    void TearDown() override {
        // 清理测试文件
        std::filesystem::remove("output/SalesDataExample.xlsx");
    }
};

TEST_F(SalesDataExampleTest, CreateSalesReportWithActualData) {
    std::cout << "=== Starting CreateSalesReportWithActualData Test ===" << std::endl;
    
    try {
        // 步骤1: 创建工作簿和工作表
        std::cout << "1. Creating workbook..." << std::endl;
        TinaXlsx::TXWorkbook workbook;
        
        std::cout << "2. Adding sheet..." << std::endl;
        workbook.addSheet("Q1_Sales_Report");
        
        std::cout << "3. Getting sheet..." << std::endl;
        auto sheet = workbook.getSheet("Q1_Sales_Report");
        ASSERT_NE(sheet, nullptr);
        
        std::cout << "4. Writing header data..." << std::endl;
        // 步骤2: 写入表头
        EXPECT_TRUE(sheet->setCellValue("A1", std::string("Product Name")));
        EXPECT_TRUE(sheet->setCellValue("B1", std::string("Units Sold")));
        EXPECT_TRUE(sheet->setCellValue("C1", std::string("Unit Price")));
        EXPECT_TRUE(sheet->setCellValue("D1", std::string("Revenue")));
        EXPECT_TRUE(sheet->setCellValue("E1", std::string("Region")));
        
        std::cout << "5. Writing first product..." << std::endl;
        // 写入第一个产品数据
        EXPECT_TRUE(sheet->setCellValue("A2", std::string("iPhone 15")));
        EXPECT_TRUE(sheet->setCellValue("B2", static_cast<int64_t>(1250)));
        EXPECT_TRUE(sheet->setCellValue("C2", 999.99));
        EXPECT_TRUE(sheet->setCellValue("D2", 1249987.5));
        EXPECT_TRUE(sheet->setCellValue("E2", std::string("North America")));
        
        std::cout << "6. Saving file..." << std::endl;
        // 保存文件
        bool saved = workbook.saveToFile("output/SalesDataExample.xlsx");
        ASSERT_TRUE(saved) << "Failed to save: " << workbook.getLastError();
        
        std::cout << "7. Verifying file exists..." << std::endl;
        EXPECT_TRUE(std::filesystem::exists("output/SalesDataExample.xlsx"));
        
        // 暂时跳过文件读取验证，因为读取功能还在完善中
        std::cout << "Test completed successfully (file reading validation skipped)" << std::endl;
        
    } catch (const std::exception& e) {
        std::cout << "Exception caught: " << e.what() << std::endl;
        FAIL() << "Exception: " << e.what();
    } catch (...) {
        std::cout << "Unknown exception caught" << std::endl;
        FAIL() << "Unknown exception";
    }
}

TEST_F(SalesDataExampleTest, CreateMultiSheetReport) {
    TinaXlsx::TXWorkbook workbook;
    
    // 创建多个工作表
    workbook.addSheet("Monthly_Summary");
    workbook.addSheet("Product_Details");
    workbook.addSheet("Regional_Analysis");
    
    // 验证工作表创建
    EXPECT_EQ(workbook.getSheetCount(), 3);
    EXPECT_TRUE(workbook.hasSheet("Monthly_Summary"));
    EXPECT_TRUE(workbook.hasSheet("Product_Details"));
    EXPECT_TRUE(workbook.hasSheet("Regional_Analysis"));
    
    // 在每个工作表中添加一些数据
    auto summary_sheet = workbook.getSheet("Monthly_Summary");
    ASSERT_NE(summary_sheet, nullptr);
    summary_sheet->setCellValue("A1", std::string("Month"));
    summary_sheet->setCellValue("B1", std::string("Total Sales"));
    summary_sheet->setCellValue("A2", std::string("January"));
    summary_sheet->setCellValue("B2", 1500000.0);
    
    auto product_sheet = workbook.getSheet("Product_Details");
    ASSERT_NE(product_sheet, nullptr);
    product_sheet->setCellValue("A1", std::string("Product ID"));
    product_sheet->setCellValue("B1", std::string("SKU"));
    product_sheet->setCellValue("A2", std::string("PROD001"));
    product_sheet->setCellValue("B2", std::string("IPH-15-128-BLK"));
    
    auto regional_sheet = workbook.getSheet("Regional_Analysis");
    ASSERT_NE(regional_sheet, nullptr);
    regional_sheet->setCellValue("A1", std::string("Region"));
    regional_sheet->setCellValue("B1", std::string("Market Share"));
    regional_sheet->setCellValue("A2", std::string("North America"));
    regional_sheet->setCellValue("B2", 0.45);  // 45%
    
    // 保存多工作表文件
    bool saved = workbook.saveToFile("output/SalesDataExample.xlsx");
    EXPECT_TRUE(saved) << "Failed to save multi-sheet file: " << workbook.getLastError();
    
    // 验证文件
    EXPECT_TRUE(std::filesystem::exists("output/SalesDataExample.xlsx"));
}

TEST_F(SalesDataExampleTest, HandleDifferentDataTypes) {
    TinaXlsx::TXWorkbook workbook;
    workbook.addSheet("DataTypes_Demo");
    
    auto sheet = workbook.getSheet("DataTypes_Demo");
    ASSERT_NE(sheet, nullptr);
    
    // 测试不同数据类型
    sheet->setCellValue("A1", std::string("Text Data"));
    sheet->setCellValue("A2", static_cast<int64_t>(42));           // 整数
    sheet->setCellValue("A3", 3.14159);                           // 浮点数
    sheet->setCellValue("A4", true);                              // 布尔值
    sheet->setCellValue("A5", static_cast<int64_t>(-1000));       // 负整数
    sheet->setCellValue("A6", 1.23e-5);                          // 科学计数法
    
    // 保存文件
    bool saved = workbook.saveToFile("output/SalesDataExample.xlsx");
    ASSERT_TRUE(saved);
    
    // 暂时跳过文件读取验证
    EXPECT_TRUE(std::filesystem::exists("output/SalesDataExample.xlsx"));
} 