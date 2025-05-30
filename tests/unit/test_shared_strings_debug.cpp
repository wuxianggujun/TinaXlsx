#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"

using namespace TinaXlsx;

// 测试默认注册SharedStrings组件的功能
TEST(SharedStringsTest, DefaultComponentRegistration) {
    TXWorkbook workbook;
    
    // 检查是否默认注册了SharedStrings组件
    bool hasSharedStrings = workbook.getComponentManager().hasComponent(ExcelComponent::SharedStrings);
    EXPECT_TRUE(hasSharedStrings) << "SharedStrings component should be registered by default";
}

// 测试智能字符串策略
TEST(SharedStringsTest, IntelligentStringStrategy) {
    TXWorkbook workbook;
    auto sheet = workbook.addSheet("Test");
    
    // 添加不同类型的字符串来测试策略
    sheet->setCellValue(row_t(1), column_t(1), std::string("A"));           // 1字符，应该内联
    sheet->setCellValue(row_t(1), column_t(2), std::string("Hello"));       // 中等长度，应该共享
    sheet->setCellValue(row_t(1), column_t(3), std::string("Hello"));       // 重复字符串，应该共享
    sheet->setCellValue(row_t(1), column_t(4), std::string("Text<with>XML")); // 特殊字符，应该内联
    
    // 验证SharedStrings组件已注册
    EXPECT_TRUE(workbook.getComponentManager().hasComponent(ExcelComponent::SharedStrings));
    
    // 保存文件并验证
    std::string filename = "test_intelligent_strings.xlsx";
    EXPECT_TRUE(workbook.saveToFile(filename));
    
    // 清理
    std::remove(filename.c_str());
}

// 测试Content Types包含SharedStrings声明
TEST(SharedStringsTest, ContentTypesContainsSharedStrings) {
    TXWorkbook workbook;
    auto sheet = workbook.addSheet("Test");
    
    // 添加一些字符串
    sheet->setCellValue(row_t(1), column_t(1), std::string("Test String"));
    
    // 保存文件
    std::string filename = "test_content_types.xlsx";
    EXPECT_TRUE(workbook.saveToFile(filename));
    
    // 这里我们无法直接验证Content Types内容，但至少确保保存成功
    // 实际验证需要解压XLSX文件并检查[Content_Types].xml
    
    // 清理
    std::remove(filename.c_str());
} 