//
// 数据完整性调试测试 (GTest版本)
//
#include <gtest/gtest.h>
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include <iostream>

class DataIntegrityTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 测试前准备
    }

    void TearDown() override {
        // 测试后清理
    }
};

TEST_F(DataIntegrityTest, BasicDataIntegrity) {
    // 创建简单的工作簿
    TinaXlsx::TXWorkbook workbook;
    auto sheet = workbook.addSheet("测试数据");
    ASSERT_NE(sheet, nullptr);
    
    // 添加测试数据
    sheet->setCellValue("A1", std::string("文本测试"));
    sheet->setCellValue("B1", static_cast<int64_t>(123));
    sheet->setCellValue("C1", 3.14159);
    
    // 立即验证数据是否写入成功
    auto value_a1 = sheet->getCellValue("A1");
    auto value_b1 = sheet->getCellValue("B1");
    auto value_c1 = sheet->getCellValue("C1");
    
    // 验证 variant 类型
    EXPECT_EQ(value_a1.index(), 1) << "A1 应该是 string 类型 (index 1)";  // string
    EXPECT_EQ(value_b1.index(), 2) << "B1 应该是 int64_t 类型 (index 2)"; // int64_t
    EXPECT_EQ(value_c1.index(), 3) << "C1 应该是 double 类型 (index 3)";  // double
    
    // 验证具体值
    EXPECT_TRUE(std::holds_alternative<std::string>(value_a1));
    EXPECT_TRUE(std::holds_alternative<int64_t>(value_b1));
    EXPECT_TRUE(std::holds_alternative<double>(value_c1));
    
    if (std::holds_alternative<std::string>(value_a1)) {
        EXPECT_EQ(std::get<std::string>(value_a1), "文本测试");
    }
    if (std::holds_alternative<int64_t>(value_b1)) {
        EXPECT_EQ(std::get<int64_t>(value_b1), 123);
    }
    if (std::holds_alternative<double>(value_c1)) {
        EXPECT_DOUBLE_EQ(std::get<double>(value_c1), 3.14159);
    }
}

TEST_F(DataIntegrityTest, EmptyAndNullValues) {
    TinaXlsx::TXWorkbook workbook;
    auto sheet = workbook.addSheet("空值测试");
    ASSERT_NE(sheet, nullptr);
    
    // 测试空字符串
    sheet->setCellValue("A1", std::string(""));
    
    // 测试零值
    sheet->setCellValue("B1", static_cast<int64_t>(0));
    sheet->setCellValue("C1", 0.0);
    
    // 验证空值处理
    auto val_a1 = sheet->getCellValue("A1");
    auto val_b1 = sheet->getCellValue("B1");
    auto val_c1 = sheet->getCellValue("C1");
    
    EXPECT_TRUE(std::holds_alternative<std::string>(val_a1));
    EXPECT_TRUE(std::holds_alternative<int64_t>(val_b1));
    EXPECT_TRUE(std::holds_alternative<double>(val_c1));
    
    if (std::holds_alternative<std::string>(val_a1)) {
        EXPECT_EQ(std::get<std::string>(val_a1), "");
    }
    if (std::holds_alternative<int64_t>(val_b1)) {
        EXPECT_EQ(std::get<int64_t>(val_b1), 0);
    }
    if (std::holds_alternative<double>(val_c1)) {
        EXPECT_DOUBLE_EQ(std::get<double>(val_c1), 0.0);
    }
}

TEST_F(DataIntegrityTest, UnicodeStringIntegrity) {
    TinaXlsx::TXWorkbook workbook;
    auto sheet = workbook.addSheet("Unicode测试");
    ASSERT_NE(sheet, nullptr);
    
    // 测试各种Unicode字符
    std::vector<std::string> unicode_strings = {
        "中文测试",
        "English Test",
        "🚀 Emoji Test",
        "Русский текст",
        "العربية",
        "αβγδε",
        "Special chars: !@#$%^&*()"
    };
    
    // 写入Unicode字符串
    for (size_t i = 0; i < unicode_strings.size(); ++i) {
        std::string address = "A" + std::to_string(i + 1);
        sheet->setCellValue(address, unicode_strings[i]);
    }
    
    // 验证Unicode字符串完整性
    for (size_t i = 0; i < unicode_strings.size(); ++i) {
        std::string address = "A" + std::to_string(i + 1);
        auto value = sheet->getCellValue(address);
        
        EXPECT_TRUE(std::holds_alternative<std::string>(value)) 
            << "位置 " << address << " 应该包含字符串";
        
        if (std::holds_alternative<std::string>(value)) {
            EXPECT_EQ(std::get<std::string>(value), unicode_strings[i])
                << "位置 " << address << " 的Unicode字符串不匹配";
        }
    }
}

TEST_F(DataIntegrityTest, LargeNumberIntegrity) {
    TinaXlsx::TXWorkbook workbook;
    auto sheet = workbook.addSheet("大数值测试");
    ASSERT_NE(sheet, nullptr);
    
    // 测试各种边界值
    std::vector<int64_t> int_values = {
        0, 1, -1,
        1000000, -1000000,
        INT64_MAX, INT64_MIN,
        9223372036854775807LL,  // INT64_MAX
        -9223372036854775808LL  // INT64_MIN
    };
    
    std::vector<double> double_values = {
        0.0, 1.0, -1.0,
        3.14159265359, -3.14159265359,
        1.23456789e10, -1.23456789e10,
        1.23456789e-10, -1.23456789e-10
    };
    
    // 写入整数值
    for (size_t i = 0; i < int_values.size(); ++i) {
        std::string address = "A" + std::to_string(i + 1);
        sheet->setCellValue(address, int_values[i]);
    }
    
    // 写入浮点数值
    for (size_t i = 0; i < double_values.size(); ++i) {
        std::string address = "B" + std::to_string(i + 1);
        sheet->setCellValue(address, double_values[i]);
    }
    
    // 验证整数值
    for (size_t i = 0; i < int_values.size(); ++i) {
        std::string address = "A" + std::to_string(i + 1);
        auto value = sheet->getCellValue(address);
        
        EXPECT_TRUE(std::holds_alternative<int64_t>(value))
            << "位置 " << address << " 应该包含int64_t";
        
        if (std::holds_alternative<int64_t>(value)) {
            EXPECT_EQ(std::get<int64_t>(value), int_values[i])
                << "位置 " << address << " 的整数值不匹配";
        }
    }
    
    // 验证浮点数值
    for (size_t i = 0; i < double_values.size(); ++i) {
        std::string address = "B" + std::to_string(i + 1);
        auto value = sheet->getCellValue(address);
        
        EXPECT_TRUE(std::holds_alternative<double>(value))
            << "位置 " << address << " 应该包含double";
        
        if (std::holds_alternative<double>(value)) {
            EXPECT_DOUBLE_EQ(std::get<double>(value), double_values[i])
                << "位置 " << address << " 的浮点数值不匹配";
        }
    }
}