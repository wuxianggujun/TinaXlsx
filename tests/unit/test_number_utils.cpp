//
// @file test_number_utils.cpp
// @brief TXNumberUtils 高性能数值工具类测试
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXNumberUtils.hpp"
#include <chrono>
#include <vector>
#include <random>

using namespace TinaXlsx;

class NumberUtilsTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 生成测试数据
        std::random_device rd;
        std::mt19937 gen(rd());
        std::uniform_real_distribution<double> dis(-1000000.0, 1000000.0);
        
        testNumbers.reserve(10000);
        testStrings.reserve(10000);
        
        for (int i = 0; i < 10000; ++i) {
            double num = dis(gen);
            testNumbers.push_back(num);
            testStrings.push_back(std::to_string(num));
        }
    }

    std::vector<double> testNumbers;
    std::vector<std::string> testStrings;
};

// ==================== 基础解析测试 ====================

TEST_F(NumberUtilsTest, ParseDoubleBasic) {
    // 测试基本解析功能
    EXPECT_EQ(TXNumberUtils::parseDouble("123.456"), 123.456);
    EXPECT_EQ(TXNumberUtils::parseDouble("0"), 0.0);
    EXPECT_EQ(TXNumberUtils::parseDouble("-123.456"), -123.456);
    EXPECT_EQ(TXNumberUtils::parseDouble("1e6"), 1000000.0);
    EXPECT_EQ(TXNumberUtils::parseDouble("1.23e-4"), 0.000123);
    
    // 测试无效输入
    EXPECT_EQ(TXNumberUtils::parseDouble("abc"), std::nullopt);
    EXPECT_EQ(TXNumberUtils::parseDouble(""), std::nullopt);
    EXPECT_EQ(TXNumberUtils::parseDouble("123abc"), std::nullopt);
}

TEST_F(NumberUtilsTest, ParseFloatBasic) {
    EXPECT_FLOAT_EQ(TXNumberUtils::parseFloat("123.456").value_or(0), 123.456f);
    EXPECT_FLOAT_EQ(TXNumberUtils::parseFloat("0").value_or(0), 0.0f);
    EXPECT_FLOAT_EQ(TXNumberUtils::parseFloat("-123.456").value_or(0), -123.456f);
}

TEST_F(NumberUtilsTest, ParseInt64Basic) {
    EXPECT_EQ(TXNumberUtils::parseInt64("123"), 123);
    EXPECT_EQ(TXNumberUtils::parseInt64("0"), 0);
    EXPECT_EQ(TXNumberUtils::parseInt64("-123"), -123);
    
    // 测试非整数
    EXPECT_EQ(TXNumberUtils::parseInt64("123.456"), std::nullopt);
    EXPECT_EQ(TXNumberUtils::parseInt64("abc"), std::nullopt);
}

// ==================== 格式化测试 ====================

TEST_F(NumberUtilsTest, FormatForExcelXml) {
    // 测试整数格式化
    EXPECT_EQ(TXNumberUtils::formatForExcelXml(3000.0), "3000");
    EXPECT_EQ(TXNumberUtils::formatForExcelXml(0.0), "0");
    EXPECT_EQ(TXNumberUtils::formatForExcelXml(-1000.0), "-1000");
    
    // 测试小数格式化
    EXPECT_EQ(TXNumberUtils::formatForExcelXml(123.45), "123.45");
    EXPECT_EQ(TXNumberUtils::formatForExcelXml(123.40), "123.4");
    EXPECT_EQ(TXNumberUtils::formatForExcelXml(123.00), "123");
    
    // 测试精度
    EXPECT_EQ(TXNumberUtils::formatForExcelXml(123.456789), "123.46");
}

TEST_F(NumberUtilsTest, FormatDoubleWithOptions) {
    TXNumberUtils::FormatOptions options;
    
    // 测试精度控制
    options.precision = 3;
    options.removeTrailingZeros = false;
    EXPECT_EQ(TXNumberUtils::formatDouble(123.456789, options), "123.457");
    
    // 测试移除尾随零
    options.removeTrailingZeros = true;
    EXPECT_EQ(TXNumberUtils::formatDouble(123.400, options), "123.4");
    
    // 测试千位分隔符
    options.useThousandSeparator = true;
    options.precision = 2;
    EXPECT_EQ(TXNumberUtils::formatDouble(1234567.89, options), "1,234,567.89");
}

// ==================== 工具方法测试 ====================

TEST_F(NumberUtilsTest, IsValidNumber) {
    EXPECT_TRUE(TXNumberUtils::isValidNumber("123.456"));
    EXPECT_TRUE(TXNumberUtils::isValidNumber("0"));
    EXPECT_TRUE(TXNumberUtils::isValidNumber("-123"));
    EXPECT_TRUE(TXNumberUtils::isValidNumber("1e6"));
    
    EXPECT_FALSE(TXNumberUtils::isValidNumber("abc"));
    EXPECT_FALSE(TXNumberUtils::isValidNumber(""));
    EXPECT_FALSE(TXNumberUtils::isValidNumber("123abc"));
}

TEST_F(NumberUtilsTest, IsInteger) {
    EXPECT_TRUE(TXNumberUtils::isInteger(123.0));
    EXPECT_TRUE(TXNumberUtils::isInteger(0.0));
    EXPECT_TRUE(TXNumberUtils::isInteger(-123.0));
    
    EXPECT_FALSE(TXNumberUtils::isInteger(123.456));
    EXPECT_FALSE(TXNumberUtils::isInteger(0.1));
}

TEST_F(NumberUtilsTest, RemoveTrailingZeros) {
    EXPECT_EQ(TXNumberUtils::removeTrailingZeros("123.000"), "123");
    EXPECT_EQ(TXNumberUtils::removeTrailingZeros("123.450"), "123.45");
    EXPECT_EQ(TXNumberUtils::removeTrailingZeros("123.456"), "123.456");
    EXPECT_EQ(TXNumberUtils::removeTrailingZeros("123"), "123");
}

// ==================== 性能测试 ====================

TEST_F(NumberUtilsTest, ParsePerformance) {
    const int iterations = 10000;
    
    // 测试 fast_float 解析性能
    auto start = std::chrono::high_resolution_clock::now();
    
    for (int i = 0; i < iterations; ++i) {
        double result;
        TXNumberUtils::parseDouble(testStrings[i % testStrings.size()], result);
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto fastFloatTime = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    // 测试标准库解析性能
    start = std::chrono::high_resolution_clock::now();
    
    for (int i = 0; i < iterations; ++i) {
        try {
            std::stod(testStrings[i % testStrings.size()]);
        } catch (...) {
            // 忽略错误
        }
    }
    
    end = std::chrono::high_resolution_clock::now();
    auto stdTime = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "🚀 性能测试结果 (" << iterations << " 次解析):" << std::endl;
    std::cout << "  fast_float: " << fastFloatTime.count() << " μs" << std::endl;
    std::cout << "  std::stod:  " << stdTime.count() << " μs" << std::endl;
    std::cout << "  性能提升:   " << (double)stdTime.count() / fastFloatTime.count() << "x" << std::endl;
    
    // fast_float 应该比标准库更快
    EXPECT_LT(fastFloatTime.count(), stdTime.count());
}

TEST_F(NumberUtilsTest, FormatPerformance) {
    const int iterations = 10000;
    
    // 测试我们的格式化性能
    auto start = std::chrono::high_resolution_clock::now();
    
    for (int i = 0; i < iterations; ++i) {
        TXNumberUtils::formatForExcelXml(testNumbers[i % testNumbers.size()]);
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto ourTime = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    // 测试标准库格式化性能
    start = std::chrono::high_resolution_clock::now();
    
    for (int i = 0; i < iterations; ++i) {
        std::to_string(testNumbers[i % testNumbers.size()]);
    }
    
    end = std::chrono::high_resolution_clock::now();
    auto stdTime = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "📊 格式化性能测试结果 (" << iterations << " 次格式化):" << std::endl;
    std::cout << "  TXNumberUtils: " << ourTime.count() << " μs" << std::endl;
    std::cout << "  std::to_string: " << stdTime.count() << " μs" << std::endl;
    
    // 输出结果用于分析
    EXPECT_GT(ourTime.count(), 0);
    EXPECT_GT(stdTime.count(), 0);
}

// ==================== Excel兼容性测试 ====================

TEST_F(NumberUtilsTest, ExcelCompatibility) {
    // 测试与Excel筛选兼容的格式
    struct TestCase {
        double input;
        std::string expected;
    };
    
    std::vector<TestCase> testCases = {
        {3000.0, "3000"},
        {3000.000000, "3000"},
        {123.45, "123.45"},
        {123.40, "123.4"},
        {0.0, "0"},
        {-1000.0, "-1000"},
        {1234.567890, "1234.57"}
    };
    
    for (const auto& testCase : testCases) {
        std::string result = TXNumberUtils::formatForExcelXml(testCase.input);
        EXPECT_EQ(result, testCase.expected) 
            << "Input: " << testCase.input 
            << ", Expected: " << testCase.expected 
            << ", Got: " << result;
    }
}

// ==================== 错误处理测试 ====================

TEST_F(NumberUtilsTest, ErrorHandling) {
    double result;
    
    // 测试各种错误情况
    EXPECT_EQ(TXNumberUtils::parseDouble("", result), TXNumberUtils::ParseResult::Empty);
    EXPECT_EQ(TXNumberUtils::parseDouble("   ", result), TXNumberUtils::ParseResult::Empty);
    EXPECT_EQ(TXNumberUtils::parseDouble("abc", result), TXNumberUtils::ParseResult::InvalidFormat);
    EXPECT_EQ(TXNumberUtils::parseDouble("123abc", result), TXNumberUtils::ParseResult::InvalidFormat);
    
    // 测试错误描述
    EXPECT_EQ(TXNumberUtils::getParseErrorDescription(TXNumberUtils::ParseResult::Success), "Success");
    EXPECT_EQ(TXNumberUtils::getParseErrorDescription(TXNumberUtils::ParseResult::InvalidFormat), "Invalid number format");
    EXPECT_EQ(TXNumberUtils::getParseErrorDescription(TXNumberUtils::ParseResult::OutOfRange), "Number out of range");
    EXPECT_EQ(TXNumberUtils::getParseErrorDescription(TXNumberUtils::ParseResult::Empty), "Empty string");
}
