/**
 * @file test_reader_new.cpp
 * @brief 测试基于minizip-ng的Reader实现
 */

#include <gtest/gtest.h>
#include "TinaXlsx/Reader.hpp"
#include "TinaXlsx/Exception.hpp"
#include <fstream>
#include <filesystem>

using namespace TinaXlsx;

class ReaderNewTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 设置测试环境
    }
    
    void TearDown() override {
        // 清理测试环境
    }
};

// 测试Reader构造函数异常处理
TEST_F(ReaderNewTest, ConstructorWithInvalidFile) {
    // 测试打开不存在的文件
    EXPECT_THROW(Reader reader("nonexistent.xlsx"), FileException);
}

// 测试静态工具函数
TEST_F(ReaderNewTest, StaticUtilityFunctions) {
    // 测试空行判断
    RowData emptyRow;
    EXPECT_TRUE(Reader::isEmptyRow(emptyRow));
    
    RowData rowWithEmptyCells = {CellValue(), CellValue(), CellValue()};
    EXPECT_TRUE(Reader::isEmptyRow(rowWithEmptyCells));
    
    RowData rowWithData = {CellValue(std::string("Hello")), CellValue(), CellValue()};
    EXPECT_FALSE(Reader::isEmptyRow(rowWithData));
    
    // 测试空单元格判断
    EXPECT_TRUE(Reader::isEmptyCell(CellValue()));
    EXPECT_TRUE(Reader::isEmptyCell(CellValue(std::string(""))));
    EXPECT_FALSE(Reader::isEmptyCell(CellValue(std::string("Hello"))));
    EXPECT_FALSE(Reader::isEmptyCell(CellValue(static_cast<Integer>(123))));
    EXPECT_FALSE(Reader::isEmptyCell(CellValue(3.14)));
    EXPECT_FALSE(Reader::isEmptyCell(CellValue(true)));
}

// 测试字符串到CellValue的转换
TEST_F(ReaderNewTest, StringToCellValueConversion) {
    // 空字符串
    auto value1 = Reader::stringToCellValue("");
    EXPECT_TRUE(std::holds_alternative<std::monostate>(value1));
    
    // 整数
    auto value2 = Reader::stringToCellValue("42");
    EXPECT_TRUE(std::holds_alternative<Integer>(value2));
    EXPECT_EQ(std::get<Integer>(value2), 42);
    
    // 浮点数
    auto value3 = Reader::stringToCellValue("3.14");
    EXPECT_TRUE(std::holds_alternative<double>(value3));
    EXPECT_DOUBLE_EQ(std::get<double>(value3), 3.14);
    
    // 布尔值
    auto value4 = Reader::stringToCellValue("true");
    EXPECT_TRUE(std::holds_alternative<bool>(value4));
    EXPECT_TRUE(std::get<bool>(value4));
    
    auto value5 = Reader::stringToCellValue("false");
    EXPECT_TRUE(std::holds_alternative<bool>(value5));
    EXPECT_FALSE(std::get<bool>(value5));
    
    // 普通字符串
    auto value6 = Reader::stringToCellValue("hello");
    EXPECT_TRUE(std::holds_alternative<std::string>(value6));
    EXPECT_EQ(std::get<std::string>(value6), "hello");
}

// 测试CellValue到字符串的转换
TEST_F(ReaderNewTest, CellValueToStringConversion) {
    EXPECT_EQ(Reader::cellValueToString(std::monostate{}), "");
    EXPECT_EQ(Reader::cellValueToString(std::string("hello")), "hello");
    EXPECT_EQ(Reader::cellValueToString(static_cast<Integer>(42)), "42");
    EXPECT_EQ(Reader::cellValueToString(3.14), "3.140000");
    EXPECT_EQ(Reader::cellValueToString(true), "true");
    EXPECT_EQ(Reader::cellValueToString(false), "false");
}

// 基本功能测试（无需真实Excel文件）
TEST_F(ReaderNewTest, BasicFunctionality) {
    // 这里测试不需要真实Excel文件的基本功能
    
    // 由于我们没有真实的Excel文件，这个测试主要验证接口的存在
    // 在实际使用中，需要提供真实的Excel文件进行测试
    
    try {
        // 创建一个简单的文件来测试（虽然不是有效的Excel文件）
        std::string testFile = "test.txt";
        std::ofstream file(testFile);
        file << "test";
        file.close();
        
        // 这应该会抛出异常，因为不是有效的Excel文件
        EXPECT_THROW({
            Reader reader(testFile);
        }, FileException);
        
        // 清理测试文件
        std::filesystem::remove(testFile);
    } catch (...) {
        // 预期的异常
    }
} 