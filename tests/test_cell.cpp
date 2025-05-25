#include <gtest/gtest.h>
#include "TinaXlsx/TXCell.hpp"

class TXCellTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 在每个测试开始前的设置
    }

    void TearDown() override {
        // 在每个测试结束后的清理
    }
};

TEST_F(TXCellTest, DefaultConstructor) {
    TinaXlsx::TXCell cell;
    
    EXPECT_TRUE(cell.isEmpty());
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Empty);
    EXPECT_EQ(cell.getStringValue(), "");
    EXPECT_EQ(cell.getNumberValue(), 0.0);
    EXPECT_EQ(cell.getIntegerValue(), 0);
    EXPECT_FALSE(cell.getBooleanValue());
}

TEST_F(TXCellTest, StringValue) {
    TinaXlsx::TXCell cell;
    
    // 测试设置字符串值
    cell.setStringValue("Hello, World!");
    EXPECT_FALSE(cell.isEmpty());
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::String);
    EXPECT_EQ(cell.getStringValue(), "Hello, World!");
    
    // 测试赋值操作符
    cell = "Test String";
    EXPECT_EQ(cell.getStringValue(), "Test String");
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::String);
    
    // 测试类型转换操作符
    std::string str = cell;
    EXPECT_EQ(str, "Test String");
}

TEST_F(TXCellTest, NumberValue) {
    TinaXlsx::TXCell cell;
    
    // 测试设置数字值
    cell.setNumberValue(3.14159);
    EXPECT_FALSE(cell.isEmpty());
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Number);
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), 3.14159);
    
    // 测试赋值操作符
    cell = 2.71828;
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), 2.71828);
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Number);
    
    // 测试类型转换操作符
    double num = cell;
    EXPECT_DOUBLE_EQ(num, 2.71828);
}

TEST_F(TXCellTest, IntegerValue) {
    TinaXlsx::TXCell cell;
    
    // 测试设置整数值
    cell.setIntegerValue(42);
    EXPECT_FALSE(cell.isEmpty());
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Integer);
    EXPECT_EQ(cell.getIntegerValue(), 42);
    
    // 测试赋值操作符（int64_t）
    cell = static_cast<int64_t>(1000000);
    EXPECT_EQ(cell.getIntegerValue(), 1000000);
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Integer);
    
    // 测试赋值操作符（int）
    cell = 123;
    EXPECT_EQ(cell.getIntegerValue(), 123);
    
    // 测试类型转换操作符
    int64_t num = cell;
    EXPECT_EQ(num, 123);
}

TEST_F(TXCellTest, BooleanValue) {
    TinaXlsx::TXCell cell;
    
    // 测试设置布尔值
    cell.setBooleanValue(true);
    EXPECT_FALSE(cell.isEmpty());
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Boolean);
    EXPECT_TRUE(cell.getBooleanValue());
    
    // 测试赋值操作符
    cell = false;
    EXPECT_FALSE(cell.getBooleanValue());
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Boolean);
    
    // 测试类型转换操作符
    bool val = cell;
    EXPECT_FALSE(val);
}

TEST_F(TXCellTest, TypeConversion) {
    TinaXlsx::TXCell cell;
    
    // 字符串转数字
    cell = "123.45";
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), 123.45);
    EXPECT_EQ(cell.getIntegerValue(), 123);
    
    // 数字转字符串
    cell = 456.78;
    EXPECT_EQ(cell.getStringValue(), "456.780000");
    
    // 布尔值转字符串
    cell = true;
    EXPECT_EQ(cell.getStringValue(), "TRUE");
    
    cell = false;
    EXPECT_EQ(cell.getStringValue(), "FALSE");
    
    // 字符串转布尔值
    cell = "true";
    EXPECT_TRUE(cell.getBooleanValue());
    
    cell = "false";
    EXPECT_FALSE(cell.getBooleanValue());
    
    cell = "1";
    EXPECT_TRUE(cell.getBooleanValue());
    
    cell = "0";
    EXPECT_FALSE(cell.getBooleanValue());
}

TEST_F(TXCellTest, Formula) {
    TinaXlsx::TXCell cell;
    
    // 测试设置公式
    cell.setFormula("SUM(A1:A10)");
    EXPECT_TRUE(cell.isFormula());
    EXPECT_EQ(cell.getFormula(), "SUM(A1:A10)");
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Formula);
    
    // 设置值后公式应该被清除
    cell.setValue(std::string("Normal value"));
    EXPECT_FALSE(cell.isFormula());
    EXPECT_TRUE(cell.getFormula().empty());
}

TEST_F(TXCellTest, NumberFormat) {
    TinaXlsx::TXCell cell;
    
    // 测试默认格式
    EXPECT_EQ(cell.getNumberFormat(), TinaXlsx::TXCell::NumberFormat::General);
    
    // 测试设置数字格式
    cell.setNumberFormat(TinaXlsx::TXCell::NumberFormat::Currency);
    EXPECT_EQ(cell.getNumberFormat(), TinaXlsx::TXCell::NumberFormat::Currency);
    
    // 测试自定义格式
    cell.setCustomFormat("#,##0.00");
    EXPECT_EQ(cell.getCustomFormat(), "#,##0.00");
}

TEST_F(TXCellTest, ToString) {
    TinaXlsx::TXCell cell;
    
    cell = "Test String";
    EXPECT_EQ(cell.toString(), "Test String");
    
    cell = 123.45;
    EXPECT_EQ(cell.toString(), "123.450000");
    
    cell = static_cast<int64_t>(789);
    EXPECT_EQ(cell.toString(), "789");
    
    cell = true;
    EXPECT_EQ(cell.toString(), "TRUE");
}

TEST_F(TXCellTest, FromString) {
    TinaXlsx::TXCell cell;
    
    // 测试自动类型检测
    EXPECT_TRUE(cell.fromString("Hello", true));
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::String);
    EXPECT_EQ(cell.getStringValue(), "Hello");
    
    EXPECT_TRUE(cell.fromString("123", true));
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Integer);
    EXPECT_EQ(cell.getIntegerValue(), 123);
    
    EXPECT_TRUE(cell.fromString("123.45", true));
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Number);
    EXPECT_DOUBLE_EQ(cell.getNumberValue(), 123.45);
    
    EXPECT_TRUE(cell.fromString("true", true));
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Boolean);
    EXPECT_TRUE(cell.getBooleanValue());
    
    EXPECT_TRUE(cell.fromString("false", true));
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Boolean);
    EXPECT_FALSE(cell.getBooleanValue());
    
    // 测试不自动检测类型
    EXPECT_TRUE(cell.fromString("123", false));
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::String);
    EXPECT_EQ(cell.getStringValue(), "123");
    
    // 测试空字符串
    EXPECT_TRUE(cell.fromString("", true));
    EXPECT_TRUE(cell.isEmpty());
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Empty);
}

TEST_F(TXCellTest, Clear) {
    TinaXlsx::TXCell cell;
    
    cell = "Test Value";
    cell.setFormula("SUM(A1:A10)");
    cell.setNumberFormat(TinaXlsx::TXCell::NumberFormat::Currency);
    cell.setCustomFormat("#,##0.00");
    
    EXPECT_FALSE(cell.isEmpty());
    
    cell.clear();
    
    EXPECT_TRUE(cell.isEmpty());
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::Empty);
    EXPECT_TRUE(cell.getFormula().empty());
    EXPECT_EQ(cell.getNumberFormat(), TinaXlsx::TXCell::NumberFormat::General);
    EXPECT_TRUE(cell.getCustomFormat().empty());
}

TEST_F(TXCellTest, CopyConstructor) {
    TinaXlsx::TXCell cell1;
    cell1 = "Original Value";
    cell1.setFormula("SUM(A1:A10)");
    cell1.setNumberFormat(TinaXlsx::TXCell::NumberFormat::Currency);
    
    // 测试拷贝构造
    TinaXlsx::TXCell cell2(cell1);
    EXPECT_EQ(cell2.getStringValue(), "Original Value");
    EXPECT_EQ(cell2.getFormula(), "SUM(A1:A10)");
    EXPECT_EQ(cell2.getNumberFormat(), TinaXlsx::TXCell::NumberFormat::Currency);
    
    // 修改原始单元格不应影响拷贝
    cell1 = "Modified Value";
    EXPECT_EQ(cell1.getStringValue(), "Modified Value");
    EXPECT_EQ(cell2.getStringValue(), "Original Value");
}

TEST_F(TXCellTest, AssignmentOperator) {
    TinaXlsx::TXCell cell1, cell2;
    
    cell1 = "Test Value";
    cell1.setNumberFormat(TinaXlsx::TXCell::NumberFormat::Percentage);
    
    // 测试赋值操作符
    cell2 = cell1;
    EXPECT_EQ(cell2.getStringValue(), "Test Value");
    EXPECT_EQ(cell2.getNumberFormat(), TinaXlsx::TXCell::NumberFormat::Percentage);
    
    // 修改原始单元格不应影响赋值后的单元格
    cell1 = "Different Value";
    EXPECT_EQ(cell1.getStringValue(), "Different Value");
    EXPECT_EQ(cell2.getStringValue(), "Test Value");
}

TEST_F(TXCellTest, MoveConstructor) {
    TinaXlsx::TXCell cell1;
    cell1 = "Move Test";
    
    // 测试移动构造
    TinaXlsx::TXCell cell2(std::move(cell1));
    EXPECT_EQ(cell2.getStringValue(), "Move Test");
    
    // 原始单元格应该处于有效但未指定状态
    // 我们不能对其内容做假设，但它应该仍然可用
}

TEST_F(TXCellTest, MoveAssignment) {
    TinaXlsx::TXCell cell1, cell2;
    cell1 = "Move Assignment Test";
    
    // 测试移动赋值
    cell2 = std::move(cell1);
    EXPECT_EQ(cell2.getStringValue(), "Move Assignment Test");
}

TEST_F(TXCellTest, ComparisonOperators) {
    TinaXlsx::TXCell cell1, cell2;
    
    // 测试相等比较
    cell1 = "Same Value";
    cell2 = "Same Value";
    EXPECT_TRUE(cell1 == cell2);
    EXPECT_FALSE(cell1 != cell2);
    
    // 测试不等比较
    cell2 = "Different Value";
    EXPECT_FALSE(cell1 == cell2);
    EXPECT_TRUE(cell1 != cell2);
    
    // 测试不同类型的比较
    cell1 = 123;
    cell2 = 123.0;
    // 注意：这里测试的是variant的比较，int64_t和double是不同类型
    EXPECT_FALSE(cell1 == cell2);
    
    // 测试相同数值不同类型
    cell1 = static_cast<int64_t>(100);
    cell2 = static_cast<int64_t>(100);
    EXPECT_TRUE(cell1 == cell2);
}

TEST_F(TXCellTest, ValueConstructor) {
    // 测试使用值构造
    TinaXlsx::TXCell::CellValue value = std::string("Constructor Test");
    TinaXlsx::TXCell cell(value);
    
    EXPECT_EQ(cell.getStringValue(), "Constructor Test");
    EXPECT_EQ(cell.getType(), TinaXlsx::TXCell::CellType::String);
    EXPECT_FALSE(cell.isEmpty());
} 