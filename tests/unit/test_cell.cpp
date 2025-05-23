#include <gtest/gtest.h>
#include <gmock/gmock.h>
#include "TinaXlsx/Cell.hpp"
#include "TinaXlsx/Types.hpp"

namespace TinaXlsx {
namespace Test {

/**
 * @brief Cell类的单元测试
 */
class CellTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 创建测试用的Cell对象
        cell = std::make_unique<Cell>();
    }

    void TearDown() override {
        cell.reset();
    }

    std::unique_ptr<Cell> cell;
};

/**
 * @brief 测试Cell的默认构造函数
 */
TEST_F(CellTest, DefaultConstructor) {
    Cell test_cell;
    EXPECT_TRUE(test_cell.isEmpty());
    EXPECT_EQ(test_cell.toString(), "");
}

/**
 * @brief 测试设置和获取字符串值
 */
TEST_F(CellTest, StringValue) {
    std::string test_value = "Hello World";
    
    cell->setValue(CellValue{test_value});
    
    EXPECT_TRUE(cell->isString());
    EXPECT_FALSE(cell->isEmpty());
    auto result = cell->getString();
    ASSERT_TRUE(result.has_value());
    EXPECT_EQ(result.value(), test_value);
}

/**
 * @brief 测试设置和获取整数值
 */
TEST_F(CellTest, IntegerValue) {
    int64_t test_value = 42;
    
    cell->setValue(CellValue{test_value});
    
    EXPECT_TRUE(cell->isInteger());
    EXPECT_FALSE(cell->isEmpty());
    auto result = cell->getInteger();
    ASSERT_TRUE(result.has_value());
    EXPECT_EQ(result.value(), test_value);
}

/**
 * @brief 测试设置和获取浮点数值
 */
TEST_F(CellTest, DoubleValue) {
    double test_value = 3.14159;
    
    cell->setValue(CellValue{test_value});
    
    EXPECT_TRUE(cell->isNumber());
    EXPECT_FALSE(cell->isEmpty());
    auto result = cell->getNumber();
    ASSERT_TRUE(result.has_value());
    EXPECT_DOUBLE_EQ(result.value(), test_value);
}

/**
 * @brief 测试设置和获取布尔值
 */
TEST_F(CellTest, BoolValue) {
    bool test_value = true;
    
    cell->setValue(CellValue{test_value});
    
    EXPECT_TRUE(cell->isBoolean());
    EXPECT_FALSE(cell->isEmpty());
    auto result = cell->getBoolean();
    ASSERT_TRUE(result.has_value());
    EXPECT_EQ(result.value(), test_value);
    
    // 测试false值
    cell->setValue(CellValue{false});
    result = cell->getBoolean();
    ASSERT_TRUE(result.has_value());
    EXPECT_EQ(result.value(), false);
}

/**
 * @brief 测试清空Cell
 */
TEST_F(CellTest, ClearCell) {
    // 先设置一个值
    cell->setValue(CellValue{std::string("Test Value")});
    EXPECT_TRUE(cell->isString());
    
    // 设置空值
    cell->setValue(CellValue{std::monostate{}});
    EXPECT_TRUE(cell->isEmpty());
    EXPECT_EQ(cell->toString(), "");
}

/**
 * @brief 测试Cell是否为空
 */
TEST_F(CellTest, IsEmpty) {
    // 新建的Cell应该是空的
    EXPECT_TRUE(cell->isEmpty());
    
    // 设置值后不应该是空的
    cell->setValue(CellValue{std::string("Not Empty")});
    EXPECT_FALSE(cell->isEmpty());
    
    // 设置空值后应该是空的
    cell->setValue(CellValue{std::monostate{}});
    EXPECT_TRUE(cell->isEmpty());
}

/**
 * @brief 测试Cell的拷贝构造函数
 */
TEST_F(CellTest, CopyConstructor) {
    cell->setValue(CellValue{std::string("Original Value")});
    
    Cell copied_cell(*cell);
    
    EXPECT_EQ(copied_cell.isString(), cell->isString());
    auto original_str = cell->getString();
    auto copied_str = copied_cell.getString();
    ASSERT_TRUE(original_str.has_value());
    ASSERT_TRUE(copied_str.has_value());
    EXPECT_EQ(copied_str.value(), original_str.value());
}

/**
 * @brief 测试Cell的赋值操作符
 */
TEST_F(CellTest, AssignmentOperator) {
    cell->setValue(CellValue{123.456});
    
    Cell assigned_cell;
    assigned_cell = *cell;
    
    EXPECT_EQ(assigned_cell.isNumber(), cell->isNumber());
    auto original_num = cell->getNumber();
    auto assigned_num = assigned_cell.getNumber();
    ASSERT_TRUE(original_num.has_value());
    ASSERT_TRUE(assigned_num.has_value());
    EXPECT_DOUBLE_EQ(assigned_num.value(), original_num.value());
}

/**
 * @brief 测试类型检查方法
 */
TEST_F(CellTest, TypeChecking) {
    // 测试字符串类型
    cell->setValue(CellValue{std::string("Test")});
    EXPECT_TRUE(cell->isString());
    EXPECT_FALSE(cell->isNumber());
    EXPECT_FALSE(cell->isInteger());
    EXPECT_FALSE(cell->isBoolean());
    EXPECT_FALSE(cell->isEmpty());
    
    // 测试数字类型
    cell->setValue(CellValue{3.14});
    EXPECT_FALSE(cell->isString());
    EXPECT_TRUE(cell->isNumber());
    EXPECT_FALSE(cell->isInteger());
    EXPECT_FALSE(cell->isBoolean());
    EXPECT_FALSE(cell->isEmpty());
    
    // 测试整数类型
    cell->setValue(CellValue{int64_t(42)});
    EXPECT_FALSE(cell->isString());
    EXPECT_FALSE(cell->isNumber());
    EXPECT_TRUE(cell->isInteger());
    EXPECT_FALSE(cell->isBoolean());
    EXPECT_FALSE(cell->isEmpty());
    
    // 测试布尔类型
    cell->setValue(CellValue{true});
    EXPECT_FALSE(cell->isString());
    EXPECT_FALSE(cell->isNumber());
    EXPECT_FALSE(cell->isInteger());
    EXPECT_TRUE(cell->isBoolean());
    EXPECT_FALSE(cell->isEmpty());
}

/**
 * @brief 测试toString方法
 */
TEST_F(CellTest, ToString) {
    // 测试字符串
    cell->setValue(CellValue{std::string("Hello")});
    EXPECT_EQ(cell->toString(), "Hello");
    
    // 测试空值
    cell->setValue(CellValue{std::monostate{}});
    EXPECT_EQ(cell->toString(), "");
}

} // namespace Test
} // namespace TinaXlsx 