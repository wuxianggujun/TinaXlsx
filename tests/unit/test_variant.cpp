//
// @file test_variant.cpp
// @brief TXVariant 单元测试
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXVariant.hpp"

using namespace TinaXlsx;

class VariantTest : public ::testing::Test {
protected:
    void SetUp() override {}
    void TearDown() override {}
};

TEST_F(VariantTest, DefaultConstruction) {
    TXVariant v;
    EXPECT_EQ(v.getType(), TXVariant::Type::Empty);
    EXPECT_TRUE(v.isEmpty());
}

TEST_F(VariantTest, NumberConstruction) {
    TXVariant v(42.5);
    EXPECT_EQ(v.getType(), TXVariant::Type::Number);
    EXPECT_FALSE(v.isEmpty());
    EXPECT_DOUBLE_EQ(v.asDouble(), 42.5);
}

TEST_F(VariantTest, StringConstruction) {
    TXVariant v("Hello World");
    EXPECT_EQ(v.getType(), TXVariant::Type::String);
    EXPECT_FALSE(v.isEmpty());
    EXPECT_EQ(v.asString(), "Hello World");
}

TEST_F(VariantTest, BooleanConstruction) {
    TXVariant v_true(true);
    TXVariant v_false(false);
    
    EXPECT_EQ(v_true.getType(), TXVariant::Type::Boolean);
    EXPECT_EQ(v_false.getType(), TXVariant::Type::Boolean);
    EXPECT_TRUE(v_true.asBool());
    EXPECT_FALSE(v_false.asBool());
}

TEST_F(VariantTest, FormulaConstruction) {
    TXVariant v("=A1+B1", TXVariant::Type::Formula);
    EXPECT_EQ(v.getType(), TXVariant::Type::Formula);
    EXPECT_EQ(v.asString(), "=A1+B1");
}

TEST_F(VariantTest, CopyConstruction) {
    TXVariant original(3.14159);
    TXVariant copy(original);
    
    EXPECT_EQ(copy.getType(), TXVariant::Type::Number);
    EXPECT_DOUBLE_EQ(copy.asDouble(), 3.14159);
}

TEST_F(VariantTest, Assignment) {
    TXVariant v1(42);
    TXVariant v2;
    
    v2 = v1;
    EXPECT_EQ(v2.getType(), TXVariant::Type::Number);
    EXPECT_DOUBLE_EQ(v2.asDouble(), 42.0);
}

TEST_F(VariantTest, TypeConversions) {
    // 数值到字符串
    TXVariant num(123.45);
    std::string str = num.toString();
    EXPECT_EQ(str, "123.45");
    
    // 布尔到字符串
    TXVariant bool_true(true);
    TXVariant bool_false(false);
    EXPECT_EQ(bool_true.toString(), "TRUE");
    EXPECT_EQ(bool_false.toString(), "FALSE");
    
    // 字符串数值到数值
    TXVariant str_num("987.65");
    EXPECT_DOUBLE_EQ(str_num.toNumber(), 987.65);
}

TEST_F(VariantTest, Equality) {
    TXVariant v1(42.0);
    TXVariant v2(42.0);
    TXVariant v3(43.0);
    
    EXPECT_TRUE(v1 == v2);
    EXPECT_FALSE(v1 == v3);
    EXPECT_TRUE(v1 != v3);
}

TEST_F(VariantTest, Performance) {
    // 测试大量Variant操作的性能
    const size_t COUNT = 10000;
    std::vector<TXVariant> variants;
    variants.reserve(COUNT);
    
    auto start = std::chrono::high_resolution_clock::now();
    
    for (size_t i = 0; i < COUNT; ++i) {
        variants.emplace_back(static_cast<double>(i));
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    EXPECT_EQ(variants.size(), COUNT);
    EXPECT_LT(duration.count(), 10000) << "10000个Variant创建应在10ms内完成";
    
    std::cout << "Variant性能: " << COUNT << "个对象在 " 
              << duration.count() << "μs 内创建" << std::endl;
} 