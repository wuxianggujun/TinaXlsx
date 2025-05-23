#include <gtest/gtest.h>
#include <gmock/gmock.h>
#include "TinaXlsx/TinaXlsx.hpp"

namespace TinaXlsx {
namespace Test {

/**
 * @brief Basic test class to verify test framework works properly
 */
class BasicTest : public ::testing::Test {
protected:
    void SetUp() override {
        // Setup before tests
    }

    void TearDown() override {
        // Cleanup after tests
    }
};

/**
 * @brief Test if Google Test framework works
 */
TEST_F(BasicTest, GoogleTestFrameworkWorks) {
    EXPECT_EQ(1 + 1, 2);
    EXPECT_TRUE(true);
    EXPECT_FALSE(false);
}

/**
 * @brief Test if project headers can be included properly
 */
TEST_F(BasicTest, IncludeHeadersWork) {
    // If headers have syntax errors, compilation will fail
    EXPECT_TRUE(true) << "Headers included successfully";
}

/**
 * @brief Test basic math operations
 */
TEST_F(BasicTest, BasicMathOperations) {
    EXPECT_EQ(2 + 3, 5);
    EXPECT_EQ(10 - 4, 6);
    EXPECT_EQ(3 * 4, 12);
    EXPECT_EQ(15 / 3, 5);
}

/**
 * @brief Test string operations
 */
TEST_F(BasicTest, StringOperations) {
    std::string test_str = "Hello";
    test_str += " World";
    
    EXPECT_EQ(test_str, "Hello World");
    EXPECT_EQ(test_str.length(), 11);
    EXPECT_TRUE(test_str.find("World") != std::string::npos);
}

} // namespace Test
} // namespace TinaXlsx 