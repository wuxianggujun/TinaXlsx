//
// Created by wuxianggujun on 2025/5/25.
//
#include <gtest/gtest.h>
#include "TinaXlsx/TXWorkbook.hpp"

// 测试示例，测试TXWorkbook类的基本功能
class TXWorkbookTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 在每个测试用例前调用，做初始化
    }

    void TearDown() override {
        // 在每个测试用例后调用，做清理工作
    }

    // 可共享的变量
    TinaXlsx::TXWorkbook workbook_;
};

// 测试基本功能：创建Workbook实例
TEST_F(TXWorkbookTest, CreateWorkbookInstance) {
    EXPECT_NO_THROW(TinaXlsx::TXWorkbook workbook);
}

// 测试功能：加载XLSX文件
TEST_F(TXWorkbookTest, LoadXlsxFile) {
    // 首先创建一个测试文件
    TinaXlsx::TXWorkbook temp_workbook;
    temp_workbook.addSheet("TestSheet");
    bool saved = temp_workbook.saveToFile("test.xlsx");
    ASSERT_TRUE(saved);
    
    // 然后加载这个文件
    bool loaded = workbook_.loadFromFile("test.xlsx");
    EXPECT_TRUE(loaded);
}

// 测试功能：写入XLSX文件
TEST_F(TXWorkbookTest, SaveXlsxFile) {
    workbook_.addSheet("TestSheet");
    bool saved = workbook_.saveToFile("output.xlsx");
    EXPECT_TRUE(saved);
}

// 测试异常情况：加载不存在的文件
TEST_F(TXWorkbookTest, LoadNonExistentFile) {
    bool loaded = workbook_.loadFromFile("nonexistent.xlsx");
    EXPECT_FALSE(loaded);
}
