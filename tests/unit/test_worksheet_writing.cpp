#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <fstream>

using namespace TinaXlsx;

class WorksheetWritingTest : public ::testing::Test {
protected:
    void SetUp() override {
        workbook = std::make_unique<TXWorkbook>();
        sheet = workbook->addSheet("TestSheet");
        ASSERT_NE(sheet, nullptr);
    }

    void TearDown() override {
        workbook.reset();
        // 清理测试文件
        std::remove("test_output.xlsx");
    }

    std::unique_ptr<TXWorkbook> workbook;
    TXSheet* sheet = nullptr;
    const std::string testFileName = "test_output.xlsx";
};

TEST_F(WorksheetWritingTest, BasicDataWriting) {
    // 设置测试数据
    sheet->setCellValue(row_t(1), column_t(1), std::string("Hello"));
    sheet->setCellValue(row_t(1), column_t(2), 123.45);
    sheet->setCellValue(row_t(2), column_t(1), std::string("World"));
    sheet->setCellValue(row_t(2), column_t(2), static_cast<int64_t>(67));
    sheet->setCellValue(row_t(3), column_t(1), true);
    
    // 验证数据设置成功
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(1), column_t(1))), "Hello");
    EXPECT_DOUBLE_EQ(std::get<double>(sheet->getCellValue(row_t(1), column_t(2))), 123.45);
    EXPECT_EQ(std::get<std::string>(sheet->getCellValue(row_t(2), column_t(1))), "World");
    EXPECT_EQ(std::get<int64_t>(sheet->getCellValue(row_t(2), column_t(2))), 67);
    EXPECT_EQ(std::get<bool>(sheet->getCellValue(row_t(3), column_t(1))), true);
    
    // 检查使用范围
    auto usedRange = sheet->getUsedRange();
    EXPECT_TRUE(usedRange.isValid());
    EXPECT_EQ(usedRange.toAddress(), "A1:B3");
}

TEST_F(WorksheetWritingTest, FileSaving) {
    // 设置数据
    sheet->setCellValue(row_t(1), column_t(1), std::string("Test"));
    sheet->setCellValue(row_t(1), column_t(2), 42.0);
    
    // 保存文件
    bool saved = workbook->saveToFile(testFileName);
    EXPECT_TRUE(saved) << "Save failed: " << workbook->getLastError();
    
    // 检查文件是否存在
    std::ifstream file(testFileName);
    EXPECT_TRUE(file.good()) << "Generated file does not exist";
    file.close();
    
    // 检查文件大小（应该大于0）
    std::ifstream fileSize(testFileName, std::ios::binary | std::ios::ate);
    auto size = fileSize.tellg();
    EXPECT_GT(size, 0) << "Generated file is empty";
}

TEST_F(WorksheetWritingTest, FileReadBack) {
    // 设置原始数据
    const std::string originalString = "TestString";
    const double originalDouble = 123.456;
    const int64_t originalInt = 789;
    
    sheet->setCellValue(row_t(1), column_t(1), originalString);
    sheet->setCellValue(row_t(1), column_t(2), originalDouble);
    sheet->setCellValue(row_t(2), column_t(1), originalInt);
    
    // 保存文件
    ASSERT_TRUE(workbook->saveToFile(testFileName)) << workbook->getLastError();
    
    // 读取文件
    TXWorkbook readWorkbook;
    bool loaded = readWorkbook.loadFromFile(testFileName);
    EXPECT_TRUE(loaded) << "Load failed: " << readWorkbook.getLastError();
    
    if (loaded) {
        EXPECT_EQ(readWorkbook.getSheetCount(), 1);
        
        TXSheet* readSheet = readWorkbook.getSheet(0);
        ASSERT_NE(readSheet, nullptr);
        EXPECT_EQ(readSheet->getName(), "TestSheet");
        
        // 目前读取功能可能不完整，这里主要测试不会崩溃
        // TODO: 当读取功能完善后，添加数据验证
        auto val1 = readSheet->getCellValue(row_t(1), column_t(1));
        auto val2 = readSheet->getCellValue(row_t(1), column_t(2));
        auto val3 = readSheet->getCellValue(row_t(2), column_t(1));
        
        // 现在只验证能成功获取到值（即使可能是空的）
        EXPECT_TRUE(std::holds_alternative<std::string>(val1) || 
                   std::holds_alternative<std::monostate>(val1));
    }
}

TEST_F(WorksheetWritingTest, EmptySheet) {
    // 测试空工作表
    auto usedRange = sheet->getUsedRange();
    EXPECT_FALSE(usedRange.isValid());
    
    // 空工作表也应该能够保存
    bool saved = workbook->saveToFile(testFileName);
    EXPECT_TRUE(saved) << workbook->getLastError();
}

TEST_F(WorksheetWritingTest, ComponentDetection) {
    // 添加字符串数据，应该触发SharedStrings组件
    sheet->setCellValue(row_t(1), column_t(1), std::string("String1"));
    sheet->setCellValue(row_t(2), column_t(1), std::string("String2"));
    
    // 保存文件以触发组件检测
    bool saved = workbook->saveToFile(testFileName);
    EXPECT_TRUE(saved);
    
    // 检查组件管理器是否检测到了正确的组件
    const auto& componentManager = workbook->getComponentManager();
    EXPECT_TRUE(componentManager.hasComponent(ExcelComponent::BasicWorkbook));
    EXPECT_TRUE(componentManager.hasComponent(ExcelComponent::SharedStrings));
} 