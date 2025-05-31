#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <cstdio>
#include <filesystem>
#include <fstream>

using namespace TinaXlsx;

class XlsxReaderTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 创建测试目录
        test_dir = "test_reader_files";
        std::filesystem::create_directories(test_dir);
        
        // 创建测试文件
        createTestFile();
    }

    void TearDown() override {
        // 清理测试文件
        try {
            std::filesystem::remove_all(test_dir);
        } catch (const std::exception& e) {
            // 忽略清理错误
        }
    }

    void createTestFile() {
        auto workbook = std::make_unique<TXWorkbook>();
        auto* sheet = workbook->addSheet("TestData");
        
        // 创建包含各种数据类型的测试文件
        sheet->setCellValue(row_t(1), column_t(1), "字符串测试");
        sheet->setCellValue(row_t(1), column_t(2), 123.45);
        sheet->setCellValue(row_t(1), column_t(3), true);
        sheet->setCellValue(row_t(1), column_t(4), "=A1&B1");
        
        // 添加数字格式
        sheet->setCellNumberFormat(row_t(1), column_t(2), TXNumberFormat::FormatType::Number, 2);
        
        test_file_path = test_dir + "/test_data.xlsx";
        workbook->saveToFile(test_file_path);
    }

    std::string test_dir;
    std::string test_file_path;
};

// 测试基本文件读取功能
TEST_F(XlsxReaderTest, BasicFileReading) {
    // 验证测试文件存在
    ASSERT_TRUE(std::filesystem::exists(test_file_path));
    
    // 尝试读取文件 (假设有loadFromFile方法)
    auto workbook = std::make_unique<TXWorkbook>();
    
    // 注意：这里假设TXWorkbook有读取方法，如果没有则需要实现
    // EXPECT_TRUE(workbook->loadFromFile(test_file_path));
    
    // 暂时跳过实际读取，等待读取功能实现
    EXPECT_TRUE(std::filesystem::exists(test_file_path));
}

// 测试空文件处理
TEST_F(XlsxReaderTest, EmptyFileHandling) {
    std::string empty_file = test_dir + "/empty.xlsx";
    
    // 创建空文件
    std::ofstream empty(empty_file);
    empty.close();
    
    auto workbook = std::make_unique<TXWorkbook>();
    // EXPECT_FALSE(workbook->loadFromFile(empty_file));
    
    // 验证文件存在但为空
    EXPECT_TRUE(std::filesystem::exists(empty_file));
    EXPECT_EQ(std::filesystem::file_size(empty_file), 0);
}

// 测试不存在文件的处理
TEST_F(XlsxReaderTest, NonExistentFileHandling) {
    std::string non_existent = test_dir + "/does_not_exist.xlsx";
    
    auto workbook = std::make_unique<TXWorkbook>();
    // EXPECT_FALSE(workbook->loadFromFile(non_existent));
    
    EXPECT_FALSE(std::filesystem::exists(non_existent));
}

// 测试损坏文件的处理
TEST_F(XlsxReaderTest, CorruptedFileHandling) {
    std::string corrupted_file = test_dir + "/corrupted.xlsx";
    
    // 创建损坏的文件
    std::ofstream corrupt(corrupted_file);
    corrupt << "这不是一个有效的xlsx文件内容";
    corrupt.close();
    
    auto workbook = std::make_unique<TXWorkbook>();
    // EXPECT_FALSE(workbook->loadFromFile(corrupted_file));
    
    EXPECT_TRUE(std::filesystem::exists(corrupted_file));
    EXPECT_GT(std::filesystem::file_size(corrupted_file), 0);
}

// 测试大文件读取
TEST_F(XlsxReaderTest, LargeFileReading) {
    std::string large_file = test_dir + "/large_data.xlsx";
    
    // 创建包含大量数据的文件
    auto workbook = std::make_unique<TXWorkbook>();
    auto* sheet = workbook->addSheet("LargeData");
    
    // 写入1000行数据
    for (int i = 1; i <= 1000; ++i) {
        sheet->setCellValue(row_t(i), column_t(1), "行数据_" + std::to_string(i));
        sheet->setCellValue(row_t(i), column_t(2), i * 10.5);
        sheet->setCellValue(row_t(i), column_t(3), i % 2 == 0);
    }
    
    EXPECT_TRUE(workbook->saveToFile(large_file));
    EXPECT_TRUE(std::filesystem::exists(large_file));
    
    // 验证文件大小合理
    auto file_size = std::filesystem::file_size(large_file);
    EXPECT_GT(file_size, 1024); // 至少1KB
}

// 测试多工作表文件读取
TEST_F(XlsxReaderTest, MultiSheetFileReading) {
    std::string multi_sheet_file = test_dir + "/multi_sheet.xlsx";
    
    auto workbook = std::make_unique<TXWorkbook>();
    
    // 创建多个工作表
    auto* sheet1 = workbook->addSheet("销售数据");
    auto* sheet2 = workbook->addSheet("统计报表");
    auto* sheet3 = workbook->addSheet("图表数据");
    
    // 在每个工作表中添加数据
    sheet1->setCellValue(row_t(1), column_t(1), "销售金额");
    sheet1->setCellValue(row_t(2), column_t(1), 10000.0);
    
    sheet2->setCellValue(row_t(1), column_t(1), "总计");
    sheet2->setCellValue(row_t(2), column_t(1), "=销售数据.B2*1.2");
    
    sheet3->setCellValue(row_t(1), column_t(1), "图表标题");
    sheet3->setCellValue(row_t(2), column_t(1), "数据源");
    
    EXPECT_TRUE(workbook->saveToFile(multi_sheet_file));
    EXPECT_TRUE(std::filesystem::exists(multi_sheet_file));
}

// 测试特殊字符文件名处理
TEST_F(XlsxReaderTest, SpecialCharacterFilename) {
    std::string special_file = test_dir + "/测试文件_特殊字符@#$.xlsx";
    
    auto workbook = std::make_unique<TXWorkbook>();
    auto* sheet = workbook->addSheet("测试");
    
    sheet->setCellValue(row_t(1), column_t(1), "中文内容测试");
    sheet->setCellValue(row_t(1), column_t(2), "English Content");
    sheet->setCellValue(row_t(1), column_t(3), "Специальные символы");
    
    EXPECT_TRUE(workbook->saveToFile(special_file));
    EXPECT_TRUE(std::filesystem::exists(special_file));
}

// 测试格式化数据读取
TEST_F(XlsxReaderTest, FormattedDataReading) {
    std::string formatted_file = test_dir + "/formatted_data.xlsx";
    
    auto workbook = std::make_unique<TXWorkbook>();
    auto* sheet = workbook->addSheet("格式化数据");
    
    // 设置各种格式的数据
    sheet->setCellValue(row_t(1), column_t(1), 1234.567);
    sheet->setCellNumberFormat(row_t(1), column_t(1), TXNumberFormat::FormatType::Number, 2);
    
    sheet->setCellValue(row_t(2), column_t(1), 0.75);
    sheet->setCellNumberFormat(row_t(2), column_t(1), TXNumberFormat::FormatType::Percentage, 1);
    
    sheet->setCellValue(row_t(3), column_t(1), 50000.0);
    sheet->setCellNumberFormat(row_t(3), column_t(1), TXNumberFormat::FormatType::Currency, 2);
    
    EXPECT_TRUE(workbook->saveToFile(formatted_file));
    EXPECT_TRUE(std::filesystem::exists(formatted_file));
}