//
// Performance Test
// 测试性能相关功能
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include <filesystem>
#include <iostream>
#include <chrono>

class PerformanceTest : public ::testing::Test {
protected:
    void SetUp() override {
        std::filesystem::create_directories("output");
    }

    void TearDown() override {
        std::filesystem::remove("output/performance_test.xlsx");
    }
};

/**
 * @brief 测试大量数据写入性能
 */
TEST_F(PerformanceTest, LargeDataWritePerformance) {
    std::cout << "\n=== 大量数据写入性能测试 ===\n";
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("性能测试");
    ASSERT_NE(sheet, nullptr);
    
    // 写入10000行数据
    const int rowCount = 10000;
    const int colCount = 5;
    
    std::cout << "开始写入 " << rowCount << " 行 × " << colCount << " 列数据...\n";
    
    for (int row = 1; row <= rowCount; ++row) {
        for (int col = 1; col <= colCount; ++col) {
            std::string value;
            switch (col) {
                case 1: value = "Item_" + std::to_string(row); break;
                case 2: value = std::to_string(row * 10); break;
                case 3: value = std::to_string(row * 0.5); break;
                case 4: value = std::to_string(row * 15.5); break;
                case 5: value = "Category_" + std::to_string(row % 10); break;
            }
            
            bool success = sheet->setCellValue(TinaXlsx::row_t(row), TinaXlsx::column_t(col), value);
            EXPECT_TRUE(success);
        }
        
        // 每1000行显示进度
        if (row % 1000 == 0) {
            std::cout << "已写入 " << row << " 行数据\n";
        }
    }    
    auto write_time = std::chrono::high_resolution_clock::now();
    auto write_duration = std::chrono::duration_cast<std::chrono::milliseconds>(write_time - start_time);
    std::cout << "数据写入耗时: " << write_duration.count() << " ms\n";
    
    // 保存文件
    std::cout << "开始保存文件...\n";
    bool saved = workbook.saveToFile("output/performance_test.xlsx");
    EXPECT_TRUE(saved) << "保存失败: " << workbook.getLastError();
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
    auto save_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - write_time);
    
    std::cout << "文件保存耗时: " << save_duration.count() << " ms\n";
    std::cout << "总耗时: " << total_duration.count() << " ms\n";
    
    // 验证文件大小
    auto file_size = std::filesystem::file_size("output/performance_test.xlsx");
    std::cout << "生成文件大小: " << file_size << " bytes\n";
    
    std::cout << "大量数据写入性能测试通过！\n";
}

/**
 * @brief 测试批量操作性能
 */
TEST_F(PerformanceTest, BatchOperationPerformance) {
    std::cout << "\n=== 批量操作性能测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("批量操作");
    ASSERT_NE(sheet, nullptr);
    
    const int dataCount = 5000;
    
    // 准备批量数据
    std::vector<std::pair<TinaXlsx::TXCoordinate, TinaXlsx::TXSheet::CellValue>> batch_data;
    batch_data.reserve(dataCount);
    
    auto prep_start = std::chrono::high_resolution_clock::now();
    
    for (int i = 1; i <= dataCount; ++i) {
        batch_data.emplace_back(
            TinaXlsx::TXCoordinate(TinaXlsx::row_t(i), TinaXlsx::column_t(1)),
            std::string("批量数据_") + std::to_string(i)
        );
    }
    
    auto prep_end = std::chrono::high_resolution_clock::now();
    auto prep_duration = std::chrono::duration_cast<std::chrono::milliseconds>(prep_end - prep_start);
    std::cout << "数据准备耗时: " << prep_duration.count() << " ms\n";
    
    // 执行批量写入
    auto batch_start = std::chrono::high_resolution_clock::now();
    size_t success_count = sheet->setCellValues(batch_data);
    auto batch_end = std::chrono::high_resolution_clock::now();
    
    auto batch_duration = std::chrono::duration_cast<std::chrono::milliseconds>(batch_end - batch_start);
    std::cout << "批量写入耗时: " << batch_duration.count() << " ms\n";
    std::cout << "成功写入: " << success_count << " / " << dataCount << " 个单元格\n";
    
    EXPECT_EQ(dataCount, success_count);
    
    // 保存文件
    bool saved = workbook.saveToFile("output/performance_test.xlsx");
    EXPECT_TRUE(saved);
    
    std::cout << "批量操作性能测试通过！\n";
}/**
 * @brief 测试合并单元格性能
 */
TEST_F(PerformanceTest, MergedCellsPerformance) {
    std::cout << "\n=== 合并单元格性能测试 ===\n";
    
    TinaXlsx::TXWorkbook workbook;
    auto* sheet = workbook.addSheet("合并性能");
    ASSERT_NE(sheet, nullptr);
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 创建100个合并区域
    const int mergeCount = 100;
    int successCount = 0;
    
    for (int i = 0; i < mergeCount; ++i) {
        int startRow = i * 3 + 1;
        int endRow = startRow + 1;
        int startCol = (i % 10) + 1;
        int endCol = startCol + 1;
        
        bool success = sheet->mergeCells(
            TinaXlsx::row_t(startRow), TinaXlsx::column_t(startCol),
            TinaXlsx::row_t(endRow), TinaXlsx::column_t(endCol)
        );
        
        if (success) {
            successCount++;
            // 在合并区域设置数据
            sheet->setCellValue(TinaXlsx::row_t(startRow), TinaXlsx::column_t(startCol), 
                               std::string("合并_") + std::to_string(i));
        }
    }
    
    auto merge_time = std::chrono::high_resolution_clock::now();
    auto merge_duration = std::chrono::duration_cast<std::chrono::milliseconds>(merge_time - start_time);
    
    std::cout << "合并操作耗时: " << merge_duration.count() << " ms\n";
    std::cout << "成功合并: " << successCount << " / " << mergeCount << " 个区域\n";
    std::cout << "总合并区域数: " << sheet->getMergeCount() << "\n";
    
    EXPECT_GT(successCount, mergeCount * 0.8); // 至少80%成功
    
    // 保存文件
    bool saved = workbook.saveToFile("output/performance_test.xlsx");
    EXPECT_TRUE(saved);
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
    std::cout << "总耗时: " << total_duration.count() << " ms\n";
    
    std::cout << "合并单元格性能测试通过！\n";
}/**
 * @brief 测试多工作表性能
 */
TEST_F(PerformanceTest, MultiSheetPerformance) {
    std::cout << "\n=== 多工作表性能测试 ===\n";
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    TinaXlsx::TXWorkbook workbook;
    
    const int sheetCount = 10;
    const int rowsPerSheet = 1000;
    
    for (int s = 1; s <= sheetCount; ++s) {
        std::string sheetName = "Sheet_" + std::to_string(s);
        auto* sheet = workbook.addSheet(sheetName);
        ASSERT_NE(sheet, nullptr);
        
        // 在每个工作表中写入数据
        for (int r = 1; r <= rowsPerSheet; ++r) {
            sheet->setCellValue(TinaXlsx::row_t(r), TinaXlsx::column_t(1), 
                               std::string("数据_") + std::to_string(r));
            sheet->setCellValue(TinaXlsx::row_t(r), TinaXlsx::column_t(2), 
                               static_cast<double>(r * s));
        }
        
        std::cout << "完成工作表 " << s << " / " << sheetCount << "\n";
    }
    
    auto data_time = std::chrono::high_resolution_clock::now();
    auto data_duration = std::chrono::duration_cast<std::chrono::milliseconds>(data_time - start_time);
    std::cout << "数据写入耗时: " << data_duration.count() << " ms\n";
    
    // 验证工作表数量
    EXPECT_EQ(sheetCount, workbook.getSheetCount());
    
    // 保存文件
    bool saved = workbook.saveToFile("output/performance_test.xlsx");
    EXPECT_TRUE(saved);
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
    auto save_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - data_time);
    
    std::cout << "文件保存耗时: " << save_duration.count() << " ms\n";
    std::cout << "总耗时: " << total_duration.count() << " ms\n";
    
    std::cout << "多工作表性能测试通过！\n";
}