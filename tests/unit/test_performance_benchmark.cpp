#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "test_file_generator.hpp"
#include <chrono>
#include <cstdio>
#include <filesystem>
#include <iostream>
#include <memory>

using namespace TinaXlsx;

class PerformanceBenchmarkTest : public TestWithFileGeneration<PerformanceBenchmarkTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<PerformanceBenchmarkTest>::SetUp();

        // 记录测试开始时间
        start_time = std::chrono::high_resolution_clock::now();
    }

    void TearDown() override {
        TestWithFileGeneration<PerformanceBenchmarkTest>::TearDown();
    }

    // 时间测量辅助方法
    template<typename Func>
    double measureExecutionTime(Func&& func) {
        auto start = std::chrono::high_resolution_clock::now();
        func();
        auto end = std::chrono::high_resolution_clock::now();
        
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
        return duration.count() / 1000.0; // 返回毫秒
    }

    void printPerformanceReport(const std::string& test_name, double time_ms, 
                               int operations = 1, const std::string& extra_info = "") {
        std::cout << "\n=== " << test_name << " 性能报告 ===" << std::endl;
        std::cout << "总时间: " << time_ms << " ms" << std::endl;
        if (operations > 1) {
            std::cout << "操作数量: " << operations << std::endl;
            std::cout << "平均每操作: " << (time_ms / operations) << " ms" << std::endl;
        }
        if (!extra_info.empty()) {
            std::cout << "额外信息: " << extra_info << std::endl;
        }
        std::cout << "=========================" << std::endl;
    }

    std::chrono::high_resolution_clock::time_point start_time;
};

// 测试单元格写入性能
TEST_F(PerformanceBenchmarkTest, CellWritingPerformance) {
    const int ROWS = 1000;
    const int COLS = 10;

    auto workbook = createWorkbook("CellWritingBenchmark");
    auto* sheet = workbook->addSheet("性能测试");

    // 添加测试信息
    addTestInfo(sheet, "CellWritingPerformance", "单元格写入性能测试 - " + std::to_string(ROWS) + "x" + std::to_string(COLS));

    double time_ms = measureExecutionTime([&]() {
        // 写入大量单元格数据
        for (int row = 1; row <= ROWS; ++row) {
            for (int col = 1; col <= COLS; ++col) {
                if (col == 1) {
                    sheet->setCellValue(row_t(row), column_t(col), "行_" + std::to_string(row));
                } else if (col == 2) {
                    sheet->setCellValue(row_t(row), column_t(col), row * col * 1.5);
                } else if (col == 3) {
                    sheet->setCellValue(row_t(row), column_t(col), row % 2 == 0);
                } else {
                    sheet->setCellValue(row_t(row), column_t(col), static_cast<double>(row + col));
                }
            }
        }
    });

    bool saved = saveWorkbook(workbook, "CellWritingBenchmark");
    EXPECT_TRUE(saved);
    
    int total_operations = ROWS * COLS;
    std::string filePath = getFilePath("CellWritingBenchmark");
    auto file_size = std::filesystem::file_size(filePath);
    std::string extra_info = "文件大小: " + std::to_string(file_size) + " bytes";

    printPerformanceReport("单元格写入", time_ms, total_operations, extra_info);

    // 性能断言 - 平均每个单元格操作应该在合理时间内完成
    double avg_time_per_cell = time_ms / total_operations;
    EXPECT_LT(avg_time_per_cell, 1.0); // 每个单元格操作应少于1ms
}

// 测试数字格式化性能
TEST_F(PerformanceBenchmarkTest, NumberFormattingPerformance) {
    std::string output_file = benchmark_dir + "/formatting_benchmark.xlsx";
    
    const int ROWS = 500;
    
    double time_ms = measureExecutionTime([&]() {
        auto workbook = std::make_unique<TXWorkbook>();
        auto* sheet = workbook->addSheet("格式化测试");
        
        // 测试各种数字格式化
        for (int row = 1; row <= ROWS; ++row) {
            double value = row * 123.456;
            
            // 基本数字格式
            sheet->setCellValue(row_t(row), column_t(1), value);
            sheet->setCellNumberFormat(row_t(row), column_t(1), TXNumberFormat::FormatType::Number, 2);
            
            // 百分比格式
            sheet->setCellValue(row_t(row), column_t(2), value / 100.0);
            sheet->setCellNumberFormat(row_t(row), column_t(2), TXNumberFormat::FormatType::Percentage, 1);
            
            // 货币格式
            sheet->setCellValue(row_t(row), column_t(3), value);
            sheet->setCellNumberFormat(row_t(row), column_t(3), TXNumberFormat::FormatType::Currency, 2);
            
            // 自定义格式
            sheet->setCellValue(row_t(row), column_t(4), value);
            sheet->setCellCustomFormat(row_t(row), column_t(4), "#,##0.00_);[红色](#,##0.00)");
        }
        
        workbook->saveToFile(output_file);
    });
    
    EXPECT_TRUE(std::filesystem::exists(output_file));
    
    int total_operations = ROWS * 4; // 4种格式
    printPerformanceReport("数字格式化", time_ms, total_operations);
    
    // 性能断言
    double avg_time_per_format = time_ms / total_operations;
    EXPECT_LT(avg_time_per_format, 2.0); // 每次格式化应少于2ms
}// 测试多工作表创建性能
TEST_F(PerformanceBenchmarkTest, MultiSheetCreationPerformance) {
    std::string output_file = benchmark_dir + "/multi_sheet_benchmark.xlsx";
    
    const int SHEET_COUNT = 50;
    const int ROWS_PER_SHEET = 100;
    
    double time_ms = measureExecutionTime([&]() {
        auto workbook = std::make_unique<TXWorkbook>();
        
        for (int sheet_idx = 1; sheet_idx <= SHEET_COUNT; ++sheet_idx) {
            std::string sheet_name = "工作表_" + std::to_string(sheet_idx);
            auto* sheet = workbook->addSheet(sheet_name);
            
            // 在每个工作表中添加一些数据
            for (int row = 1; row <= ROWS_PER_SHEET; ++row) {
                sheet->setCellValue(row_t(row), column_t(1), "数据_" + std::to_string(row));
                sheet->setCellValue(row_t(row), column_t(2), sheet_idx * row);
            }
        }
        
        workbook->saveToFile(output_file);
    });
    
    EXPECT_TRUE(std::filesystem::exists(output_file));
    
    int total_sheets = SHEET_COUNT;
    int total_cells = SHEET_COUNT * ROWS_PER_SHEET * 2; // 每行2个单元格
    
    std::string extra_info = "工作表数: " + std::to_string(total_sheets) + 
                           ", 总单元格数: " + std::to_string(total_cells);
    
    printPerformanceReport("多工作表创建", time_ms, total_sheets, extra_info);
    
    // 性能断言
    double avg_time_per_sheet = time_ms / total_sheets;
    EXPECT_LT(avg_time_per_sheet, 50.0); // 每个工作表创建应少于50ms
}

// 测试大量字符串写入性能
TEST_F(PerformanceBenchmarkTest, StringWritingPerformance) {
    std::string output_file = benchmark_dir + "/string_benchmark.xlsx";
    
    const int STRING_COUNT = 2000;
    
    double time_ms = measureExecutionTime([&]() {
        auto workbook = std::make_unique<TXWorkbook>();
        auto* sheet = workbook->addSheet("字符串测试");
        
        for (int i = 1; i <= STRING_COUNT; ++i) {
            std::string long_string = "这是一个很长的字符串测试内容，包含中文和英文 English content " + 
                                    std::to_string(i) + " 用于测试字符串处理性能。";
            
            sheet->setCellValue(row_t(i), column_t(1), long_string);
            sheet->setCellValue(row_t(i), column_t(2), "简短文本_" + std::to_string(i));
        }
        
        workbook->saveToFile(output_file);
    });
    
    EXPECT_TRUE(std::filesystem::exists(output_file));
    
    printPerformanceReport("字符串写入", time_ms, STRING_COUNT);
    
    // 性能断言
    double avg_time_per_string = time_ms / STRING_COUNT;
    EXPECT_LT(avg_time_per_string, 1.5); // 每个字符串操作应少于1.5ms
}

// 测试公式写入性能
TEST_F(PerformanceBenchmarkTest, FormulaWritingPerformance) {
    std::string output_file = benchmark_dir + "/formula_benchmark.xlsx";
    
    const int FORMULA_COUNT = 300;
    
    double time_ms = measureExecutionTime([&]() {
        auto workbook = std::make_unique<TXWorkbook>();
        auto* sheet = workbook->addSheet("公式测试");
        
        // 先写入一些基础数据
        for (int i = 1; i <= 100; ++i) {
            sheet->setCellValue(row_t(i), column_t(1), i * 10.0);
            sheet->setCellValue(row_t(i), column_t(2), i * 5.0);
        }
        
        // 写入各种公式
        for (int i = 1; i <= FORMULA_COUNT; ++i) {
            if (i <= 100) {
                // 简单算术公式
                sheet->setCellValue(row_t(i), column_t(3), "=A" + std::to_string(i) + "+B" + std::to_string(i));
            } else if (i <= 200) {
                // SUM公式
                sheet->setCellValue(row_t(i), column_t(3), "=SUM(A1:A" + std::to_string(i-100) + ")");
            } else {
                // 复合公式
                sheet->setCellValue(row_t(i), column_t(3), "=IF(A" + std::to_string(i-200) + ">50,A" + 
                                  std::to_string(i-200) + "*2,A" + std::to_string(i-200) + "/2)");
            }
        }
        
        workbook->saveToFile(output_file);
    });
    
    EXPECT_TRUE(std::filesystem::exists(output_file));
    
    printPerformanceReport("公式写入", time_ms, FORMULA_COUNT);
    
    // 性能断言
    double avg_time_per_formula = time_ms / FORMULA_COUNT;
    EXPECT_LT(avg_time_per_formula, 2.0); // 每个公式操作应少于2ms
}

// 测试内存使用性能
TEST_F(PerformanceBenchmarkTest, MemoryUsagePerformance) {
    std::string output_file = benchmark_dir + "/memory_benchmark.xlsx";
    
    const int LARGE_DATASET_SIZE = 5000;
    
    double time_ms = measureExecutionTime([&]() {
        auto workbook = std::make_unique<TXWorkbook>();
        auto* sheet = workbook->addSheet("内存测试");
        
        // 创建大量数据测试内存使用
        for (int row = 1; row <= LARGE_DATASET_SIZE; ++row) {
            sheet->setCellValue(row_t(row), column_t(1), "数据行_" + std::to_string(row));
            sheet->setCellValue(row_t(row), column_t(2), row * 3.14159);
            sheet->setCellValue(row_t(row), column_t(3), row % 2 == 0);
            sheet->setCellValue(row_t(row), column_t(4), "=B" + std::to_string(row) + "*2");
            
            // 添加格式化
            if (row % 100 == 0) {
                sheet->setCellNumberFormat(row_t(row), column_t(2), TXNumberFormat::FormatType::Number, 3);
            }
        }
        
        workbook->saveToFile(output_file);
    });
    
    EXPECT_TRUE(std::filesystem::exists(output_file));
    
    auto file_size = std::filesystem::file_size(output_file);
    std::string extra_info = "数据集大小: " + std::to_string(LARGE_DATASET_SIZE) + 
                           " 行, 文件大小: " + std::to_string(file_size) + " bytes";
    
    printPerformanceReport("大数据集处理", time_ms, LARGE_DATASET_SIZE, extra_info);
    
    // 性能断言
    EXPECT_LT(time_ms, 10000.0); // 整个操作应在10秒内完成
    EXPECT_GT(file_size, 50000); // 文件应有合理大小
}

// 测试文件保存性能
TEST_F(PerformanceBenchmarkTest, FileSavePerformance) {
    const int SAVE_COUNT = 20;
    
    double time_ms = measureExecutionTime([&]() {
        for (int i = 1; i <= SAVE_COUNT; ++i) {
            auto workbook = std::make_unique<TXWorkbook>();
            auto* sheet = workbook->addSheet("保存测试_" + std::to_string(i));
            
            // 添加一些数据
            for (int row = 1; row <= 50; ++row) {
                sheet->setCellValue(row_t(row), column_t(1), "测试数据_" + std::to_string(row));
                sheet->setCellValue(row_t(row), column_t(2), row * i);
            }
            
            std::string output_file = benchmark_dir + "/save_test_" + std::to_string(i) + ".xlsx";
            workbook->saveToFile(output_file);
            
            EXPECT_TRUE(std::filesystem::exists(output_file));
        }
    });
    
    printPerformanceReport("文件保存", time_ms, SAVE_COUNT);
    
    // 性能断言
    double avg_time_per_save = time_ms / SAVE_COUNT;
    EXPECT_LT(avg_time_per_save, 200.0); // 每次保存应少于200ms
}
