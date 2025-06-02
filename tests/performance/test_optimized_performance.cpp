#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <chrono>
#include <vector>
#include <random>
#include <iostream>
#include <iomanip>

using namespace TinaXlsx;

class OptimizedPerformanceTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 设置随机数生成器
        gen_.seed(12345); // 固定种子确保可重复性
    }
    
    void TearDown() override {
        // 清理测试文件
        for (const auto& filename : testFiles_) {
            std::remove(filename.c_str());
        }
    }
    
    // 生成随机字符串
    std::string generateRandomString(size_t length) {
        const std::string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        std::uniform_int_distribution<> dis(0, chars.size() - 1);
        std::string result;
        result.reserve(length);
        for (size_t i = 0; i < length; ++i) {
            result += chars[dis(gen_)];
        }
        return result;
    }
    
    // 生成随机数值
    double generateRandomNumber() {
        std::uniform_real_distribution<double> dis(-1000000.0, 1000000.0);
        return dis(gen_);
    }
    
    // 测量执行时间
    template<typename Func>
    std::chrono::microseconds measureTime(Func&& func) {
        auto start = std::chrono::high_resolution_clock::now();
        func();
        auto end = std::chrono::high_resolution_clock::now();
        return std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    }
    
    // 打印性能结果
    void printPerformanceResult(const std::string& testName, 
                               std::chrono::microseconds duration,
                               size_t itemCount,
                               const std::string& unit = "items") {
        double durationMs = duration.count() / 1000.0;
        double itemsPerSecond = itemCount * 1000000.0 / duration.count();
        double timePerItem = duration.count() / static_cast<double>(itemCount);
        
        std::cout << std::fixed << std::setprecision(2);
        std::cout << "[性能] " << testName << ":\n";
        std::cout << "  总时间: " << durationMs << "ms\n";
        std::cout << "  处理量: " << itemCount << " " << unit << "\n";
        std::cout << "  吞吐量: " << itemsPerSecond << " " << unit << "/秒\n";
        std::cout << "  平均时间: " << timePerItem << "μs/" << unit << "\n\n";
    }

protected:
    std::mt19937 gen_;
    std::vector<std::string> testFiles_;
};

// 测试批量操作性能
TEST_F(OptimizedPerformanceTest, BatchOperationsPerformance) {
    TXWorkbook workbook;
    auto sheet = workbook.addSheet("BatchTest");
    
    const size_t rowCount = 1000;
    const size_t colCount = 100;
    
    std::cout << "=== 批量操作性能测试 ===\n\n";
    
    // 测试1：逐个设置 vs 批量设置
    {
        std::vector<std::vector<cell_value_t>> testData;
        testData.reserve(rowCount);
        
        for (size_t r = 0; r < rowCount; ++r) {
            std::vector<cell_value_t> rowData;
            rowData.reserve(colCount);
            for (size_t c = 0; c < colCount; ++c) {
                if (c % 3 == 0) {
                    rowData.emplace_back(generateRandomString(10));
                } else if (c % 3 == 1) {
                    rowData.emplace_back(generateRandomNumber());
                } else {
                    rowData.emplace_back(static_cast<int64_t>(r * colCount + c));
                }
            }
            testData.push_back(std::move(rowData));
        }
        
        // 逐个设置（传统方式）
        auto sheet1 = workbook.addSheet("Individual");
        auto individualTime = measureTime([&]() {
            for (size_t r = 0; r < rowCount; ++r) {
                for (size_t c = 0; c < colCount; ++c) {
                    sheet1->setCellValue(row_t(r + 1), column_t(c + 1), testData[r][c]);
                }
            }
        });
        
        // 批量设置（优化方式）
        auto sheet2 = workbook.addSheet("Batch");
        auto batchTime = measureTime([&]() {
            sheet2->setRangeValues(row_t(1), column_t(1), testData);
        });
        
        size_t totalCells = rowCount * colCount;
        printPerformanceResult("逐个设置单元格", individualTime, totalCells, "cells");
        printPerformanceResult("批量设置单元格", batchTime, totalCells, "cells");
        
        double speedup = static_cast<double>(individualTime.count()) / batchTime.count();
        std::cout << "批量操作加速比: " << std::fixed << std::setprecision(2) << speedup << "x\n\n";
        
        // 验证批量操作的正确性
        EXPECT_EQ(sheet1->getCellValue(row_t(1), column_t(1)), sheet2->getCellValue(row_t(1), column_t(1)));
        EXPECT_EQ(sheet1->getCellValue(row_t(rowCount), column_t(colCount)), 
                  sheet2->getCellValue(row_t(rowCount), column_t(colCount)));
    }
}

// 测试文件保存性能
TEST_F(OptimizedPerformanceTest, FileSavePerformance) {
    std::cout << "=== 文件保存性能测试 ===\n\n";
    
    // 创建不同大小的工作簿进行测试
    std::vector<std::pair<size_t, size_t>> testSizes = {
        {100, 50},    // 5,000 cells
        {500, 100},   // 50,000 cells
        {1000, 200}   // 200,000 cells
    };
    
    for (const auto& [rows, cols] : testSizes) {
        TXWorkbook workbook;
        auto sheet = workbook.addSheet("SaveTest");
        
        // 生成测试数据
        std::vector<std::vector<cell_value_t>> testData;
        testData.reserve(rows);
        
        for (size_t r = 0; r < rows; ++r) {
            std::vector<cell_value_t> rowData;
            rowData.reserve(cols);
            for (size_t c = 0; c < cols; ++c) {
                if (c % 4 == 0) {
                    rowData.emplace_back(generateRandomString(8));
                } else if (c % 4 == 1) {
                    rowData.emplace_back(generateRandomNumber());
                } else if (c % 4 == 2) {
                    rowData.emplace_back(static_cast<int64_t>(r * cols + c));
                } else {
                    rowData.emplace_back(r % 2 == 0);
                }
            }
            testData.push_back(std::move(rowData));
        }
        
        // 使用批量操作填充数据
        sheet->setRangeValues(row_t(1), column_t(1), testData);
        
        // 测试保存性能
        std::string filename = "performance_test_" + std::to_string(rows) + "x" + std::to_string(cols) + ".xlsx";
        testFiles_.push_back(filename);
        
        auto saveTime = measureTime([&]() {
            EXPECT_TRUE(workbook.saveToFile(filename));
        });
        
        size_t totalCells = rows * cols;
        printPerformanceResult("保存文件 (" + std::to_string(rows) + "x" + std::to_string(cols) + ")", 
                              saveTime, totalCells, "cells");
        
        // 检查文件大小
        std::ifstream file(filename, std::ios::binary | std::ios::ate);
        if (file.is_open()) {
            size_t fileSize = file.tellg();
            double fileSizeMB = fileSize / (1024.0 * 1024.0);
            double bytesPerCell = static_cast<double>(fileSize) / totalCells;
            std::cout << "  文件大小: " << std::fixed << std::setprecision(2) << fileSizeMB << " MB\n";
            std::cout << "  每单元格: " << std::fixed << std::setprecision(1) << bytesPerCell << " bytes\n\n";
        }
    }
}

// 测试内存使用效率
TEST_F(OptimizedPerformanceTest, MemoryEfficiencyTest) {
    std::cout << "=== 内存效率测试 ===\n\n";
    
    TXWorkbook workbook;
    auto sheet = workbook.addSheet("MemoryTest");
    
    const size_t batchSize = 10000;
    const size_t numBatches = 10;
    
    // 分批添加数据，观察内存增长
    for (size_t batch = 0; batch < numBatches; ++batch) {
        std::vector<cell_value_t> rowData;
        rowData.reserve(batchSize);
        
        for (size_t i = 0; i < batchSize; ++i) {
            if (i % 3 == 0) {
                rowData.emplace_back(generateRandomString(5));
            } else if (i % 3 == 1) {
                rowData.emplace_back(generateRandomNumber());
            } else {
                rowData.emplace_back(static_cast<int64_t>(batch * batchSize + i));
            }
        }
        
        auto addTime = measureTime([&]() {
            sheet->setRowValues(row_t(batch + 1), column_t(1), rowData);
        });
        
        size_t totalCells = (batch + 1) * batchSize;
        std::cout << "批次 " << (batch + 1) << "/" << numBatches 
                  << ": " << addTime.count() << "μs, 总单元格: " << totalCells << "\n";
    }
    
    std::cout << "\n内存效率测试完成\n\n";
}

// 测试字符串池性能
TEST_F(OptimizedPerformanceTest, StringPoolPerformance) {
    std::cout << "=== 字符串池性能测试 ===\n\n";
    
    TXWorkbook workbook;
    auto sheet = workbook.addSheet("StringTest");
    
    // 生成重复字符串数据
    std::vector<std::string> baseStrings;
    for (int i = 0; i < 100; ++i) {
        baseStrings.push_back("String_" + std::to_string(i));
    }
    
    const size_t rows = 1000;
    const size_t cols = 50;
    
    std::vector<std::vector<cell_value_t>> stringData;
    stringData.reserve(rows);
    
    for (size_t r = 0; r < rows; ++r) {
        std::vector<cell_value_t> rowData;
        rowData.reserve(cols);
        for (size_t c = 0; c < cols; ++c) {
            // 重复使用基础字符串，测试字符串池效果
            rowData.emplace_back(baseStrings[c % baseStrings.size()]);
        }
        stringData.push_back(std::move(rowData));
    }
    
    auto stringTime = measureTime([&]() {
        sheet->setRangeValues(row_t(1), column_t(1), stringData);
    });
    
    size_t totalCells = rows * cols;
    printPerformanceResult("字符串池测试", stringTime, totalCells, "string cells");
    
    // 保存并检查文件大小（字符串池应该显著减少文件大小）
    std::string filename = "string_pool_test.xlsx";
    testFiles_.push_back(filename);
    
    auto saveTime = measureTime([&]() {
        EXPECT_TRUE(workbook.saveToFile(filename));
    });
    
    printPerformanceResult("字符串池文件保存", saveTime, totalCells, "cells");
    
    std::ifstream file(filename, std::ios::binary | std::ios::ate);
    if (file.is_open()) {
        size_t fileSize = file.tellg();
        double fileSizeMB = fileSize / (1024.0 * 1024.0);
        std::cout << "字符串池文件大小: " << std::fixed << std::setprecision(2) << fileSizeMB << " MB\n\n";
    }
}
