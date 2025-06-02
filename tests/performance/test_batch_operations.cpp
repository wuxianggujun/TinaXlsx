#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <chrono>
#include <vector>
#include <random>
#include <iostream>
#include <iomanip>
#include <fstream>
#include <cstdio>

using namespace TinaXlsx;

class BatchOperationsTest : public ::testing::Test {
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
        std::uniform_real_distribution<double> dis(-1000.0, 1000.0);
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

// 测试批量操作 vs 逐个操作的性能差异
TEST_F(BatchOperationsTest, IndividualVsBatchPerformance) {
    std::cout << "=== 批量操作 vs 逐个操作性能对比 ===\n\n";
    
    const size_t rows = 500;
    const size_t cols = 100;
    const size_t totalCells = rows * cols;
    
    // 生成测试数据
    std::vector<std::vector<cell_value_t>> testData;
    testData.reserve(rows);
    
    for (size_t r = 0; r < rows; ++r) {
        std::vector<cell_value_t> rowData;
        rowData.reserve(cols);
        for (size_t c = 0; c < cols; ++c) {
            if (c % 3 == 0) {
                rowData.emplace_back(generateRandomString(8));
            } else if (c % 3 == 1) {
                rowData.emplace_back(generateRandomNumber());
            } else {
                rowData.emplace_back(static_cast<int64_t>(r * cols + c));
            }
        }
        testData.push_back(std::move(rowData));
    }
    
    // 测试1：逐个设置（传统方式）
    TXWorkbook workbook1;
    auto sheet1 = workbook1.addSheet("Individual");
    
    auto individualTime = measureTime([&]() {
        for (size_t r = 0; r < rows; ++r) {
            for (size_t c = 0; c < cols; ++c) {
                sheet1->setCellValue(row_t(r + 1), column_t(c + 1), testData[r][c]);
            }
        }
    });
    
    printPerformanceResult("逐个设置单元格", individualTime, totalCells, "cells");
    
    // 测试2：批量设置（优化方式）
    TXWorkbook workbook2;
    auto sheet2 = workbook2.addSheet("Batch");
    
    auto batchTime = measureTime([&]() {
        sheet2->setRangeValues(row_t(1), column_t(1), testData);
    });
    
    printPerformanceResult("批量设置单元格", batchTime, totalCells, "cells");
    
    // 计算加速比
    double speedup = static_cast<double>(individualTime.count()) / batchTime.count();
    std::cout << "🚀 批量操作加速比: " << std::fixed << std::setprecision(2) << speedup << "x\n\n";
    
    // 验证数据正确性
    EXPECT_EQ(sheet1->getCellValue(row_t(1), column_t(1)), sheet2->getCellValue(row_t(1), column_t(1)));
    EXPECT_EQ(sheet1->getCellValue(row_t(rows), column_t(cols)), 
              sheet2->getCellValue(row_t(rows), column_t(cols)));
    
    std::cout << "✅ 数据正确性验证通过\n\n";
}

// 测试行批量操作性能
TEST_F(BatchOperationsTest, RowBatchPerformance) {
    std::cout << "=== 行批量操作性能测试 ===\n\n";
    
    TXWorkbook workbook;
    auto sheet = workbook.addSheet("RowBatch");
    
    const size_t numRows = 1000;
    const size_t colsPerRow = 50;
    
    // 生成行数据
    std::vector<std::vector<cell_value_t>> rowsData;
    rowsData.reserve(numRows);
    
    for (size_t r = 0; r < numRows; ++r) {
        std::vector<cell_value_t> rowData;
        rowData.reserve(colsPerRow);
        for (size_t c = 0; c < colsPerRow; ++c) {
            if (c % 2 == 0) {
                rowData.emplace_back(generateRandomString(6));
            } else {
                rowData.emplace_back(generateRandomNumber());
            }
        }
        rowsData.push_back(std::move(rowData));
    }
    
    // 测试行批量设置
    auto rowBatchTime = measureTime([&]() {
        for (size_t r = 0; r < numRows; ++r) {
            sheet->setRowValues(row_t(r + 1), column_t(1), rowsData[r]);
        }
    });
    
    size_t totalCells = numRows * colsPerRow;
    printPerformanceResult("行批量操作", rowBatchTime, totalCells, "cells");
    
    // 验证数据
    EXPECT_EQ(sheet->getCellValue(row_t(1), column_t(1)), rowsData[0][0]);
    EXPECT_EQ(sheet->getCellValue(row_t(numRows), column_t(colsPerRow)), 
              rowsData[numRows-1][colsPerRow-1]);
    
    std::cout << "✅ 行批量操作验证通过\n\n";
}

// 测试文件保存性能
TEST_F(BatchOperationsTest, SavePerformanceComparison) {
    std::cout << "=== 文件保存性能对比 ===\n\n";
    
    const size_t rows = 200;
    const size_t cols = 100;
    
    // 生成测试数据
    std::vector<std::vector<cell_value_t>> testData;
    testData.reserve(rows);
    
    for (size_t r = 0; r < rows; ++r) {
        std::vector<cell_value_t> rowData;
        rowData.reserve(cols);
        for (size_t c = 0; c < cols; ++c) {
            if (c % 4 == 0) {
                rowData.emplace_back(generateRandomString(5));
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
    
    // 创建两个相同的工作簿
    TXWorkbook workbook1, workbook2;
    auto sheet1 = workbook1.addSheet("Individual");
    auto sheet2 = workbook2.addSheet("Batch");
    
    // 使用不同方式填充数据
    for (size_t r = 0; r < rows; ++r) {
        for (size_t c = 0; c < cols; ++c) {
            sheet1->setCellValue(row_t(r + 1), column_t(c + 1), testData[r][c]);
        }
    }
    
    sheet2->setRangeValues(row_t(1), column_t(1), testData);
    
    // 测试保存性能
    std::string filename1 = "individual_save_test.xlsx";
    std::string filename2 = "batch_save_test.xlsx";
    testFiles_.push_back(filename1);
    testFiles_.push_back(filename2);
    
    auto saveTime1 = measureTime([&]() {
        EXPECT_TRUE(workbook1.saveToFile(filename1));
    });
    
    auto saveTime2 = measureTime([&]() {
        EXPECT_TRUE(workbook2.saveToFile(filename2));
    });
    
    size_t totalCells = rows * cols;
    printPerformanceResult("逐个操作后保存", saveTime1, totalCells, "cells");
    printPerformanceResult("批量操作后保存", saveTime2, totalCells, "cells");
    
    // 检查文件大小
    std::ifstream file1(filename1, std::ios::binary | std::ios::ate);
    std::ifstream file2(filename2, std::ios::binary | std::ios::ate);
    
    if (file1.is_open() && file2.is_open()) {
        size_t size1 = file1.tellg();
        size_t size2 = file2.tellg();
        
        std::cout << "文件大小对比:\n";
        std::cout << "  逐个操作: " << size1 << " bytes\n";
        std::cout << "  批量操作: " << size2 << " bytes\n";
        std::cout << "  大小差异: " << static_cast<int>(size2) - static_cast<int>(size1) << " bytes\n\n";
        
        // 文件大小应该基本相同（允许小幅差异）
        EXPECT_NEAR(size1, size2, size1 * 0.01); // 允许1%的差异
    }
    
    std::cout << "✅ 文件保存测试完成\n\n";
}

// 测试内存使用效率
TEST_F(BatchOperationsTest, MemoryEfficiencyTest) {
    std::cout << "=== 内存使用效率测试 ===\n\n";
    
    TXWorkbook workbook;
    auto sheet = workbook.addSheet("MemoryTest");
    
    const size_t batchSize = 5000;
    const size_t numBatches = 5;
    
    std::cout << "分批添加数据，观察性能稳定性:\n";
    
    for (size_t batch = 0; batch < numBatches; ++batch) {
        std::vector<cell_value_t> rowData;
        rowData.reserve(batchSize);
        
        for (size_t i = 0; i < batchSize; ++i) {
            if (i % 3 == 0) {
                rowData.emplace_back(generateRandomString(4));
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
        double timePerCell = static_cast<double>(addTime.count()) / batchSize;
        
        std::cout << "批次 " << (batch + 1) << "/" << numBatches 
                  << ": " << addTime.count() << "μs"
                  << ", 平均: " << std::fixed << std::setprecision(2) << timePerCell << "μs/cell"
                  << ", 总单元格: " << totalCells << "\n";
    }
    
    std::cout << "\n✅ 内存效率测试完成\n\n";
}
