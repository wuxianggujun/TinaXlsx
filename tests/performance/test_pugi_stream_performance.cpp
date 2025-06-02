#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <chrono>
#include <vector>
#include <random>
#include <iostream>
#include <iomanip>
#include <fstream>
#include <cstdio>
#include <numeric>
#include <algorithm>

using namespace TinaXlsx;

class PugiStreamPerformanceTest : public ::testing::Test {
protected:
    void SetUp() override {
        gen_.seed(12345);
    }
    
    void TearDown() override {
        for (const auto& filename : testFiles_) {
            std::remove(filename.c_str());
        }
    }
    
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
    
    double generateRandomNumber() {
        std::uniform_real_distribution<double> dis(-1000.0, 1000.0);
        return dis(gen_);
    }
    
    template<typename Func>
    std::chrono::microseconds measureTime(Func&& func) {
        auto start = std::chrono::high_resolution_clock::now();
        func();
        auto end = std::chrono::high_resolution_clock::now();
        return std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    }
    
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

// 测试新的pugixml流式写入器性能
TEST_F(PugiStreamPerformanceTest, PugiStreamWriterPerformance) {
    std::cout << "=== Pugixml流式写入器性能测试 ===\n\n";
    
    // 测试不同数据量的保存性能
    std::vector<std::pair<size_t, size_t>> testSizes = {
        {50, 25},     // 1,250 cells - DOM方式
        {150, 75},    // 11,250 cells - 流式写入器
        {300, 100},   // 30,000 cells - 流式写入器
        {500, 150}    // 75,000 cells - 流式写入器
    };
    
    for (const auto& [rows, cols] : testSizes) {
        std::cout << "--- 测试数据量: " << rows << "x" << cols << " (" << (rows * cols) << " cells) ---\n";
        
        TXWorkbook workbook;
        auto sheet = workbook.addSheet("PugiStreamTest");
        
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
        
        // 使用批量操作填充数据
        auto fillTime = measureTime([&]() {
            sheet->setRangeValues(row_t(1), column_t(1), testData);
        });
        
        size_t totalCells = rows * cols;
        printPerformanceResult("数据填充", fillTime, totalCells, "cells");
        
        // 测试保存性能
        std::string filename = "pugi_stream_test_" + std::to_string(rows) + "x" + std::to_string(cols) + ".xlsx";
        testFiles_.push_back(filename);
        
        auto saveTime = measureTime([&]() {
            EXPECT_TRUE(workbook.saveToFile(filename));
        });
        
        printPerformanceResult("文件保存", saveTime, totalCells, "cells");
        
        // 检查文件大小和完整性
        std::ifstream file(filename, std::ios::binary | std::ios::ate);
        if (file.is_open()) {
            size_t fileSize = file.tellg();
            double fileSizeMB = fileSize / (1024.0 * 1024.0);
            double bytesPerCell = static_cast<double>(fileSize) / totalCells;
            
            std::cout << "文件信息:\n";
            std::cout << "  文件大小: " << std::fixed << std::setprecision(2) << fileSizeMB << " MB\n";
            std::cout << "  每单元格: " << std::fixed << std::setprecision(1) << bytesPerCell << " bytes\n";
            
            // 判断使用的写入策略
            if (totalCells > 5000) {
                std::cout << "  写入策略: Pugixml流式写入器 (高性能)\n";
            } else {
                std::cout << "  写入策略: DOM方式 (兼容性)\n";
            }
            std::cout << "\n";
        }
        
        // 验证文件生成成功
        std::ifstream verifyFile(filename, std::ios::binary);
        if (verifyFile.is_open()) {
            verifyFile.seekg(0, std::ios::end);
            size_t verifySize = verifyFile.tellg();
            if (verifySize > 1000) {
                std::cout << "✅ 文件生成成功，大小: " << verifySize << " bytes\n";
            } else {
                std::cout << "⚠️  文件大小异常: " << verifySize << " bytes\n";
            }
        } else {
            std::cout << "❌ 文件生成失败\n";
        }
        
        std::cout << "\n";
    }
}

// 测试极大数据量性能
TEST_F(PugiStreamPerformanceTest, ExtremeDataPerformance) {
    std::cout << "=== 极大数据量性能测试 ===\n\n";
    
    const size_t rows = 1000;
    const size_t cols = 200;
    const size_t totalCells = rows * cols;
    
    std::cout << "测试数据量: " << rows << "x" << cols << " (" << totalCells << " cells)\n\n";
    
    TXWorkbook workbook;
    auto sheet = workbook.addSheet("ExtremeTest");
    
    // 生成大量测试数据
    std::vector<std::vector<cell_value_t>> testData;
    testData.reserve(rows);
    
    auto dataGenTime = measureTime([&]() {
        for (size_t r = 0; r < rows; ++r) {
            std::vector<cell_value_t> rowData;
            rowData.reserve(cols);
            for (size_t c = 0; c < cols; ++c) {
                if (c % 5 == 0) {
                    rowData.emplace_back(generateRandomString(4));
                } else if (c % 5 == 1) {
                    rowData.emplace_back(generateRandomNumber());
                } else if (c % 5 == 2) {
                    rowData.emplace_back(static_cast<int64_t>(r * cols + c));
                } else if (c % 5 == 3) {
                    rowData.emplace_back(r % 2 == 0);
                } else {
                    rowData.emplace_back(generateRandomNumber() * 0.01); // 小数
                }
            }
            testData.push_back(std::move(rowData));
        }
    });
    
    printPerformanceResult("数据生成", dataGenTime, totalCells, "cells");
    
    // 批量填充数据
    auto fillTime = measureTime([&]() {
        sheet->setRangeValues(row_t(1), column_t(1), testData);
    });
    
    printPerformanceResult("数据填充", fillTime, totalCells, "cells");
    
    // 保存文件
    std::string filename = "extreme_pugi_test.xlsx";
    testFiles_.push_back(filename);
    
    auto saveTime = measureTime([&]() {
        EXPECT_TRUE(workbook.saveToFile(filename));
    });
    
    printPerformanceResult("文件保存", saveTime, totalCells, "cells");
    
    // 文件信息
    std::ifstream file(filename, std::ios::binary | std::ios::ate);
    if (file.is_open()) {
        size_t fileSize = file.tellg();
        double fileSizeMB = fileSize / (1024.0 * 1024.0);
        double bytesPerCell = static_cast<double>(fileSize) / totalCells;
        
        std::cout << "极大数据量文件信息:\n";
        std::cout << "  文件大小: " << std::fixed << std::setprecision(2) << fileSizeMB << " MB\n";
        std::cout << "  每单元格: " << std::fixed << std::setprecision(1) << bytesPerCell << " bytes\n";
        std::cout << "  压缩比: " << std::fixed << std::setprecision(1) 
                  << (static_cast<double>(totalCells * 10) / fileSize) << ":1\n";
        
        std::cout << "\n✅ 极大数据量测试完成\n\n";
    }
}

// 测试性能对比
TEST_F(PugiStreamPerformanceTest, PerformanceComparison) {
    std::cout << "=== 性能对比测试 ===\n\n";
    
    const size_t rows = 200;
    const size_t cols = 100;
    const size_t totalCells = rows * cols;
    const size_t numTests = 3;
    
    std::vector<std::chrono::microseconds> saveTimes;
    saveTimes.reserve(numTests);
    
    for (size_t test = 0; test < numTests; ++test) {
        TXWorkbook workbook;
        auto sheet = workbook.addSheet("ComparisonTest");
        
        // 生成测试数据
        std::vector<std::vector<cell_value_t>> testData;
        testData.reserve(rows);
        
        for (size_t r = 0; r < rows; ++r) {
            std::vector<cell_value_t> rowData;
            rowData.reserve(cols);
            for (size_t c = 0; c < cols; ++c) {
                if (c % 3 == 0) {
                    rowData.emplace_back(generateRandomString(6));
                } else if (c % 3 == 1) {
                    rowData.emplace_back(generateRandomNumber());
                } else {
                    rowData.emplace_back(static_cast<int64_t>(test * 1000000 + r * cols + c));
                }
            }
            testData.push_back(std::move(rowData));
        }
        
        // 填充数据
        sheet->setRangeValues(row_t(1), column_t(1), testData);
        
        // 测试保存性能
        std::string filename = "comparison_test_" + std::to_string(test) + ".xlsx";
        testFiles_.push_back(filename);
        
        auto saveTime = measureTime([&]() {
            EXPECT_TRUE(workbook.saveToFile(filename));
        });
        
        saveTimes.push_back(saveTime);
        
        double timePerCell = static_cast<double>(saveTime.count()) / totalCells;
        std::cout << "测试 " << (test + 1) << "/" << numTests 
                  << ": " << saveTime.count() << "μs"
                  << ", 平均: " << std::fixed << std::setprecision(2) << timePerCell << "μs/cell\n";
    }
    
    // 计算统计信息
    auto minTime = *std::min_element(saveTimes.begin(), saveTimes.end());
    auto maxTime = *std::max_element(saveTimes.begin(), saveTimes.end());
    auto totalTime = std::accumulate(saveTimes.begin(), saveTimes.end(), std::chrono::microseconds(0));
    auto avgTime = totalTime / numTests;
    
    double variation = static_cast<double>(maxTime.count() - minTime.count()) / avgTime.count() * 100;
    
    std::cout << "\n性能统计:\n";
    std::cout << "  最快: " << minTime.count() << "μs\n";
    std::cout << "  最慢: " << maxTime.count() << "μs\n";
    std::cout << "  平均: " << avgTime.count() << "μs\n";
    std::cout << "  变异系数: " << std::fixed << std::setprecision(1) << variation << "%\n";
    
    if (variation > 30.0) {
        std::cout << "⚠️  性能变异较大，可能存在性能波动\n";
    } else {
        std::cout << "✅ 性能稳定性良好\n";
    }
    
    std::cout << "\n✅ 性能对比测试完成\n\n";
}
