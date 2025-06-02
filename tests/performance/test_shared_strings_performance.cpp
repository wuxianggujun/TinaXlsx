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
#include <set>

using namespace TinaXlsx;

class SharedStringsPerformanceTest : public ::testing::Test {
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
        const std::string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789 ";
        std::uniform_int_distribution<> dis(0, chars.size() - 1);
        std::string result;
        result.reserve(length);
        for (size_t i = 0; i < length; ++i) {
            result += chars[dis(gen_)];
        }
        return result;
    }
    
    // 生成具有一定重复率的字符串集合
    std::vector<std::string> generateStringSet(size_t count, double duplicateRate = 0.3) {
        std::vector<std::string> strings;
        strings.reserve(count);
        
        // 生成基础字符串集合
        size_t uniqueCount = static_cast<size_t>(count * (1.0 - duplicateRate));
        std::vector<std::string> uniqueStrings;
        uniqueStrings.reserve(uniqueCount);
        
        for (size_t i = 0; i < uniqueCount; ++i) {
            size_t length = 5 + (i % 20); // 5-24字符长度
            uniqueStrings.push_back(generateRandomString(length));
        }
        
        // 生成最终字符串集合（包含重复）
        std::uniform_int_distribution<> indexDis(0, uniqueStrings.size() - 1);
        for (size_t i = 0; i < count; ++i) {
            strings.push_back(uniqueStrings[indexDis(gen_)]);
        }
        
        return strings;
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

// 测试SharedStrings流式写入器性能
TEST_F(SharedStringsPerformanceTest, SharedStringsPerformanceTest) {
    std::cout << "=== SharedStrings流式写入器性能测试 ===\n\n";
    
    // 测试不同字符串数量的性能表现
    std::vector<size_t> testCounts = {
        100,    // 小数据量
        500,    // 中小数据量
        1000,   // 中等数据量
        2000,   // 中大数据量
        5000,   // 大数据量
        10000   // 超大数据量
    };
    
    for (size_t stringCount : testCounts) {
        std::cout << "--- 测试字符串数量: " << stringCount << " ---\n";
        
        TXWorkbook workbook;
        auto sheet = workbook.addSheet("StringTest");
        
        // 生成测试字符串
        auto testStrings = generateStringSet(stringCount, 0.4); // 40%重复率
        
        // 创建包含大量字符串的数据
        size_t rows = std::min(stringCount, static_cast<size_t>(1000));
        size_t cols = (stringCount + rows - 1) / rows; // 向上取整
        
        std::vector<std::vector<cell_value_t>> testData;
        testData.reserve(rows);
        
        size_t stringIndex = 0;
        for (size_t r = 0; r < rows && stringIndex < stringCount; ++r) {
            std::vector<cell_value_t> rowData;
            rowData.reserve(cols);
            for (size_t c = 0; c < cols && stringIndex < stringCount; ++c) {
                rowData.emplace_back(testStrings[stringIndex++]);
            }
            testData.push_back(std::move(rowData));
        }
        
        // 填充数据
        auto fillTime = measureTime([&]() {
            sheet->setRangeValues(row_t(1), column_t(1), testData);
        });
        
        printPerformanceResult("数据填充", fillTime, stringCount, "strings");
        
        // 保存文件
        std::string filename = "shared_strings_test_" + std::to_string(stringCount) + ".xlsx";
        testFiles_.push_back(filename);
        
        auto saveTime = measureTime([&]() {
            EXPECT_TRUE(workbook.saveToFile(filename));
        });
        
        printPerformanceResult("文件保存", saveTime, stringCount, "strings");
        
        // 检查文件信息
        std::ifstream file(filename, std::ios::binary | std::ios::ate);
        if (file.is_open()) {
            size_t fileSize = file.tellg();
            double fileSizeMB = fileSize / (1024.0 * 1024.0);
            
            std::cout << "文件信息:\n";
            std::cout << "  文件大小: " << std::fixed << std::setprecision(2) << fileSizeMB << " MB\n";
            std::cout << "  写入策略: SharedStrings流式写入器 (高性能)\n";
            std::cout << "\n";
        }
        
        // 计算去重后的字符串数量
        std::set<std::string> uniqueStrings(testStrings.begin(), testStrings.end());
        std::cout << "字符串统计:\n";
        std::cout << "  总字符串: " << stringCount << "\n";
        std::cout << "  唯一字符串: " << uniqueStrings.size() << "\n";
        std::cout << "  重复率: " << std::fixed << std::setprecision(1) 
                  << (1.0 - static_cast<double>(uniqueStrings.size()) / stringCount) * 100 << "%\n\n";
    }
}

// 测试大量字符串性能
TEST_F(SharedStringsPerformanceTest, LargeStringSetPerformance) {
    std::cout << "=== 大量字符串性能测试 ===\n\n";
    
    const size_t stringCount = 20000;
    std::cout << "测试字符串数量: " << stringCount << "\n\n";
    
    TXWorkbook workbook;
    auto sheet = workbook.addSheet("LargeStringTest");
    
    // 生成大量测试字符串
    auto testStrings = generateStringSet(stringCount, 0.5); // 50%重复率
    
    auto stringGenTime = measureTime([&]() {
        // 字符串生成时间已包含在generateStringSet中
    });
    
    // 创建数据结构
    size_t rows = 200;
    size_t cols = stringCount / rows;
    
    std::vector<std::vector<cell_value_t>> testData;
    testData.reserve(rows);
    
    auto dataStructTime = measureTime([&]() {
        size_t stringIndex = 0;
        for (size_t r = 0; r < rows && stringIndex < stringCount; ++r) {
            std::vector<cell_value_t> rowData;
            rowData.reserve(cols);
            for (size_t c = 0; c < cols && stringIndex < stringCount; ++c) {
                rowData.emplace_back(testStrings[stringIndex++]);
            }
            testData.push_back(std::move(rowData));
        }
    });
    
    printPerformanceResult("数据结构构建", dataStructTime, stringCount, "strings");
    
    // 填充数据到工作表
    auto fillTime = measureTime([&]() {
        sheet->setRangeValues(row_t(1), column_t(1), testData);
    });
    
    printPerformanceResult("数据填充", fillTime, stringCount, "strings");
    
    // 保存文件
    std::string filename = "large_shared_strings_test.xlsx";
    testFiles_.push_back(filename);
    
    auto saveTime = measureTime([&]() {
        EXPECT_TRUE(workbook.saveToFile(filename));
    });
    
    printPerformanceResult("文件保存", saveTime, stringCount, "strings");
    
    // 文件信息
    std::ifstream file(filename, std::ios::binary | std::ios::ate);
    if (file.is_open()) {
        size_t fileSize = file.tellg();
        double fileSizeMB = fileSize / (1024.0 * 1024.0);
        
        std::cout << "大量字符串文件信息:\n";
        std::cout << "  文件大小: " << std::fixed << std::setprecision(2) << fileSizeMB << " MB\n";
        std::cout << "  每字符串: " << std::fixed << std::setprecision(1) 
                  << (static_cast<double>(fileSize) / stringCount) << " bytes\n";
        
        std::set<std::string> uniqueStrings(testStrings.begin(), testStrings.end());
        std::cout << "  唯一字符串: " << uniqueStrings.size() << "\n";
        std::cout << "  压缩效率: " << std::fixed << std::setprecision(1)
                  << (static_cast<double>(uniqueStrings.size()) / stringCount) * 100 << "%\n";
        
        std::cout << "\n✅ 大量字符串测试完成\n\n";
    }
}

// 测试字符串长度对性能的影响
TEST_F(SharedStringsPerformanceTest, StringLengthImpactTest) {
    std::cout << "=== 字符串长度影响测试 ===\n\n";
    
    std::vector<std::pair<std::string, size_t>> lengthTests = {
        {"短字符串", 5},
        {"中等字符串", 20},
        {"长字符串", 100},
        {"超长字符串", 500}
    };
    
    const size_t stringCount = 5000; // 使用流式写入器
    
    for (const auto& [testName, avgLength] : lengthTests) {
        std::cout << "--- " << testName << " (平均长度: " << avgLength << ") ---\n";
        
        TXWorkbook workbook;
        auto sheet = workbook.addSheet("LengthTest");
        
        // 生成指定长度的字符串
        std::vector<std::string> testStrings;
        testStrings.reserve(stringCount);
        
        auto stringGenTime = measureTime([&]() {
            for (size_t i = 0; i < stringCount; ++i) {
                size_t length = avgLength + (i % 10) - 5; // 长度变化±5
                if (length < 1) length = 1;
                testStrings.push_back(generateRandomString(length));
            }
        });
        
        printPerformanceResult("字符串生成", stringGenTime, stringCount, "strings");
        
        // 创建数据并填充
        std::vector<std::vector<cell_value_t>> testData;
        size_t rows = 100;
        size_t cols = stringCount / rows;
        testData.reserve(rows);
        
        size_t stringIndex = 0;
        for (size_t r = 0; r < rows && stringIndex < stringCount; ++r) {
            std::vector<cell_value_t> rowData;
            rowData.reserve(cols);
            for (size_t c = 0; c < cols && stringIndex < stringCount; ++c) {
                rowData.emplace_back(testStrings[stringIndex++]);
            }
            testData.push_back(std::move(rowData));
        }
        
        auto fillTime = measureTime([&]() {
            sheet->setRangeValues(row_t(1), column_t(1), testData);
        });
        
        printPerformanceResult("数据填充", fillTime, stringCount, "strings");
        
        // 保存文件
        std::string filename = "string_length_test_" + std::to_string(avgLength) + ".xlsx";
        testFiles_.push_back(filename);
        
        auto saveTime = measureTime([&]() {
            EXPECT_TRUE(workbook.saveToFile(filename));
        });
        
        printPerformanceResult("文件保存", saveTime, stringCount, "strings");
        
        // 计算平均字符串长度
        size_t totalLength = 0;
        for (const auto& str : testStrings) {
            totalLength += str.length();
        }
        double actualAvgLength = static_cast<double>(totalLength) / stringCount;
        
        std::cout << "字符串统计:\n";
        std::cout << "  实际平均长度: " << std::fixed << std::setprecision(1) << actualAvgLength << "\n";
        std::cout << "  总字符数: " << totalLength << "\n\n";
    }
}

// 测试性能稳定性
TEST_F(SharedStringsPerformanceTest, PerformanceStabilityTest) {
    std::cout << "=== SharedStrings性能稳定性测试 ===\n\n";
    
    const size_t stringCount = 3000; // 使用流式写入器
    const size_t numTests = 5;
    
    std::vector<std::chrono::microseconds> saveTimes;
    saveTimes.reserve(numTests);
    
    for (size_t test = 0; test < numTests; ++test) {
        TXWorkbook workbook;
        auto sheet = workbook.addSheet("StabilityTest");
        
        // 生成测试字符串
        auto testStrings = generateStringSet(stringCount, 0.3);
        
        // 创建数据
        std::vector<std::vector<cell_value_t>> testData;
        size_t rows = 60;
        size_t cols = stringCount / rows;
        testData.reserve(rows);
        
        size_t stringIndex = 0;
        for (size_t r = 0; r < rows && stringIndex < stringCount; ++r) {
            std::vector<cell_value_t> rowData;
            rowData.reserve(cols);
            for (size_t c = 0; c < cols && stringIndex < stringCount; ++c) {
                rowData.emplace_back(testStrings[stringIndex++]);
            }
            testData.push_back(std::move(rowData));
        }
        
        // 填充数据
        sheet->setRangeValues(row_t(1), column_t(1), testData);
        
        // 测试保存性能
        std::string filename = "stability_test_" + std::to_string(test) + ".xlsx";
        testFiles_.push_back(filename);
        
        auto saveTime = measureTime([&]() {
            EXPECT_TRUE(workbook.saveToFile(filename));
        });
        
        saveTimes.push_back(saveTime);
        
        double timePerString = static_cast<double>(saveTime.count()) / stringCount;
        std::cout << "测试 " << (test + 1) << "/" << numTests 
                  << ": " << saveTime.count() << "μs"
                  << ", 平均: " << std::fixed << std::setprecision(2) << timePerString << "μs/string\n";
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
    
    std::cout << "\n✅ 性能稳定性测试完成\n\n";
}
