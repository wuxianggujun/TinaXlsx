//
// TinaXlsx 简化性能测试
// 专注于核心功能的性能测试，避免复杂依赖
//

#include <gtest/gtest.h>
#include <chrono>
#include <memory>
#include <random>
#include <sstream>
#include <iomanip>
#include <filesystem>

#ifdef _WIN32
#include <windows.h>
#include <psapi.h>
#pragma comment(lib, "psapi.lib")
#endif

#include "TinaXlsx/TinaXlsx.hpp"

using namespace TinaXlsx;
using namespace std::chrono;

class SimplePerformanceTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 创建输出目录
        std::filesystem::create_directories("test_output/performance");
        
        // 初始化随机数生成器
        rng_.seed(12345);
        
        test_start_time_ = high_resolution_clock::now();
        std::cout << "\n=== 简化性能测试开始 ===" << std::endl;
    }

    void TearDown() override {
        auto test_end_time = high_resolution_clock::now();
        auto total_duration = duration_cast<milliseconds>(test_end_time - test_start_time_);
        std::cout << "=== 测试总耗时: " << total_duration.count() << "ms ===" << std::endl;
    }

    // 性能计时器
    class PerformanceTimer {
    public:
        PerformanceTimer(const std::string& name) : name_(name) {
            start_time_ = high_resolution_clock::now();
        }
        
        ~PerformanceTimer() {
            auto end_time = high_resolution_clock::now();
            auto duration = duration_cast<microseconds>(end_time - start_time_);
            std::cout << "[性能] " << name_ << ": " << duration.count() << "μs" << std::endl;
        }
        
    private:
        std::string name_;
        high_resolution_clock::time_point start_time_;
    };

    // 内存使用监控
    size_t getCurrentMemoryUsage() {
        #ifdef _WIN32
        PROCESS_MEMORY_COUNTERS pmc;
        if (GetProcessMemoryInfo(GetCurrentProcess(), &pmc, sizeof(pmc))) {
            return pmc.WorkingSetSize;
        }
        #endif
        return 0;
    }

    // 生成随机字符串
    std::string generateRandomString(size_t length) {
        const std::string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        std::uniform_int_distribution<> dis(0, chars.size() - 1);
        
        std::string result;
        result.reserve(length);
        for (size_t i = 0; i < length; ++i) {
            result += chars[dis(rng_)];
        }
        return result;
    }

    // 生成随机数值
    double generateRandomNumber() {
        std::uniform_real_distribution<double> dis(0.0, 1000000.0);
        return dis(rng_);
    }

    // 格式化内存大小
    std::string formatMemorySize(size_t bytes) {
        const char* units[] = {"B", "KB", "MB", "GB"};
        int unit = 0;
        double size = static_cast<double>(bytes);
        
        while (size >= 1024.0 && unit < 3) {
            size /= 1024.0;
            unit++;
        }
        
        std::ostringstream oss;
        oss << std::fixed << std::setprecision(2) << size << " " << units[unit];
        return oss.str();
    }

protected:
    std::mt19937 rng_;
    high_resolution_clock::time_point test_start_time_;
};

// 测试1: 基础数据写入性能
TEST_F(SimplePerformanceTest, BasicDataWritePerformance) {
    std::cout << "\n--- 测试1: 基础数据写入性能 ---" << std::endl;
    
    const int ROWS = 10000;
    const int COLS = 10;
    
    size_t initial_memory = getCurrentMemoryUsage();
    std::cout << "初始内存使用: " << formatMemorySize(initial_memory) << std::endl;
    
    TXWorkbook workbook;
    TXSheet* sheet = workbook.addSheet("性能测试");
    
    {
        PerformanceTimer timer("基础数据写入(" + std::to_string(ROWS * COLS) + "个单元格)");
        
        for (int row = 1; row <= ROWS; ++row) {
            for (int col = 1; col <= COLS; ++col) {
                if (col % 2 == 0) {
                    // 数值
                    sheet->setCellValue(row_t(row), column_t(col), generateRandomNumber());
                } else {
                    // 字符串
                    sheet->setCellValue(row_t(row), column_t(col), generateRandomString(10));
                }
            }
            
            if (row % 1000 == 0) {
                size_t current_memory = getCurrentMemoryUsage();
                std::cout << "进度: " << row << "/" << ROWS 
                         << ", 内存: " << formatMemorySize(current_memory) << std::endl;
            }
        }
    }
    
    size_t after_write_memory = getCurrentMemoryUsage();
    std::cout << "写入后内存使用: " << formatMemorySize(after_write_memory) << std::endl;
    std::cout << "内存增长: " << formatMemorySize(after_write_memory - initial_memory) << std::endl;
    
    {
        PerformanceTimer timer("文件保存");
        bool success = workbook.saveToFile("test_output/performance/basic_performance_test.xlsx");
        EXPECT_TRUE(success) << "文件保存失败: " << workbook.getLastError();
    }
    
    auto file_size = std::filesystem::file_size("test_output/performance/basic_performance_test.xlsx");
    std::cout << "生成文件大小: " << formatMemorySize(file_size) << std::endl;
}

// 测试2: 纯数值性能测试
TEST_F(SimplePerformanceTest, NumericOnlyPerformance) {
    std::cout << "\n--- 测试2: 纯数值性能测试 ---" << std::endl;
    
    const int ROWS = 20000;
    const int COLS = 5;
    
    size_t initial_memory = getCurrentMemoryUsage();
    
    TXWorkbook workbook;
    TXSheet* sheet = workbook.addSheet("纯数值测试");
    
    {
        PerformanceTimer timer("纯数值写入");
        
        for (int row = 1; row <= ROWS; ++row) {
            for (int col = 1; col <= COLS; ++col) {
                sheet->setCellValue(row_t(row), column_t(col), static_cast<double>(row * col + 0.123));
            }
        }
    }
    
    size_t after_write_memory = getCurrentMemoryUsage();
    std::cout << "数值写入后内存: " << formatMemorySize(after_write_memory) << std::endl;
    std::cout << "内存增长: " << formatMemorySize(after_write_memory - initial_memory) << std::endl;
    
    {
        PerformanceTimer timer("数值文件保存");
        bool success = workbook.saveToFile("test_output/performance/numeric_only_test.xlsx");
        EXPECT_TRUE(success);
    }
    
    auto file_size = std::filesystem::file_size("test_output/performance/numeric_only_test.xlsx");
    std::cout << "数值文件大小: " << formatMemorySize(file_size) << std::endl;
}

// 测试3: 纯字符串性能测试
TEST_F(SimplePerformanceTest, StringOnlyPerformance) {
    std::cout << "\n--- 测试3: 纯字符串性能测试 ---" << std::endl;
    
    const int ROWS = 10000;
    const int COLS = 5;
    
    size_t initial_memory = getCurrentMemoryUsage();
    
    TXWorkbook workbook;
    TXSheet* sheet = workbook.addSheet("纯字符串测试");
    
    {
        PerformanceTimer timer("纯字符串写入");
        
        for (int row = 1; row <= ROWS; ++row) {
            for (int col = 1; col <= COLS; ++col) {
                sheet->setCellValue(row_t(row), column_t(col), 
                    "Row" + std::to_string(row) + "_Col" + std::to_string(col));
            }
        }
    }
    
    size_t after_write_memory = getCurrentMemoryUsage();
    std::cout << "字符串写入后内存: " << formatMemorySize(after_write_memory) << std::endl;
    std::cout << "内存增长: " << formatMemorySize(after_write_memory - initial_memory) << std::endl;
    
    {
        PerformanceTimer timer("字符串文件保存");
        bool success = workbook.saveToFile("test_output/performance/string_only_test.xlsx");
        EXPECT_TRUE(success);
    }
    
    auto file_size = std::filesystem::file_size("test_output/performance/string_only_test.xlsx");
    std::cout << "字符串文件大小: " << formatMemorySize(file_size) << std::endl;
}

// 测试4: 多工作表性能测试
TEST_F(SimplePerformanceTest, MultipleSheetPerformance) {
    std::cout << "\n--- 测试4: 多工作表性能测试 ---" << std::endl;
    
    const int SHEET_COUNT = 10;
    const int ROWS_PER_SHEET = 1000;
    const int COLS_PER_SHEET = 5;
    
    size_t initial_memory = getCurrentMemoryUsage();
    
    TXWorkbook workbook;
    
    {
        PerformanceTimer timer("创建" + std::to_string(SHEET_COUNT) + "个工作表");
        
        for (int sheet_idx = 0; sheet_idx < SHEET_COUNT; ++sheet_idx) {
            std::string sheet_name = "Sheet" + std::to_string(sheet_idx + 1);
            TXSheet* sheet = workbook.addSheet(sheet_name);
            
            for (int row = 1; row <= ROWS_PER_SHEET; ++row) {
                for (int col = 1; col <= COLS_PER_SHEET; ++col) {
                    sheet->setCellValue(row_t(row), column_t(col), 
                        "S" + std::to_string(sheet_idx) + "_R" + std::to_string(row) + "_C" + std::to_string(col));
                }
            }
        }
    }
    
    size_t after_creation_memory = getCurrentMemoryUsage();
    std::cout << "创建后内存使用: " << formatMemorySize(after_creation_memory) << std::endl;
    std::cout << "内存增长: " << formatMemorySize(after_creation_memory - initial_memory) << std::endl;
    
    {
        PerformanceTimer timer("保存多工作表文件");
        bool success = workbook.saveToFile("test_output/performance/multiple_sheets_simple_test.xlsx");
        EXPECT_TRUE(success);
    }
    
    auto file_size = std::filesystem::file_size("test_output/performance/multiple_sheets_simple_test.xlsx");
    std::cout << "多工作表文件大小: " << formatMemorySize(file_size) << std::endl;
}

// 测试5: 内存使用监控测试
TEST_F(SimplePerformanceTest, MemoryUsageMonitoring) {
    std::cout << "\n--- 测试5: 内存使用监控测试 ---" << std::endl;
    
    const int ITERATIONS = 10;
    const int ROWS_PER_ITERATION = 1000;
    const int COLS_PER_ITERATION = 5;
    
    size_t initial_memory = getCurrentMemoryUsage();
    std::vector<size_t> memory_snapshots;
    
    {
        PerformanceTimer timer("内存监控测试(" + std::to_string(ITERATIONS) + "次迭代)");
        
        for (int iter = 0; iter < ITERATIONS; ++iter) {
            {
                TXWorkbook workbook;
                TXSheet* sheet = workbook.addSheet("内存测试");
                
                for (int row = 1; row <= ROWS_PER_ITERATION; ++row) {
                    for (int col = 1; col <= COLS_PER_ITERATION; ++col) {
                        sheet->setCellValue(row_t(row), column_t(col), generateRandomString(20));
                    }
                }
                
                std::string filename = "test_output/performance/memory_test_" + std::to_string(iter) + ".xlsx";
                workbook.saveToFile(filename);
                std::filesystem::remove(filename); // 立即删除以节省空间
            }
            
            size_t current_memory = getCurrentMemoryUsage();
            memory_snapshots.push_back(current_memory);
            
            std::cout << "迭代 " << (iter + 1) << "/" << ITERATIONS 
                     << ", 内存: " << formatMemorySize(current_memory) << std::endl;
        }
    }
    
    size_t final_memory = getCurrentMemoryUsage();
    size_t memory_growth = final_memory - initial_memory;
    
    std::cout << "初始内存: " << formatMemorySize(initial_memory) << std::endl;
    std::cout << "最终内存: " << formatMemorySize(final_memory) << std::endl;
    std::cout << "总内存增长: " << formatMemorySize(memory_growth) << std::endl;
    
    // 简单的内存泄漏检测
    if (memory_growth > initial_memory * 0.1) { // 增长超过10%
        std::cout << "⚠️  警告: 检测到可能的内存泄漏!" << std::endl;
    } else {
        std::cout << "✅ 内存使用稳定" << std::endl;
    }
}
