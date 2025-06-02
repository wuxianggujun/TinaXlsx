//
// TinaXlsx 极致性能测试
// 测试大数据量下的读写性能，发现性能瓶颈和内存泄漏
//

#include <gtest/gtest.h>
#include <chrono>
#include <memory>
#include <random>
#include <sstream>
#include <iomanip>
#include <filesystem>
#include <thread>
#include <atomic>
#include <vector>

#ifdef _WIN32
#include <windows.h>
#include <psapi.h>
#pragma comment(lib, "psapi.lib")
#endif

#include "TinaXlsx/TinaXlsx.hpp"

using namespace TinaXlsx;
using namespace std::chrono;

class ExtremePerformanceTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 创建输出目录
        std::filesystem::create_directories("test_output/performance");
        
        // 初始化随机数生成器
        rng_.seed(12345); // 固定种子确保可重现性
        
        // 记录测试开始时间
        test_start_time_ = high_resolution_clock::now();
        
        std::cout << "\n=== 极致性能测试开始 ===" << std::endl;
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
        // Windows平台的内存使用获取
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

// 测试1: 大量数据写入性能
TEST_F(ExtremePerformanceTest, MassiveDataWritePerformance) {
    std::cout << "\n--- 测试1: 大量数据写入性能 ---" << std::endl;
    
    const int ROWS = 50000;    // 5万行
    const int COLS = 20;       // 20列
    
    size_t initial_memory = getCurrentMemoryUsage();
    std::cout << "初始内存使用: " << formatMemorySize(initial_memory) << std::endl;
    
    TXWorkbook workbook;
    TXSheet* sheet = workbook.addSheet("大数据测试");
    
    {
        PerformanceTimer timer("大量数据写入(" + std::to_string(ROWS * COLS) + "个单元格)");
        
        // 写入数据
        for (int row = 1; row <= ROWS; ++row) {
            for (int col = 1; col <= COLS; ++col) {
                if (col % 3 == 0) {
                    // 每3列写入数值
                    sheet->setCellValue(row_t(row), column_t(col), generateRandomNumber());
                } else {
                    // 其他列写入字符串
                    sheet->setCellValue(row_t(row), column_t(col), generateRandomString(10));
                }
            }
            
            // 每1000行报告一次进度
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
    
    // 保存文件
    {
        PerformanceTimer timer("文件保存");
        bool success = workbook.saveToFile("test_output/performance/massive_data_test.xlsx");
        EXPECT_TRUE(success) << "文件保存失败: " << workbook.getLastError();
    }
    
    size_t final_memory = getCurrentMemoryUsage();
    std::cout << "保存后内存使用: " << formatMemorySize(final_memory) << std::endl;
    
    // 检查文件大小
    auto file_size = std::filesystem::file_size("test_output/performance/massive_data_test.xlsx");
    std::cout << "生成文件大小: " << formatMemorySize(file_size) << std::endl;
}

// 测试2: 大文件读取性能
TEST_F(ExtremePerformanceTest, MassiveDataReadPerformance) {
    std::cout << "\n--- 测试2: 大文件读取性能 ---" << std::endl;
    
    // 首先确保测试文件存在
    if (!std::filesystem::exists("test_output/performance/massive_data_test.xlsx")) {
        GTEST_SKIP() << "测试文件不存在，跳过读取测试";
    }
    
    size_t initial_memory = getCurrentMemoryUsage();
    std::cout << "初始内存使用: " << formatMemorySize(initial_memory) << std::endl;
    
    TXWorkbook workbook;
    
    {
        PerformanceTimer timer("大文件加载");
        bool success = workbook.loadFromFile("test_output/performance/massive_data_test.xlsx");
        EXPECT_TRUE(success) << "文件加载失败: " << workbook.getLastError();
    }
    
    size_t after_load_memory = getCurrentMemoryUsage();
    std::cout << "加载后内存使用: " << formatMemorySize(after_load_memory) << std::endl;
    std::cout << "内存增长: " << formatMemorySize(after_load_memory - initial_memory) << std::endl;
    
    // 随机访问测试
    TXSheet* sheet = workbook.getSheet(static_cast<u64>(0));
    ASSERT_NE(sheet, nullptr);
    
    {
        PerformanceTimer timer("随机单元格访问(10000次)");
        
        std::uniform_int_distribution<> row_dis(1, 50000);
        std::uniform_int_distribution<> col_dis(1, 20);
        
        for (int i = 0; i < 10000; ++i) {
            int row = row_dis(rng_);
            int col = col_dis(rng_);
            auto value = sheet->getCellValue(row_t(row), column_t(col));
            // 简单使用值以防止编译器优化
            (void)value;
        }
    }
    
    size_t final_memory = getCurrentMemoryUsage();
    std::cout << "访问后内存使用: " << formatMemorySize(final_memory) << std::endl;
}

// 测试3: 多工作表性能
TEST_F(ExtremePerformanceTest, MultipleSheetPerformance) {
    std::cout << "\n--- 测试3: 多工作表性能 ---" << std::endl;
    
    const int SHEET_COUNT = 50;
    const int ROWS_PER_SHEET = 1000;
    const int COLS_PER_SHEET = 10;
    
    size_t initial_memory = getCurrentMemoryUsage();
    
    TXWorkbook workbook;
    
    {
        PerformanceTimer timer("创建" + std::to_string(SHEET_COUNT) + "个工作表");
        
        for (int sheet_idx = 0; sheet_idx < SHEET_COUNT; ++sheet_idx) {
            std::string sheet_name = "Sheet" + std::to_string(sheet_idx + 1);
            TXSheet* sheet = workbook.addSheet(sheet_name);
            
            // 为每个工作表填充数据
            for (int row = 1; row <= ROWS_PER_SHEET; ++row) {
                for (int col = 1; col <= COLS_PER_SHEET; ++col) {
                    sheet->setCellValue(row_t(row), column_t(col), 
                                      "Sheet" + std::to_string(sheet_idx) + "_R" + std::to_string(row) + "_C" + std::to_string(col));
                }
            }
            
            if ((sheet_idx + 1) % 10 == 0) {
                size_t current_memory = getCurrentMemoryUsage();
                std::cout << "已创建 " << (sheet_idx + 1) << " 个工作表, 内存: " 
                         << formatMemorySize(current_memory) << std::endl;
            }
        }
    }
    
    size_t after_creation_memory = getCurrentMemoryUsage();
    std::cout << "创建后内存使用: " << formatMemorySize(after_creation_memory) << std::endl;
    std::cout << "内存增长: " << formatMemorySize(after_creation_memory - initial_memory) << std::endl;
    
    {
        PerformanceTimer timer("保存多工作表文件");
        bool success = workbook.saveToFile("test_output/performance/multiple_sheets_test.xlsx");
        EXPECT_TRUE(success) << "多工作表文件保存失败: " << workbook.getLastError();
    }
    
    // 检查文件大小
    auto file_size = std::filesystem::file_size("test_output/performance/multiple_sheets_test.xlsx");
    std::cout << "多工作表文件大小: " << formatMemorySize(file_size) << std::endl;
}

// 测试4: 字符串池性能测试
TEST_F(ExtremePerformanceTest, SharedStringPoolPerformance) {
    std::cout << "\n--- 测试4: 字符串池性能测试 ---" << std::endl;

    const int ROWS = 10000;
    const int COLS = 10;
    const int UNIQUE_STRINGS = 100; // 只有100个唯一字符串，大量重复

    // 预生成唯一字符串
    std::vector<std::string> unique_strings;
    for (int i = 0; i < UNIQUE_STRINGS; ++i) {
        unique_strings.push_back("重复字符串_" + std::to_string(i) + "_" + generateRandomString(20));
    }

    size_t initial_memory = getCurrentMemoryUsage();

    TXWorkbook workbook;
    TXSheet* sheet = workbook.addSheet("字符串池测试");

    {
        PerformanceTimer timer("大量重复字符串写入");

        std::uniform_int_distribution<> string_dis(0, UNIQUE_STRINGS - 1);

        for (int row = 1; row <= ROWS; ++row) {
            for (int col = 1; col <= COLS; ++col) {
                // 随机选择一个重复字符串
                const std::string& str = unique_strings[string_dis(rng_)];
                sheet->setCellValue(row_t(row), column_t(col), str);
            }

            if (row % 1000 == 0) {
                size_t current_memory = getCurrentMemoryUsage();
                std::cout << "字符串写入进度: " << row << "/" << ROWS
                         << ", 内存: " << formatMemorySize(current_memory) << std::endl;
            }
        }
    }

    size_t after_write_memory = getCurrentMemoryUsage();
    std::cout << "字符串写入后内存: " << formatMemorySize(after_write_memory) << std::endl;
    std::cout << "内存增长: " << formatMemorySize(after_write_memory - initial_memory) << std::endl;

    {
        PerformanceTimer timer("字符串池文件保存");
        bool success = workbook.saveToFile("test_output/performance/string_pool_test.xlsx");
        EXPECT_TRUE(success) << "字符串池文件保存失败: " << workbook.getLastError();
    }

    auto file_size = std::filesystem::file_size("test_output/performance/string_pool_test.xlsx");
    std::cout << "字符串池文件大小: " << formatMemorySize(file_size) << std::endl;
}

// 测试5: 数值类型性能测试
TEST_F(ExtremePerformanceTest, NumericDataPerformance) {
    std::cout << "\n--- 测试5: 数值类型性能测试 ---" << std::endl;

    const int ROWS = 10000;
    const int COLS = 10;

    size_t initial_memory = getCurrentMemoryUsage();

    TXWorkbook workbook;
    TXSheet* sheet = workbook.addSheet("数值测试");

    {
        PerformanceTimer timer("纯数值数据写入");

        for (int row = 1; row <= ROWS; ++row) {
            for (int col = 1; col <= COLS; ++col) {
                // 写入不同类型的数值
                if (col % 4 == 0) {
                    sheet->setCellValue(row_t(row), column_t(col), static_cast<i64>(row * col));
                } else if (col % 4 == 1) {
                    sheet->setCellValue(row_t(row), column_t(col), static_cast<f64>(row * col * 0.123));
                } else if (col % 4 == 2) {
                    sheet->setCellValue(row_t(row), column_t(col), static_cast<f64>(row + col));
                } else {
                    sheet->setCellValue(row_t(row), column_t(col), static_cast<i64>(row - col));
                }
            }

            if (row % 1000 == 0) {
                size_t current_memory = getCurrentMemoryUsage();
                std::cout << "数值写入进度: " << row << "/" << ROWS
                         << ", 内存: " << formatMemorySize(current_memory) << std::endl;
            }
        }
    }

    size_t after_write_memory = getCurrentMemoryUsage();
    std::cout << "数值写入后内存: " << formatMemorySize(after_write_memory) << std::endl;
    std::cout << "内存增长: " << formatMemorySize(after_write_memory - initial_memory) << std::endl;

    {
        PerformanceTimer timer("数值文件保存");
        bool success = workbook.saveToFile("test_output/performance/numeric_test.xlsx");
        EXPECT_TRUE(success) << "数值文件保存失败: " << workbook.getLastError();
    }

    auto file_size = std::filesystem::file_size("test_output/performance/numeric_test.xlsx");
    std::cout << "数值文件大小: " << formatMemorySize(file_size) << std::endl;
}

// 测试6: 内存泄漏检测
TEST_F(ExtremePerformanceTest, MemoryLeakDetection) {
    std::cout << "\n--- 测试6: 内存泄漏检测 ---" << std::endl;

    const int ITERATIONS = 100;
    const int ROWS_PER_ITERATION = 1000;
    const int COLS_PER_ITERATION = 10;

    size_t initial_memory = getCurrentMemoryUsage();
    std::vector<size_t> memory_snapshots;

    {
        PerformanceTimer timer("内存泄漏检测(" + std::to_string(ITERATIONS) + "次迭代)");

        for (int iter = 0; iter < ITERATIONS; ++iter) {
            // 创建工作簿
            {
                TXWorkbook workbook;
                TXSheet* sheet = workbook.addSheet("泄漏测试");

                // 填充数据
                for (int row = 1; row <= ROWS_PER_ITERATION; ++row) {
                    for (int col = 1; col <= COLS_PER_ITERATION; ++col) {
                        sheet->setCellValue(row_t(row), column_t(col), generateRandomString(50));
                    }
                }

                // 保存文件
                std::string filename = "test_output/performance/leak_test_" + std::to_string(iter) + ".xlsx";
                workbook.saveToFile(filename);

                // 删除文件以节省磁盘空间
                std::filesystem::remove(filename);
            } // workbook 在这里销毁

            // 记录内存使用
            size_t current_memory = getCurrentMemoryUsage();
            memory_snapshots.push_back(current_memory);

            if ((iter + 1) % 10 == 0) {
                std::cout << "迭代 " << (iter + 1) << "/" << ITERATIONS
                         << ", 内存: " << formatMemorySize(current_memory) << std::endl;
            }
        }
    }

    // 分析内存趋势
    size_t final_memory = getCurrentMemoryUsage();
    size_t memory_growth = final_memory - initial_memory;

    std::cout << "初始内存: " << formatMemorySize(initial_memory) << std::endl;
    std::cout << "最终内存: " << formatMemorySize(final_memory) << std::endl;
    std::cout << "总内存增长: " << formatMemorySize(memory_growth) << std::endl;

    // 计算内存增长趋势
    if (memory_snapshots.size() >= 10) {
        size_t first_10_avg = 0, last_10_avg = 0;

        for (int i = 0; i < 10; ++i) {
            first_10_avg += memory_snapshots[i];
            last_10_avg += memory_snapshots[memory_snapshots.size() - 10 + i];
        }
        first_10_avg /= 10;
        last_10_avg /= 10;

        std::cout << "前10次平均内存: " << formatMemorySize(first_10_avg) << std::endl;
        std::cout << "后10次平均内存: " << formatMemorySize(last_10_avg) << std::endl;

        if (last_10_avg > first_10_avg * 1.1) { // 增长超过10%
            std::cout << "⚠️  警告: 检测到可能的内存泄漏!" << std::endl;
        } else {
            std::cout << "✅ 内存使用稳定，未检测到明显泄漏" << std::endl;
        }
    }
}

// 测试7: 并发安全性测试
TEST_F(ExtremePerformanceTest, ConcurrencyStressTest) {
    std::cout << "\n--- 测试7: 并发安全性测试 ---" << std::endl;

    const int THREAD_COUNT = 4;
    const int OPERATIONS_PER_THREAD = 100; // 减少操作数量以加快测试

    size_t initial_memory = getCurrentMemoryUsage();

    {
        PerformanceTimer timer("并发操作测试");

        std::vector<std::thread> threads;
        std::atomic<int> error_count{0};

        for (int t = 0; t < THREAD_COUNT; ++t) {
            threads.emplace_back([t, &error_count, OPERATIONS_PER_THREAD, this]() {
                try {
                    for (int op = 0; op < OPERATIONS_PER_THREAD; ++op) {
                        TXWorkbook workbook;
                        TXSheet* sheet = workbook.addSheet("Thread" + std::to_string(t));

                        // 每个线程写入不同的数据
                        for (int row = 1; row <= 100; ++row) {
                            for (int col = 1; col <= 10; ++col) {
                                std::string value = "T" + std::to_string(t) + "_R" + std::to_string(row) + "_C" + std::to_string(col);
                                sheet->setCellValue(row_t(row), column_t(col), value);
                            }
                        }

                        std::string filename = "test_output/performance/thread_" + std::to_string(t) + "_" + std::to_string(op) + ".xlsx";
                        if (!workbook.saveToFile(filename)) {
                            error_count++;
                        }

                        // 立即删除文件以节省空间
                        std::filesystem::remove(filename);
                    }
                } catch (const std::exception& e) {
                    error_count++;
                    std::cout << "线程 " << t << " 异常: " << e.what() << std::endl;
                }
            });
        }

        // 等待所有线程完成
        for (auto& thread : threads) {
            thread.join();
        }

        std::cout << "并发测试完成，错误数量: " << error_count.load() << std::endl;
        EXPECT_EQ(error_count.load(), 0) << "并发测试中发生错误";
    }

    size_t final_memory = getCurrentMemoryUsage();
    std::cout << "并发测试后内存: " << formatMemorySize(final_memory) << std::endl;
}

// 测试8: 极限单元格数量测试
TEST_F(ExtremePerformanceTest, ExtremeCellCountTest) {
    std::cout << "\n--- 测试8: 极限单元格数量测试 ---" << std::endl;

    // Excel理论最大值：1,048,576行 × 16,384列
    // 我们测试一个较小但仍然很大的数量
    const int MAX_ROWS = 100000;  // 10万行
    const int MAX_COLS = 50;      // 50列

    size_t initial_memory = getCurrentMemoryUsage();

    TXWorkbook workbook;
    TXSheet* sheet = workbook.addSheet("极限测试");

    {
        PerformanceTimer timer("极限单元格写入(" + std::to_string(MAX_ROWS * MAX_COLS) + "个单元格)");

        // 使用批量写入策略
        for (int row = 1; row <= MAX_ROWS; ++row) {
            for (int col = 1; col <= MAX_COLS; ++col) {
                // 交替写入数值和字符串以测试不同类型
                if ((row + col) % 2 == 0) {
                    sheet->setCellValue(row_t(row), column_t(col), static_cast<double>(row * col));
                } else {
                    sheet->setCellValue(row_t(row), column_t(col), "R" + std::to_string(row) + "C" + std::to_string(col));
                }
            }

            // 每5000行报告一次进度和内存使用
            if (row % 5000 == 0) {
                size_t current_memory = getCurrentMemoryUsage();
                double progress = static_cast<double>(row) / MAX_ROWS * 100.0;
                std::cout << "极限测试进度: " << std::fixed << std::setprecision(1) << progress << "% "
                         << "(" << row << "/" << MAX_ROWS << "), 内存: " << formatMemorySize(current_memory) << std::endl;
            }
        }
    }

    size_t after_write_memory = getCurrentMemoryUsage();
    std::cout << "极限写入后内存: " << formatMemorySize(after_write_memory) << std::endl;
    std::cout << "内存增长: " << formatMemorySize(after_write_memory - initial_memory) << std::endl;

    // 测试随机访问性能
    {
        PerformanceTimer timer("极限数据随机访问(10000次)");

        std::uniform_int_distribution<> row_dis(1, MAX_ROWS);
        std::uniform_int_distribution<> col_dis(1, MAX_COLS);

        for (int i = 0; i < 10000; ++i) {
            int row = row_dis(rng_);
            int col = col_dis(rng_);
            auto value = sheet->getCellValue(row_t(row), column_t(col));
            (void)value; // 防止编译器优化
        }
    }

    std::cout << "⚠️  注意: 由于文件过大，跳过保存测试" << std::endl;
    std::cout << "预估文件大小: 可能超过1GB" << std::endl;
}

// 测试9: 性能回归测试
TEST_F(ExtremePerformanceTest, PerformanceRegressionTest) {
    std::cout << "\n--- 测试9: 性能回归测试 ---" << std::endl;

    // 标准测试用例，用于检测性能回归
    const int STANDARD_ROWS = 10000;
    const int STANDARD_COLS = 10;

    struct BenchmarkResult {
        std::string operation;
        int64_t duration_us;
        size_t memory_used;
    };

    std::vector<BenchmarkResult> results;

    // 基准测试1: 纯数值写入
    {
        TXWorkbook workbook;
        TXSheet* sheet = workbook.addSheet("数值基准");

        size_t start_memory = getCurrentMemoryUsage();
        auto start_time = high_resolution_clock::now();

        for (int row = 1; row <= STANDARD_ROWS; ++row) {
            for (int col = 1; col <= STANDARD_COLS; ++col) {
                sheet->setCellValue(row_t(row), column_t(col), static_cast<double>(row * col));
            }
        }

        auto end_time = high_resolution_clock::now();
        size_t end_memory = getCurrentMemoryUsage();

        results.push_back({
            "数值写入",
            duration_cast<microseconds>(end_time - start_time).count(),
            end_memory - start_memory
        });
    }

    // 基准测试2: 纯字符串写入
    {
        TXWorkbook workbook;
        TXSheet* sheet = workbook.addSheet("字符串基准");

        size_t start_memory = getCurrentMemoryUsage();
        auto start_time = high_resolution_clock::now();

        for (int row = 1; row <= STANDARD_ROWS; ++row) {
            for (int col = 1; col <= STANDARD_COLS; ++col) {
                sheet->setCellValue(row_t(row), column_t(col), "Cell_" + std::to_string(row) + "_" + std::to_string(col));
            }
        }

        auto end_time = high_resolution_clock::now();
        size_t end_memory = getCurrentMemoryUsage();

        results.push_back({
            "字符串写入",
            duration_cast<microseconds>(end_time - start_time).count(),
            end_memory - start_memory
        });
    }

    // 基准测试3: 混合数据写入
    {
        TXWorkbook workbook;
        TXSheet* sheet = workbook.addSheet("混合基准");

        size_t start_memory = getCurrentMemoryUsage();
        auto start_time = high_resolution_clock::now();

        for (int row = 1; row <= STANDARD_ROWS; ++row) {
            for (int col = 1; col <= STANDARD_COLS; ++col) {
                if ((row + col) % 2 == 0) {
                    sheet->setCellValue(row_t(row), column_t(col), generateRandomNumber());
                } else {
                    sheet->setCellValue(row_t(row), column_t(col), generateRandomString(20));
                }
            }
        }

        auto end_time = high_resolution_clock::now();
        size_t end_memory = getCurrentMemoryUsage();

        results.push_back({
            "混合写入",
            duration_cast<microseconds>(end_time - start_time).count(),
            end_memory - start_memory
        });
    }

    // 输出基准测试结果
    std::cout << "\n=== 性能基准测试结果 ===" << std::endl;
    std::cout << std::left << std::setw(15) << "操作类型"
              << std::setw(15) << "耗时(μs)"
              << std::setw(15) << "内存使用"
              << std::setw(20) << "每单元格耗时(ns)" << std::endl;
    std::cout << std::string(65, '-') << std::endl;

    for (const auto& result : results) {
        double ns_per_cell = static_cast<double>(result.duration_us * 1000) / (STANDARD_ROWS * STANDARD_COLS);
        std::cout << std::left << std::setw(15) << result.operation
                  << std::setw(15) << result.duration_us
                  << std::setw(15) << formatMemorySize(result.memory_used)
                  << std::setw(20) << std::fixed << std::setprecision(2) << ns_per_cell << std::endl;
    }

    // 性能警告阈值
    const int64_t WARNING_THRESHOLD_US = 5000000; // 5秒
    for (const auto& result : results) {
        if (result.duration_us > WARNING_THRESHOLD_US) {
            std::cout << "⚠️  警告: " << result.operation << " 性能可能存在问题 (>"
                     << WARNING_THRESHOLD_US / 1000 << "ms)" << std::endl;
        }
    }
}
