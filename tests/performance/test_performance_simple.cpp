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
#include "TinaXlsx/TXMemoryLeakDetector.hpp"

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

        // 🚀 测试串行批量保存
        TXBatchWorksheetWriter::BatchConfig serialConfig;
        serialConfig.enableParallelGeneration = false;
        serialConfig.enableSharedStringOptim = true;
        serialConfig.maxConcurrentThreads = 1;

        auto serialStart = std::chrono::high_resolution_clock::now();
        bool serialSuccess = workbook.saveToFileBatch("test_output/performance/multiple_sheets_serial.xlsx", serialConfig);
        auto serialEnd = std::chrono::high_resolution_clock::now();
        auto serialDuration = std::chrono::duration_cast<std::chrono::milliseconds>(serialEnd - serialStart);

        if (!serialSuccess) {
            std::cout << "❌ 串行批量保存失败: " << workbook.getLastError() << std::endl;
        }
        EXPECT_TRUE(serialSuccess);

        if (serialSuccess) {
            const auto& serialStats = workbook.getLastBatchStats();
            std::cout << "串行保存统计 - 耗时: " << serialDuration.count() << "ms"
                     << ", XML生成: " << serialStats.xmlGenerationTimeMs << "ms" << std::endl;
        }

        // 🚀 测试并行批量保存（注意：10个工作表可能不足以体现并行优势）
        TXBatchWorksheetWriter::BatchConfig parallelConfig;
        parallelConfig.enableParallelGeneration = true;
        parallelConfig.enableSharedStringOptim = true;
        parallelConfig.maxConcurrentThreads = 2; // 减少线程数，避免过度竞争

        auto parallelStart = std::chrono::high_resolution_clock::now();
        bool parallelSuccess = workbook.saveToFileBatch("test_output/performance/multiple_sheets_parallel.xlsx", parallelConfig);
        auto parallelEnd = std::chrono::high_resolution_clock::now();
        auto parallelDuration = std::chrono::duration_cast<std::chrono::milliseconds>(parallelEnd - parallelStart);

        if (!parallelSuccess) {
            std::cout << "❌ 并行批量保存失败: " << workbook.getLastError() << std::endl;
        }
        EXPECT_TRUE(parallelSuccess);

        if (parallelSuccess) {
            const auto& parallelStats = workbook.getLastBatchStats();
            std::cout << "并行保存统计 - 耗时: " << parallelDuration.count() << "ms"
                     << ", XML生成: " << parallelStats.xmlGenerationTimeMs << "ms" << std::endl;

            // 🚀 计算性能提升
            if (serialDuration.count() > 0) {
                double improvement = (double(serialDuration.count()) - double(parallelDuration.count())) / double(serialDuration.count()) * 100.0;
                std::cout << "🚀 并行处理性能提升: " << std::fixed << std::setprecision(1) << improvement << "%" << std::endl;

                if (improvement < 0) {
                    std::cout << "💡 提示: 对于 " << SHEET_COUNT << " 个工作表，串行处理更高效。" << std::endl;
                    std::cout << "   并行处理适合于更多工作表（建议 >20 个）的场景。" << std::endl;
                } else if (improvement < 10) {
                    std::cout << "💡 提示: 性能提升较小，可能是因为工作表数量不够多。" << std::endl;
                }
            }
        }
    }
    
    // 检查并行保存的文件大小
    if (std::filesystem::exists("test_output/performance/multiple_sheets_parallel.xlsx")) {
        auto file_size = std::filesystem::file_size("test_output/performance/multiple_sheets_parallel.xlsx");
        std::cout << "多工作表文件大小: " << formatMemorySize(file_size) << std::endl;
    }
}

// 测试5: 高级内存泄漏检测测试
TEST_F(SimplePerformanceTest, AdvancedMemoryLeakDetection) {
    std::cout << "\n--- 测试5: 高级内存泄漏检测测试 ---" << std::endl;

    const int ITERATIONS = 10;
    const int ROWS_PER_ITERATION = 1000;
    const int COLS_PER_ITERATION = 5;

    // 🚀 启用内存泄漏检测器
    auto& detector = TXMemoryLeakDetector::instance();
    TXMemoryLeakDetector::DetectorConfig config;
    config.enableTracking = true;
    config.enableAutoCleanup = false; // 手动控制清理
    detector.setConfig(config);
    detector.reset(); // 重置统计

    size_t initial_memory = getCurrentMemoryUsage();
    auto initial_stats = detector.getStats();

    {
        PerformanceTimer timer("高级内存监控测试(" + std::to_string(ITERATIONS) + "次迭代)");

        for (int iter = 0; iter < ITERATIONS; ++iter) {
            {
                // 🚀 使用作用域内存跟踪器
                TXScopedMemoryTracker scope_tracker(("迭代" + std::to_string(iter)).c_str());

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

                // 检查作用域内存泄漏
                if (scope_tracker.hasLeaks()) {
                    auto scope_stats = scope_tracker.getScopeStats();
                    std::cout << "⚠️  迭代 " << iter << " 检测到作用域内存泄漏: "
                             << formatMemorySize(scope_stats.currentBytes) << std::endl;
                }
            }

            size_t current_memory = getCurrentMemoryUsage();
            auto current_stats = detector.getStats();

            std::cout << "迭代 " << (iter + 1) << "/" << ITERATIONS
                     << ", 系统内存: " << formatMemorySize(current_memory)
                     << ", 跟踪内存: " << formatMemorySize(current_stats.currentBytes) << std::endl;
        }
    }

    // 🚀 详细的内存泄漏分析
    size_t final_memory = getCurrentMemoryUsage();
    auto final_stats = detector.getStats();
    auto leak_report = detector.detectLeaks();

    std::cout << "\n=== 内存分析报告 ===" << std::endl;
    std::cout << "系统内存 - 初始: " << formatMemorySize(initial_memory)
             << ", 最终: " << formatMemorySize(final_memory)
             << ", 增长: " << formatMemorySize(final_memory - initial_memory) << std::endl;

    std::cout << "跟踪内存 - 当前分配: " << final_stats.currentAllocations
             << ", 当前字节: " << formatMemorySize(final_stats.currentBytes) << std::endl;
    std::cout << "总分配次数: " << final_stats.totalAllocations
             << ", 总释放次数: " << final_stats.totalDeallocations << std::endl;
    std::cout << "峰值分配: " << final_stats.peakAllocations
             << ", 峰值内存: " << formatMemorySize(final_stats.peakBytes) << std::endl;

    if (leak_report.leakedAllocations > 0) {
        std::cout << "🚨 检测到内存泄漏: " << leak_report.leakedAllocations
                 << " 个分配, 总计 " << formatMemorySize(leak_report.totalLeakedBytes) << std::endl;

        // 显示前5个泄漏
        size_t show_count = std::min(size_t(5), leak_report.leaks.size());
        for (size_t i = 0; i < show_count; ++i) {
            const auto& [ptr, info] = leak_report.leaks[i];
            std::cout << "  泄漏 " << (i+1) << ": " << formatMemorySize(info.size);
            if (info.file) {
                std::cout << " 在 " << info.file << ":" << info.line;
            }
            std::cout << std::endl;
        }
    } else {
        std::cout << "✅ 未检测到长期内存泄漏" << std::endl;
    }

    // 🚀 使用智能内存管理器进行健康检查
    auto& smart_manager = TXSmartMemoryManager::instance();
    auto health_report = smart_manager.performHealthCheck();

    std::cout << "\n=== 内存健康报告 ===" << std::endl;
    std::cout << "内存效率: " << (health_report.memoryEfficiency * 100) << "%" << std::endl;
    std::cout << "碎片化级别: " << health_report.fragmentationLevel << std::endl;

    if (!health_report.recommendations.empty()) {
        std::cout << "优化建议:" << std::endl;
        for (const auto& recommendation : health_report.recommendations) {
            std::cout << "  - " << recommendation << std::endl;
        }
    }

    // 最终判断
    size_t memory_growth = final_memory - initial_memory;
    if (memory_growth > initial_memory * 0.1 || health_report.hasLeaks) {
        std::cout << "⚠️  警告: 检测到可能的内存泄漏!" << std::endl;
    } else {
        std::cout << "✅ 内存使用健康" << std::endl;
    }
}
