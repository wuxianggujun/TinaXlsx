//
// @file test_high_performance_logger.cpp
// @brief 🚀 高性能日志库测试
//

#include <gtest/gtest.h>
#include <TinaXlsx/TXHighPerformanceLogger.hpp>
#include <TinaXlsx/TXUnifiedMemoryManager.hpp>
#include <chrono>
#include <thread>

using namespace TinaXlsx;

class HighPerformanceLoggerTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 初始化内存管理器
        TXUnifiedMemoryManager::Config config;
        config.memory_limit = 1024 * 1024 * 1024; // 1GB
        config.enable_monitoring = false; // 关闭监控以获得更好的性能
        GlobalUnifiedMemoryManager::initialize(config);

        // 初始化日志系统
        TXGlobalLogger::initialize(GlobalUnifiedMemoryManager::getInstance());

        // 🚀 纯净版本：默认就是高性能模式
        auto logger = TXGlobalLogger::getDefault();

        std::cout << "🚀 日志系统初始化完成（纯净高性能版本）" << std::endl;
    }
    
    void TearDown() override {
        TXGlobalLogger::shutdown();
        GlobalUnifiedMemoryManager::shutdown();
    }
};

TEST_F(HighPerformanceLoggerTest, BasicLogging) {
    // 🚀 基本日志测试 - 简化版本
    std::cout << "开始基本日志测试..." << std::endl;

    try {
        auto logger = TXGlobalLogger::getDefault();
        std::cout << "获取默认日志器成功" << std::endl;

        if (!logger) {
            FAIL() << "默认日志器为空";
            return;
        }

        std::cout << "开始写入日志..." << std::endl;
        std::cout << "=== 日志输出开始 ===" << std::endl;

        TX_LOG_INFO("这是一条信息日志");
        std::cout << "信息日志写入完成" << std::endl;

        TX_LOG_WARN("这是一条警告日志: {}", "测试参数");
        std::cout << "警告日志写入完成" << std::endl;

        TX_LOG_ERROR("这是一条错误日志: {} + {} = {}", 1, 2, 3);
        std::cout << "错误日志写入完成" << std::endl;

        std::cout << "=== 日志输出结束 ===" << std::endl;

        // 刷新确保输出
        std::cout << "开始刷新日志..." << std::endl;
        logger->flush();
        std::cout << "日志刷新完成" << std::endl;

        SUCCEED(); // 如果没有崩溃就算成功

    } catch (const std::exception& e) {
        FAIL() << "日志测试异常: " << e.what();
    }
}

TEST_F(HighPerformanceLoggerTest, PerformanceTest) {
    // 🚀 性能测试：10万条日志
    const int LOG_COUNT = 100000;

    // 🚀 启用性能模式以获得最佳性能
    TXGlobalLogger::setOutputMode(TXLogOutputMode::PERFORMANCE);

    // 🚀 禁用C++流与C stdio的同步以提高性能
    std::ios_base::sync_with_stdio(false);

    auto start = std::chrono::high_resolution_clock::now();
    
    for (int i = 0; i < LOG_COUNT; ++i) {
        TX_LOG_INFO("性能测试日志 #{}: 数值={}, 字符串={}", i, i * 1.5, "测试");
    }
    
    // 等待所有日志写入完成
    TXGlobalLogger::getDefault()->flush();
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    std::cout << "🚀 日志性能测试结果:" << std::endl;
    std::cout << "  - 日志数量: " << LOG_COUNT << std::endl;
    std::cout << "  - 总耗时: " << duration.count() << "ms" << std::endl;
    std::cout << "  - 性能: " << (LOG_COUNT * 1000.0 / duration.count()) << " 条/秒" << std::endl;
    
    // 🚀 期望性能：至少10万条/秒（性能模式）
    EXPECT_LT(duration.count(), 1000) << "性能模式应该在1秒内完成10万条日志";

    // 🚀 如果性能很好，给出更严格的期望
    if (duration.count() < 500) {
        std::cout << "🎉 性能优秀！达到20万条/秒以上" << std::endl;
    }

    // 🚀 恢复默认模式
    TXGlobalLogger::setOutputMode(TXLogOutputMode::CONSOLE_ONLY);
    std::ios_base::sync_with_stdio(true);
}

TEST_F(HighPerformanceLoggerTest, FileLogging) {
    // 🚀 文件日志测试 - 使用新的输出模式API
    TXGlobalLogger::setOutputMode(TXLogOutputMode::FILE_ONLY);
    std::cout << "🚀 已设置为文件输出模式" << std::endl;

    // 写入一些测试日志
    TX_LOG_INFO("文件日志测试开始");
    TX_LOG_DEBUG("调试信息: {}", "文件写入测试");
    TX_LOG_WARN("警告: 这是文件日志测试");
    TX_LOG_ERROR("错误: 测试错误日志");

    // 刷新确保写入文件
    TXGlobalLogger::getDefault()->flush();

    // 恢复控制台输出
    TXGlobalLogger::setOutputMode(TXLogOutputMode::CONSOLE_ONLY);

    SUCCEED();
}

TEST_F(HighPerformanceLoggerTest, CustomLogger) {
    // 🚀 自定义日志器测试
    auto custom_logger = TXGlobalLogger::create("CustomLogger", TXLogLevel::DEBUG);

    custom_logger->debug("这是自定义日志器的调试信息");
    custom_logger->info("自定义日志器信息: {}", "测试参数");
    custom_logger->warn("自定义日志器警告");

    custom_logger->flush();

    SUCCEED();
}

TEST_F(HighPerformanceLoggerTest, OutputModeTest) {
    // 🚀 输出模式测试
    std::cout << "🚀 测试不同的输出模式..." << std::endl;

    // 测试控制台输出
    TXGlobalLogger::setOutputMode(TXLogOutputMode::CONSOLE_ONLY);
    TX_LOG_INFO("控制台输出模式测试");

    // 测试文件输出
    TXGlobalLogger::setOutputMode(TXLogOutputMode::FILE_ONLY);
    TX_LOG_INFO("文件输出模式测试");

    // 测试双重输出
    TXGlobalLogger::setOutputMode(TXLogOutputMode::BOTH);
    TX_LOG_INFO("双重输出模式测试");

    // 测试性能模式
    TXGlobalLogger::setOutputMode(TXLogOutputMode::PERFORMANCE);
    TX_LOG_INFO("性能模式测试");

    // 恢复默认
    TXGlobalLogger::setOutputMode(TXLogOutputMode::CONSOLE_ONLY);

    TXGlobalLogger::getDefault()->flush();

    SUCCEED();
}

TEST_F(HighPerformanceLoggerTest, PerformanceMacroTest) {
    // 🚀 性能测试宏
    TX_PERF_LOG(info, {
        // 模拟一些工作
        std::this_thread::sleep_for(std::chrono::milliseconds(10));
        int sum = 0;
        for (int i = 0; i < 1000; ++i) {
            sum += i;
        }
    });
    
    TXGlobalLogger::getDefault()->flush();
    
    SUCCEED();
}

TEST_F(HighPerformanceLoggerTest, ZeroAllocationTest) {
    // 🚀 零分配测试（理论上，实际可能有一些分配）
    const int ITERATIONS = 1000;
    
    // 预热
    for (int i = 0; i < 100; ++i) {
        TX_LOG_INFO("预热日志 {}", i);
    }
    TXGlobalLogger::getDefault()->flush();
    
    // 测试零分配性能
    auto start = std::chrono::high_resolution_clock::now();
    
    for (int i = 0; i < ITERATIONS; ++i) {
        TX_LOG_INFO("零分配测试 #{}: 值={}", i, i * 2);
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "🚀 零分配性能测试:" << std::endl;
    std::cout << "  - 每条日志平均耗时: " << (duration.count() / double(ITERATIONS)) << " 微秒" << std::endl;
    
    // 期望每条日志耗时小于100微秒
    EXPECT_LT(duration.count() / double(ITERATIONS), 100.0) << "每条日志应该在100微秒内完成";
    
    TXGlobalLogger::getDefault()->flush();
}

TEST_F(HighPerformanceLoggerTest, MultiThreadTest) {
    // 🚀 多线程测试
    const int THREAD_COUNT = 4;
    const int LOGS_PER_THREAD = 1000;
    
    std::vector<std::thread> threads;
    
    auto start = std::chrono::high_resolution_clock::now();
    
    for (int t = 0; t < THREAD_COUNT; ++t) {
        threads.emplace_back([t, LOGS_PER_THREAD]() {
            for (int i = 0; i < LOGS_PER_THREAD; ++i) {
                TX_LOG_INFO("线程 {} 日志 #{}: 数据={}", t, i, i * t);
            }
        });
    }
    
    for (auto& thread : threads) {
        thread.join();
    }
    
    TXGlobalLogger::getDefault()->flush();
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    
    std::cout << "🚀 多线程日志测试:" << std::endl;
    std::cout << "  - 线程数: " << THREAD_COUNT << std::endl;
    std::cout << "  - 总日志数: " << (THREAD_COUNT * LOGS_PER_THREAD) << std::endl;
    std::cout << "  - 总耗时: " << duration.count() << "ms" << std::endl;
    std::cout << "  - 性能: " << (THREAD_COUNT * LOGS_PER_THREAD * 1000.0 / duration.count()) << " 条/秒" << std::endl;
    
    SUCCEED();
}

// 🚀 性能基准测试
TEST_F(HighPerformanceLoggerTest, BenchmarkVsStdCout) {
    const int ITERATIONS = 10000;

    // 🚀 启用性能模式进行公平比较
    TXGlobalLogger::setOutputMode(TXLogOutputMode::PERFORMANCE);
    std::cout << "🚀 已启用性能模式进行基准测试" << std::endl;

    // 测试标准cout性能
    auto start_cout = std::chrono::high_resolution_clock::now();
    for (int i = 0; i < ITERATIONS; ++i) {
        std::cout << "标准cout测试 #" << i << ": 值=" << (i * 1.5) << std::endl;
    }
    auto end_cout = std::chrono::high_resolution_clock::now();
    auto cout_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_cout - start_cout);

    // 测试我们的日志库性能（性能模式）
    auto start_log = std::chrono::high_resolution_clock::now();
    for (int i = 0; i < ITERATIONS; ++i) {
        TX_LOG_INFO("高性能日志测试 #{}: 值={}", i, i * 1.5);
    }
    TXGlobalLogger::getDefault()->flush();
    auto end_log = std::chrono::high_resolution_clock::now();
    auto log_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_log - start_log);
    
    std::cout << "🚀 性能对比测试:" << std::endl;
    std::cout << "  - std::cout: " << cout_duration.count() << "ms" << std::endl;
    std::cout << "  - TXLogger: " << log_duration.count() << "ms" << std::endl;
    std::cout << "  - 性能提升: " << (double(cout_duration.count()) / log_duration.count()) << "x" << std::endl;
    
    // 🚀 性能模式应该与std::cout性能相当或更好
    // 允许一定的性能差异（30%以内，因为我们有额外的格式化）
    EXPECT_LE(log_duration.count(), cout_duration.count() * 1.3) << "性能模式应该接近std::cout性能";
}
