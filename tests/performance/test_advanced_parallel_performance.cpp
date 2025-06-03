//
// @file test_advanced_parallel_performance.cpp
// @brief 高级并行框架性能测试
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXAdvancedParallelFramework.hpp"
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXMemoryLeakDetector.hpp"
#include "test_file_generator.hpp"
#include <chrono>
#include <random>
#include <iostream>
#include <filesystem>

using namespace TinaXlsx;

class AdvancedParallelPerformanceTest : public TestWithFileGeneration<AdvancedParallelPerformanceTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<AdvancedParallelPerformanceTest>::SetUp();

        // 启用内存泄漏检测
        auto& detector = TXMemoryLeakDetector::instance();
        detector.reset();
        detector.startAutoCleanup();
    }

    void TearDown() override {
        // 检查内存泄漏
        auto& detector = TXMemoryLeakDetector::instance();
        auto report = detector.detectLeaks();
        if (report.leakedAllocations > 0) {
            std::cout << "⚠️ 检测到内存泄漏: " << report.leakedAllocations
                      << " 个分配, " << report.totalLeakedBytes << " 字节" << std::endl;
        }
        detector.stopAutoCleanup();

        TestWithFileGeneration<AdvancedParallelPerformanceTest>::TearDown();
    }
    
    // 生成无重复的随机测试数据
    std::vector<std::pair<TXCoordinate, cell_value_t>> generateCellData(size_t count) {
        std::set<std::pair<u32, u32>> usedCoords;
        std::vector<std::pair<TXCoordinate, cell_value_t>> data;
        data.reserve(count);

        std::random_device rd;
        std::mt19937 gen(rd());
        // 扩大范围以确保有足够的唯一坐标
        std::uniform_int_distribution<> rowDist(1, 2000);
        std::uniform_int_distribution<> colDist(1, 100);
        std::uniform_int_distribution<> valueDist(1, 100000);
        std::uniform_int_distribution<> typeChoice(0, 3); // 0=double, 1=string, 2=int, 3=bool

        std::cout << "生成 " << count << " 个无重复随机单元格..." << std::endl;

        while (data.size() < count) {
            u32 row = rowDist(gen);
            u32 col = colDist(gen);

            // 确保坐标唯一
            if (usedCoords.insert({row, col}).second) {
                TXCoordinate coord(row_t(row), column_t(col));

                // 生成不同类型的数据，增加测试复杂度
                cell_value_t value;
                switch (typeChoice(gen)) {
                    case 0: // 浮点数
                        value = static_cast<double>(valueDist(gen)) / 100.0;
                        break;
                    case 1: // 字符串
                        value = "Test_" + std::to_string(valueDist(gen));
                        break;
                    case 2: // 整数
                        value = static_cast<i64>(valueDist(gen));
                        break;
                    case 3: // 布尔值
                        value = (valueDist(gen) % 2 == 0);
                        break;
                }

                data.emplace_back(coord, value);
            }
        }

        std::cout << "✅ 成功生成 " << data.size() << " 个唯一单元格数据" << std::endl;
        return data;
    }
};

/**
 * @brief 测试简单的并行任务处理
 */
TEST_F(AdvancedParallelPerformanceTest, SimpleParallelTaskPerformance) {
    std::cout << "\n🚀 测试简单并行任务性能..." << std::endl;

    const size_t numTasks = 1000;
    const size_t taskComplexity = 1000;

    std::cout << "开始并行任务性能测试 - " << numTasks << " 任务" << std::endl;

    auto startTime = std::chrono::high_resolution_clock::now();

    // 使用简单的多线程处理
    std::vector<std::thread> workers;
    std::vector<size_t> results(numTasks);

    size_t numThreads = std::thread::hardware_concurrency();
    size_t tasksPerThread = numTasks / numThreads;

    for (size_t t = 0; t < numThreads; ++t) {
        workers.emplace_back([&, t]() {
            size_t start = t * tasksPerThread;
            size_t end = (t == numThreads - 1) ? numTasks : (t + 1) * tasksPerThread;

            for (size_t i = start; i < end; ++i) {
                // 模拟计算密集型任务
                size_t result = 0;
                for (size_t j = 0; j < taskComplexity; ++j) {
                    result += (i * j) % 1000;
                }
                results[i] = result;
            }
        });
    }

    // 等待所有线程完成
    for (auto& worker : workers) {
        worker.join();
    }

    auto endTime = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(endTime - startTime);

    // 计算总结果
    size_t totalResult = 0;
    for (size_t result : results) {
        totalResult += result;
    }

    std::cout << "✅ 任务总数: " << numTasks << std::endl;
    std::cout << "✅ 总耗时: " << duration.count() << " μs" << std::endl;
    std::cout << "✅ 平均每任务: " << (duration.count() / numTasks) << " μs" << std::endl;
    std::cout << "✅ 总结果: " << totalResult << std::endl;

    // 验证结果
    EXPECT_EQ(results.size(), numTasks);
    EXPECT_GT(totalResult, 0);
}

/**
 * @brief 测试智能并行单元格处理器
 */
TEST_F(AdvancedParallelPerformanceTest, SmartParallelCellProcessorPerformance) {
    std::cout << "\n🚀 测试智能并行单元格处理器..." << std::endl;
    
    TXWorkbook workbook;
    auto sheet = workbook.addSheet("PerformanceTest");
    ASSERT_NE(sheet, nullptr);
    
    // 生成大量测试数据
    const size_t cellCount = 100000;
    auto cellData = generateCellData(cellCount);
    
    std::cout << "生成了 " << cellCount << " 个单元格数据" << std::endl;
    
    // 配置智能处理器
    TXSmartParallelCellProcessor::ProcessorConfig config;
    config.numThreads = std::thread::hardware_concurrency();
    config.enableAdaptiveBatching = true;
    config.enableMemoryPool = true;
    config.enableCacheOptimization = true;
    
    TXSmartParallelCellProcessor processor(config);
    
    {
        std::cout << "开始智能并行单元格处理测试 - " << cellCount << " 单元格" << std::endl;
        
        auto startTime = std::chrono::high_resolution_clock::now();
        
        // 使用智能并行处理器
        auto result = processor.parallelSetCellValues(*sheet, cellData);
        
        auto endTime = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(endTime - startTime);
        
        ASSERT_TRUE(result.isOk());
        
        size_t processedCount = result.value();
        std::cout << "✅ 处理单元格数: " << processedCount << std::endl;
        std::cout << "✅ 总耗时: " << duration.count() << " μs" << std::endl;
        std::cout << "✅ 平均每单元格: " << (duration.count() / processedCount) << " μs" << std::endl;
        std::cout << "✅ 处理速度: " << (processedCount * 1000000 / duration.count()) << " 单元格/秒" << std::endl;
    }
    
    // 验证数据正确性
    std::cout << "验证数据正确性..." << std::endl;
    size_t verifiedCount = 0;
    for (const auto& [coord, cellValue] : cellData) {
        auto retrievedValue = sheet->getCellValue(coord.getRow(), coord.getCol());
        // 检查是否不是 std::monostate（空值）
        if (!std::holds_alternative<std::monostate>(retrievedValue)) {
            verifiedCount++;
        }
    }
    std::cout << "✅ 验证通过的单元格数: " << verifiedCount << std::endl;
}

/**
 * @brief 测试XLSX任务调度器
 */
TEST_F(AdvancedParallelPerformanceTest, XlsxTaskSchedulerPerformance) {
    std::cout << "\n🚀 测试XLSX任务调度器..." << std::endl;
    
    // 配置调度器
    TXXlsxTaskScheduler::SchedulerConfig config;
    config.maxConcurrentTasks = std::thread::hardware_concurrency();
    config.enableDependencyTracking = true;
    config.enableResourceMonitoring = true;
    config.enableAdaptiveScheduling = true;
    
    TXXlsxTaskScheduler scheduler(config);
    
    const size_t numTasks = 1000;
    std::vector<std::future<void>> futures;
    futures.reserve(numTasks);
    
    {
        std::cout << "开始XLSX任务调度器测试 - " << numTasks << " 任务" << std::endl;
        
        auto startTime = std::chrono::high_resolution_clock::now();
        
        // 提交不同类型的任务
        for (size_t i = 0; i < numTasks; ++i) {
            TXXlsxTaskScheduler::TaskType taskType;
            size_t estimatedMemory;
            std::chrono::microseconds estimatedTime;
            
            // 根据任务索引分配不同类型
            switch (i % 5) {
                case 0:
                    taskType = TXXlsxTaskScheduler::TaskType::CellProcessing;
                    estimatedMemory = 1024;
                    estimatedTime = std::chrono::microseconds(100);
                    break;
                case 1:
                    taskType = TXXlsxTaskScheduler::TaskType::XmlGeneration;
                    estimatedMemory = 4096;
                    estimatedTime = std::chrono::microseconds(500);
                    break;
                case 2:
                    taskType = TXXlsxTaskScheduler::TaskType::Compression;
                    estimatedMemory = 8192;
                    estimatedTime = std::chrono::microseconds(1000);
                    break;
                case 3:
                    taskType = TXXlsxTaskScheduler::TaskType::IO;
                    estimatedMemory = 2048;
                    estimatedTime = std::chrono::microseconds(2000);
                    break;
                default:
                    taskType = TXXlsxTaskScheduler::TaskType::StringProcessing;
                    estimatedMemory = 512;
                    estimatedTime = std::chrono::microseconds(50);
                    break;
            }
            
            TXXlsxTaskScheduler::TaskMetrics metrics(taskType, estimatedMemory, estimatedTime);
            
            auto future = scheduler.scheduleTask(metrics, [i]() {
                // 模拟任务执行
                std::this_thread::sleep_for(std::chrono::microseconds(10 + (i % 100)));
            });
            
            futures.push_back(std::move(future));
        }
        
        // 等待所有任务完成
        for (auto& future : futures) {
            future.get();
        }
        
        auto endTime = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(endTime - startTime);
        
        std::cout << "✅ 调度任务数: " << numTasks << std::endl;
        std::cout << "✅ 总耗时: " << duration.count() << " ms" << std::endl;
        std::cout << "✅ 平均每任务: " << (duration.count() * 1000 / numTasks) << " μs" << std::endl;
        
        // 获取调度器统计
        auto stats = scheduler.getStats();
        std::cout << "✅ 已完成任务: " << stats.tasksCompleted << std::endl;
        std::cout << "✅ 队列中任务: " << stats.tasksInQueue << std::endl;
        std::cout << "✅ 当前内存使用: " << stats.currentMemoryUsage << " bytes" << std::endl;
    }
}

/**
 * @brief 并行框架综合性能测试
 */
TEST_F(AdvancedParallelPerformanceTest, ComprehensiveParallelPerformance) {
    std::cout << "\n🚀 并行框架综合性能测试..." << std::endl;
    
    TXWorkbook workbook;
    
    // 创建多个工作表
    const size_t numSheets = 5;
    const size_t cellsPerSheet = 10000;
    
    std::vector<TXSheet*> sheets;
    for (size_t i = 0; i < numSheets; ++i) {
        auto sheet = workbook.addSheet("Sheet" + std::to_string(i + 1));
        ASSERT_NE(sheet, nullptr);
        sheets.push_back(sheet);
    }
    
    {
        std::cout << "开始综合并行测试 - " << numSheets << " 工作表" << std::endl;
        
        // 配置并行处理器
        TXSmartParallelCellProcessor::ProcessorConfig config;
        config.numThreads = std::thread::hardware_concurrency();
        config.enableAdaptiveBatching = true;
        config.enableMemoryPool = true;
        
        TXSmartParallelCellProcessor processor(config);
        
        // 并行处理所有工作表
        std::vector<std::future<TXResult<size_t>>> futures;
        
        for (size_t i = 0; i < numSheets; ++i) {
            auto cellData = generateCellData(cellsPerSheet);
            
            auto future = std::async(std::launch::async, [&processor, sheet = sheets[i], cellData]() {
                return processor.parallelSetCellValues(*sheet, cellData);
            });
            
            futures.push_back(std::move(future));
        }
        
        // 收集结果
        size_t totalProcessed = 0;
        for (auto& future : futures) {
            auto result = future.get();
            ASSERT_TRUE(result.isOk());
            totalProcessed += result.value();
        }
        
        std::cout << "✅ 总处理单元格数: " << totalProcessed << std::endl;
        std::cout << "✅ 工作表数量: " << numSheets << std::endl;
        std::cout << "✅ 平均每工作表: " << (totalProcessed / numSheets) << " 单元格" << std::endl;
    }
    
    // 添加测试信息到第一个工作表
    addTestInfo(sheets[0], "ComprehensiveParallelPerformance", "综合并行性能测试 - " + std::to_string(numSheets) + " 工作表");

    // 保存文件测试
    {
        std::cout << "开始保存综合测试文件" << std::endl;
        std::cout << "工作表数量: " << workbook.getSheetCount() << std::endl;

        // 验证工作表状态
        for (size_t i = 0; i < workbook.getSheetCount(); ++i) {
            auto* sheet = workbook.getSheet(i);
            if (sheet) {
                std::cout << "工作表 " << i << ": " << sheet->getName()
                         << ", 单元格数: " << sheet->getCellManager().getCellCount() << std::endl;
            }
        }

        std::string fullPath = getFilePath("ComprehensiveParallelTest");
        bool success = workbook.saveToFile(fullPath);
        EXPECT_TRUE(success);

        if (!success) {
            std::cout << "❌ 保存失败: " << workbook.getLastError() << std::endl;
        } else {
            std::cout << "✅ 文件保存成功: " << fullPath << std::endl;
        }
    }
}

int main(int argc, char** argv) {
    ::testing::InitGoogleTest(&argc, argv);
    
    std::cout << "🚀 高级并行框架性能测试开始..." << std::endl;
    std::cout << "CPU核心数: " << std::thread::hardware_concurrency() << std::endl;
    
    return RUN_ALL_TESTS();
}
