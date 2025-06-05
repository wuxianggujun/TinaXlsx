//
// @file test_batch_pipeline.cpp
// @brief 批处理流水线测试 - 第3周交付物验证
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXBatchPipeline.hpp"
#include "TinaXlsx/TXBatchXMLGenerator.hpp"
#include "TinaXlsx/TXAsyncProcessingFramework.hpp"
#include "TinaXlsx/TXUnifiedMemoryManager.hpp"
#include "test_file_generator.hpp"
#include <chrono>
#include <random>
#include <iostream>

using namespace TinaXlsx;

class BatchPipelineTest : public TestWithFileGeneration<BatchPipelineTest> {
protected:
    void SetUp() override {
        TestWithFileGeneration<BatchPipelineTest>::SetUp();
        
        // 配置统一内存管理器
        TXUnifiedMemoryManager::Config memory_config;
        memory_config.chunk_size = 32 * 1024 * 1024;      // 32MB块
        memory_config.memory_limit = 1024 * 1024 * 1024;  // 1GB限制
        memory_config.enable_monitoring = true;
        
        memory_manager_ = std::make_unique<TXUnifiedMemoryManager>(memory_config);
        
        // 配置批处理流水线
        TXBatchPipeline::PipelineConfig pipeline_config;
        pipeline_config.max_concurrent_batches = 8;
        pipeline_config.batch_size_threshold = 1000;
        pipeline_config.memory_limit_mb = 512;
        pipeline_config.enable_memory_optimization = true;
        pipeline_config.enable_async_processing = true;
        pipeline_config.enable_performance_monitoring = true;
        
        pipeline_ = std::make_unique<TXBatchPipeline>(pipeline_config);
        
        // 配置XML生成器
        TXBatchXMLGenerator::XMLGeneratorConfig xml_config;
        xml_config.enable_memory_pooling = true;
        xml_config.enable_parallel_generation = true;
        xml_config.batch_size = 5000;
        
        xml_generator_ = std::make_unique<TXBatchXMLGenerator>(*memory_manager_, xml_config);
        
        // 配置异步处理框架
        TXAsyncProcessingFramework::FrameworkConfig async_config;
        async_config.worker_thread_count = 4;
        async_config.enable_work_stealing = true;
        async_config.enable_priority_scheduling = true;
        async_config.memory_limit_mb = 256;
        
        async_framework_ = std::make_unique<TXAsyncProcessingFramework>(*memory_manager_, async_config);
    }
    
    void TearDown() override {
        async_framework_.reset();
        xml_generator_.reset();
        pipeline_.reset();
        memory_manager_.reset();
        TestWithFileGeneration<BatchPipelineTest>::TearDown();
    }
    
    /**
     * @brief 创建测试批次数据
     */
    std::unique_ptr<TXBatchData> createTestBatch(size_t batch_id, size_t cell_count) {
        auto batch = std::make_unique<TXBatchData>(batch_id);
        
        std::random_device rd;
        std::mt19937 gen(rd());
        std::uniform_int_distribution<> value_dist(1, 10000);
        std::uniform_int_distribution<> type_dist(0, 2);
        
        batch->cells.reserve(cell_count);
        
        for (size_t i = 0; i < cell_count; ++i) {
            TXCompactCell cell;
            
            switch (type_dist(gen)) {
                case 0: // 数字
                    cell.setValue(static_cast<double>(value_dist(gen)));
                    break;
                case 1: // 字符串
                    cell.setValue("TestString_" + std::to_string(value_dist(gen)));
                    break;
                case 2: // 布尔
                    cell.setValue(value_dist(gen) % 2 == 0);
                    break;
            }
            
            batch->cells.push_back(cell);
        }
        
        batch->estimated_size = cell_count * sizeof(TXCompactCell);
        return batch;
    }
    
    std::unique_ptr<TXUnifiedMemoryManager> memory_manager_;
    std::unique_ptr<TXBatchPipeline> pipeline_;
    std::unique_ptr<TXBatchXMLGenerator> xml_generator_;
    std::unique_ptr<TXAsyncProcessingFramework> async_framework_;
};

// ==================== 批处理流水线测试 ====================

TEST_F(BatchPipelineTest, PipelineBasicFunctionalityTest) {
    std::cout << "\n=== 批处理流水线基础功能测试 ===" << std::endl;
    
    // 启动流水线
    auto start_result = pipeline_->start();
    ASSERT_TRUE(start_result.isOk()) << "流水线启动失败: " << start_result.error().getMessage();
    
    std::cout << "✅ 流水线启动成功" << std::endl;
    
    // 创建测试批次
    std::vector<std::unique_ptr<TXBatchData>> test_batches;
    const size_t BATCH_COUNT = 5;
    const size_t CELLS_PER_BATCH = 1000;
    
    for (size_t i = 0; i < BATCH_COUNT; ++i) {
        test_batches.push_back(createTestBatch(i + 1, CELLS_PER_BATCH));
    }
    
    std::cout << "创建了 " << BATCH_COUNT << " 个测试批次，每批 " << CELLS_PER_BATCH << " 个单元格" << std::endl;
    
    // 提交批次
    std::vector<size_t> batch_ids;
    for (auto& batch : test_batches) {
        auto submit_result = pipeline_->submitBatch(std::move(batch));
        ASSERT_TRUE(submit_result.isOk()) << "批次提交失败: " << submit_result.error().getMessage();
        batch_ids.push_back(submit_result.value());
    }
    
    std::cout << "✅ 所有批次提交成功" << std::endl;
    
    // 等待处理完成
    auto wait_result = pipeline_->waitForCompletion(std::chrono::seconds(30));
    ASSERT_TRUE(wait_result.isOk()) << "等待完成失败: " << wait_result.error().getMessage();
    
    std::cout << "✅ 所有批次处理完成" << std::endl;
    
    // 获取统计信息
    auto stats = pipeline_->getStats();
    std::cout << "\n流水线统计:" << std::endl;
    std::cout << "  处理批次: " << stats.total_batches_processed << std::endl;
    std::cout << "  失败批次: " << stats.total_batches_failed << std::endl;
    std::cout << "  平均处理时间: " << stats.avg_pipeline_time.count() << " μs" << std::endl;
    std::cout << "  整体吞吐量: " << stats.overall_throughput << " 批次/秒" << std::endl;
    std::cout << "  内存效率: " << (stats.memory_efficiency * 100) << "%" << std::endl;
    
    // 验证处理结果
    EXPECT_EQ(stats.total_batches_processed, BATCH_COUNT) << "处理批次数不匹配";
    EXPECT_EQ(stats.total_batches_failed, 0) << "不应该有失败批次";
    EXPECT_GT(stats.overall_throughput, 0) << "吞吐量应该大于0";
    
    // 停止流水线
    auto stop_result = pipeline_->stop();
    ASSERT_TRUE(stop_result.isOk()) << "流水线停止失败: " << stop_result.error().getMessage();
    
    std::cout << "✅ 流水线停止成功" << std::endl;
    std::cout << "✅ 批处理流水线基础功能测试完成" << std::endl;
}

TEST_F(BatchPipelineTest, PipelinePerformanceTest) {
    std::cout << "\n=== 批处理流水线性能测试 ===" << std::endl;
    
    // 启动流水线
    (void)pipeline_->start();
    
    const size_t LARGE_BATCH_COUNT = 20;
    const size_t CELLS_PER_BATCH = 5000;
    
    std::cout << "性能测试配置:" << std::endl;
    std::cout << "  批次数量: " << LARGE_BATCH_COUNT << std::endl;
    std::cout << "  每批单元格: " << CELLS_PER_BATCH << std::endl;
    std::cout << "  总单元格数: " << (LARGE_BATCH_COUNT * CELLS_PER_BATCH) << std::endl;
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 创建并提交大量批次
    std::vector<size_t> batch_ids;
    for (size_t i = 0; i < LARGE_BATCH_COUNT; ++i) {
        auto batch = createTestBatch(i + 1, CELLS_PER_BATCH);
        auto result = pipeline_->submitBatch(std::move(batch));
        ASSERT_TRUE(result.isOk());
        batch_ids.push_back(result.value());
        
        if ((i + 1) % 5 == 0) {
            std::cout << "  已提交 " << (i + 1) << " 个批次" << std::endl;
        }
    }
    
    auto submit_time = std::chrono::high_resolution_clock::now();
    auto submit_duration = std::chrono::duration_cast<std::chrono::milliseconds>(submit_time - start_time);
    
    std::cout << "批次提交完成，耗时: " << submit_duration.count() << " ms" << std::endl;
    
    // 等待处理完成
    (void)pipeline_->waitForCompletion(std::chrono::seconds(60));
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto total_duration = std::chrono::duration_cast<std::chrono::milliseconds>(end_time - start_time);
    
    // 获取性能统计
    auto stats = pipeline_->getStats();
    
    std::cout << "\n性能测试结果:" << std::endl;
    std::cout << "  总处理时间: " << total_duration.count() << " ms" << std::endl;
    std::cout << "  处理批次: " << stats.total_batches_processed << std::endl;
    std::cout << "  平均批次处理时间: " << stats.avg_pipeline_time.count() << " μs" << std::endl;
    std::cout << "  整体吞吐量: " << stats.overall_throughput << " 批次/秒" << std::endl;
    std::cout << "  单元格处理速率: " << (stats.total_batches_processed * CELLS_PER_BATCH * 1000.0 / total_duration.count()) << " 单元格/秒" << std::endl;
    
    // 性能验证
    EXPECT_GT(stats.overall_throughput, 1.0) << "吞吐量应该大于1批次/秒";
    EXPECT_LT(stats.avg_pipeline_time.count(), 100000) << "平均处理时间应该小于100ms";
    
    (void)pipeline_->stop();
    
    std::cout << "✅ 批处理流水线性能测试完成" << std::endl;
}

// ==================== XML生成器测试 ====================

TEST_F(BatchPipelineTest, XMLGeneratorTest) {
    std::cout << "\n=== XML批量生成器测试 ===" << std::endl;
    
    // 创建测试单元格
    std::vector<TXCompactCell> test_cells;
    test_cells.reserve(1000);
    
    for (size_t i = 0; i < 1000; ++i) {
        TXCompactCell cell;
        if (i % 3 == 0) {
            cell.setValue(static_cast<double>(i * 1.5));
        } else if (i % 3 == 1) {
            cell.setValue("TestString_" + std::to_string(i));
        } else {
            cell.setValue(i % 2 == 0);
        }
        test_cells.push_back(cell);
    }
    
    std::cout << "创建了 " << test_cells.size() << " 个测试单元格" << std::endl;
    
    // 测试单个单元格XML生成
    auto single_result = xml_generator_->generateCellXML(test_cells[0]);
    ASSERT_TRUE(single_result.isOk()) << "单个单元格XML生成失败";

    std::cout << "✅ 单个单元格XML生成成功" << std::endl;
    std::cout << "示例XML: " << single_result.value().substr(0, 100) << "..." << std::endl;
    
    // 测试批量XML生成
    auto start_time = std::chrono::high_resolution_clock::now();
    
    auto batch_result = xml_generator_->generateCellsXML(test_cells);
    ASSERT_TRUE(batch_result.isOk()) << "批量XML生成失败";
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    
    std::cout << "✅ 批量XML生成成功" << std::endl;
    std::cout << "生成时间: " << duration.count() << " μs" << std::endl;
    std::cout << "生成速率: " << (test_cells.size() * 1000000.0 / duration.count()) << " 单元格/秒" << std::endl;
    std::cout << "XML大小: " << batch_result.value().size() << " 字节" << std::endl;
    
    // 获取XML生成器统计
    auto xml_stats = xml_generator_->getStats();
    std::cout << "\nXML生成器统计:" << std::endl;
    std::cout << "  总生成XML数: " << xml_stats.total_xml_generated << std::endl;
    std::cout << "  总处理单元格: " << xml_stats.total_cells_processed << std::endl;
    std::cout << "  总生成字节数: " << xml_stats.total_bytes_generated << std::endl;
    std::cout << "  平均生成时间: " << xml_stats.avg_generation_time.count() << " μs" << std::endl;
    std::cout << "  生成速率: " << xml_stats.generation_rate << " 单元格/秒" << std::endl;
    std::cout << "  内存效率: " << (xml_stats.memory_efficiency * 100) << "%" << std::endl;
    
    // 性能验证
    EXPECT_GT(xml_stats.generation_rate, 10000) << "XML生成速率应该大于1万单元格/秒";
    EXPECT_LT(xml_stats.avg_generation_time.count(), 10) << "平均生成时间应该小于10μs";
    
    std::cout << "✅ XML批量生成器测试完成" << std::endl;
}

// ==================== 异步处理框架测试 ====================

TEST_F(BatchPipelineTest, AsyncFrameworkTest) {
    std::cout << "\n=== 异步处理框架测试 ===" << std::endl;
    
    // 启动异步框架
    auto start_result = async_framework_->start();
    ASSERT_TRUE(start_result.isOk()) << "异步框架启动失败";
    
    std::cout << "✅ 异步框架启动成功" << std::endl;
    
    // 测试函数任务提交
    const size_t TASK_COUNT = 100;
    std::vector<std::future<int>> futures;
    
    std::cout << "提交 " << TASK_COUNT << " 个计算任务..." << std::endl;
    
    auto submit_start = std::chrono::high_resolution_clock::now();
    
    for (size_t i = 0; i < TASK_COUNT; ++i) {
        auto result = async_framework_->submitFunction([i]() -> int {
            // 模拟计算工作
            int sum = 0;
            for (int j = 0; j < 1000; ++j) {
                sum += (i * j) % 1000;
            }
            return sum;
        }, "ComputeTask_" + std::to_string(i));
        
        ASSERT_TRUE(result.isOk()) << "任务提交失败";
        futures.push_back(std::move(result.value()));
    }
    
    auto submit_end = std::chrono::high_resolution_clock::now();
    auto submit_duration = std::chrono::duration_cast<std::chrono::microseconds>(submit_end - submit_start);
    
    std::cout << "任务提交完成，耗时: " << submit_duration.count() << " μs" << std::endl;
    
    // 等待所有任务完成
    auto wait_start = std::chrono::high_resolution_clock::now();
    
    size_t completed_tasks = 0;
    for (auto& future : futures) {
        try {
            int result = future.get();
            completed_tasks++;
            (void)result; // 避免未使用变量警告
        } catch (const std::exception& e) {
            std::cout << "任务执行异常: " << e.what() << std::endl;
        }
    }
    
    auto wait_end = std::chrono::high_resolution_clock::now();
    auto wait_duration = std::chrono::duration_cast<std::chrono::milliseconds>(wait_end - wait_start);
    
    std::cout << "任务执行完成，耗时: " << wait_duration.count() << " ms" << std::endl;
    std::cout << "完成任务数: " << completed_tasks << "/" << TASK_COUNT << std::endl;
    
    // 获取异步框架统计
    auto async_stats = async_framework_->getStats();
    std::cout << "\n异步框架统计:" << std::endl;
    std::cout << "  总提交任务: " << async_stats.total_tasks_submitted << std::endl;
    std::cout << "  总完成任务: " << async_stats.total_tasks_completed << std::endl;
    std::cout << "  总失败任务: " << async_stats.total_tasks_failed << std::endl;
    std::cout << "  平均执行时间: " << async_stats.avg_execution_time.count() << " μs" << std::endl;
    std::cout << "  任务处理速率: " << async_stats.tasks_per_second << " 任务/秒" << std::endl;
    std::cout << "  活跃工作线程: " << async_stats.active_worker_threads << std::endl;
    
    // 验证结果
    EXPECT_EQ(completed_tasks, TASK_COUNT) << "所有任务都应该完成";
    EXPECT_EQ(async_stats.total_tasks_failed, 0) << "不应该有失败任务";
    EXPECT_GT(async_stats.tasks_per_second, 100) << "任务处理速率应该大于100任务/秒";
    
    // 停止异步框架
    auto stop_result = async_framework_->stop();
    ASSERT_TRUE(stop_result.isOk()) << "异步框架停止失败";
    
    std::cout << "✅ 异步处理框架测试完成" << std::endl;
}
