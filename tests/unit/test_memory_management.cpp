//
// @file test_memory_management.cpp
// @brief 内存管理系统测试
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXChunkAllocator.hpp"
#include "TinaXlsx/TXSmartMemoryManager.hpp"
#include "TinaXlsx/TXOptimizedSIMD.hpp"
#include <chrono>
#include <thread>
#include <iostream>
#include <numeric>

using namespace TinaXlsx;

class MemoryManagementTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 创建测试用的内存分配器
        allocator_ = std::make_unique<TXChunkAllocator>();
        
        // 设置较小的内存限制用于测试
        allocator_->setMemoryLimit(100 * 1024 * 1024); // 100MB
        allocator_->setChunkSize(10 * 1024 * 1024);    // 10MB块
    }
    
    void TearDown() override {
        allocator_.reset();
    }
    
    std::unique_ptr<TXChunkAllocator> allocator_;
};

// ==================== TXChunkAllocator 测试 ====================

TEST_F(MemoryManagementTest, ChunkAllocatorBasicAllocation) {
    std::cout << "\n=== 分块分配器基础测试 ===" << std::endl;
    
    // 测试基本分配
    void* ptr1 = allocator_->allocate(1024);
    ASSERT_NE(ptr1, nullptr);
    
    void* ptr2 = allocator_->allocate(2048);
    ASSERT_NE(ptr2, nullptr);
    
    // 检查内存使用
    size_t usage = allocator_->getTotalMemoryUsage();
    EXPECT_GT(usage, 0);
    
    std::cout << "分配后内存使用: " << (usage / 1024.0) << " KB" << std::endl;
    
    // 获取统计信息
    auto stats = allocator_->getStats();
    std::cout << "分配统计:" << std::endl;
    std::cout << "  总分配: " << stats.allocation_count << " 次" << std::endl;
    std::cout << "  失败分配: " << stats.failed_allocations << " 次" << std::endl;
    std::cout << "  活跃块数: " << stats.active_chunks << std::endl;
    
    EXPECT_EQ(stats.failed_allocations, 0);
    EXPECT_GT(stats.allocation_count, 0);
}

TEST_F(MemoryManagementTest, ChunkAllocatorBatchAllocation) {
    std::cout << "\n=== 批量分配测试（智能块大小）===" << std::endl;

    // 准备不同大小的分配请求来测试智能块选择
    std::vector<size_t> small_sizes = {1024, 2048, 4096, 8192, 1024, 512}; // 小分配
    std::vector<size_t> medium_sizes = {128 * 1024, 256 * 1024, 512 * 1024}; // 中等分配
    std::vector<size_t> large_sizes = {5 * 1024 * 1024}; // 大分配

    std::cout << "测试小分配（应使用1MB块）:" << std::endl;
    auto start = std::chrono::high_resolution_clock::now();
    auto small_ptrs = allocator_->allocateBatch(small_sizes);
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);

    size_t small_total = std::accumulate(small_sizes.begin(), small_sizes.end(), size_t(0));
    size_t usage_after_small = allocator_->getTotalMemoryUsage();

    std::cout << "  小分配时间: " << duration.count() << " 微秒" << std::endl;
    std::cout << "  请求总量: " << small_total << " 字节" << std::endl;
    std::cout << "  实际使用: " << usage_after_small << " 字节" << std::endl;
    std::cout << "  内存效率: " << (static_cast<double>(small_total) / usage_after_small * 100) << "%" << std::endl;

    // 验证小分配使用了合适的块大小（应该是1MB块）
    auto chunk_infos = allocator_->getChunkInfos();
    bool has_small_chunk = false;
    for (const auto& info : chunk_infos) {
        if (info.total_size == 1024 * 1024) { // 1MB块
            has_small_chunk = true;
            std::cout << "  ✅ 使用了1MB小块，使用率: " << (info.usage_ratio * 100) << "%" << std::endl;
        }
    }
    EXPECT_TRUE(has_small_chunk) << "应该创建1MB小块用于小分配";

    std::cout << "\n测试中等分配（应使用16MB块）:" << std::endl;
    start = std::chrono::high_resolution_clock::now();
    auto medium_ptrs = allocator_->allocateBatch(medium_sizes);
    end = std::chrono::high_resolution_clock::now();
    duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);

    size_t medium_total = std::accumulate(medium_sizes.begin(), medium_sizes.end(), size_t(0));
    size_t usage_after_medium = allocator_->getTotalMemoryUsage();

    std::cout << "  中等分配时间: " << duration.count() << " 微秒" << std::endl;
    std::cout << "  请求总量: " << medium_total << " 字节" << std::endl;
    std::cout << "  新增使用: " << (usage_after_medium - usage_after_small) << " 字节" << std::endl;
    std::cout << "  中等分配效率: " << (static_cast<double>(medium_total) / (usage_after_medium - usage_after_small) * 100) << "%" << std::endl;

    // 验证所有分配都成功
    for (size_t i = 0; i < small_sizes.size(); ++i) {
        EXPECT_NE(small_ptrs[i], nullptr) << "小分配 " << i << " 失败";
    }
    for (size_t i = 0; i < medium_sizes.size(); ++i) {
        EXPECT_NE(medium_ptrs[i], nullptr) << "中等分配 " << i << " 失败";
    }

    // 总体效率应该比之前好很多
    double overall_efficiency = static_cast<double>(small_total + medium_total) / usage_after_medium * 100;
    std::cout << "  总体内存效率: " << overall_efficiency << "%" << std::endl;
    EXPECT_GT(overall_efficiency, 50.0) << "智能块选择应该提供>50%的内存效率";
}

TEST_F(MemoryManagementTest, ChunkAllocatorMemoryLimit) {
    std::cout << "\n=== 内存限制测试 ===" << std::endl;
    
    size_t limit = allocator_->getMemoryLimit();
    std::cout << "内存限制: " << (limit / 1024.0 / 1024.0) << " MB" << std::endl;
    
    // 尝试分配接近限制的内存 - 修复：考虑块大小开销
    size_t chunk_size = allocator_->getChunkSize();
    std::cout << "块大小: " << (chunk_size / 1024.0 / 1024.0) << " MB" << std::endl;

    // 分配小于块大小的内存应该成功
    size_t safe_size = chunk_size / 4; // 2.5MB
    void* ptr1 = allocator_->allocate(safe_size);
    EXPECT_NE(ptr1, nullptr) << "第一次分配应该成功";

    void* ptr2 = allocator_->allocate(safe_size);
    EXPECT_NE(ptr2, nullptr) << "第二次分配应该成功";

    // 尝试分配超出限制的内存 - 分配很多块直到失败
    std::vector<void*> ptrs;
    void* ptr = nullptr;
    do {
        ptr = allocator_->allocate(chunk_size / 2);
        if (ptr) ptrs.push_back(ptr);
    } while (ptr != nullptr && ptrs.size() < 20); // 最多尝试20次

    EXPECT_EQ(ptr, nullptr) << "最终应该因为超出限制而分配失败";
    std::cout << "成功分配了 " << ptrs.size() << " 个大块" << std::endl;
    
    auto stats = allocator_->getStats();
    std::cout << "失败分配数: " << stats.failed_allocations << std::endl;
    EXPECT_GT(stats.failed_allocations, 0);
}

TEST_F(MemoryManagementTest, ChunkAllocatorCompaction) {
    std::cout << "\n=== 内存压缩测试 ===" << std::endl;
    
    // 分配一些内存
    std::vector<void*> ptrs;
    for (int i = 0; i < 10; ++i) {
        ptrs.push_back(allocator_->allocate(1024 * 1024)); // 1MB each
    }
    
    size_t usage_before = allocator_->getTotalMemoryUsage();
    auto chunks_before = allocator_->getChunkInfos();
    
    std::cout << "压缩前:" << std::endl;
    std::cout << "  内存使用: " << (usage_before / 1024.0 / 1024.0) << " MB" << std::endl;
    std::cout << "  块数: " << chunks_before.size() << std::endl;
    
    // 清空所有内存
    allocator_->deallocateAll();
    
    // 执行压缩
    allocator_->compact();
    
    size_t usage_after = allocator_->getTotalMemoryUsage();
    auto chunks_after = allocator_->getChunkInfos();
    
    std::cout << "压缩后:" << std::endl;
    std::cout << "  内存使用: " << (usage_after / 1024.0 / 1024.0) << " MB" << std::endl;
    std::cout << "  块数: " << chunks_after.size() << std::endl;
    
    EXPECT_LE(usage_after, usage_before);
    EXPECT_LE(chunks_after.size(), chunks_before.size());
}

// ==================== TXSmartMemoryManager 测试 ====================

TEST_F(MemoryManagementTest, SmartMemoryManagerBasic) {
    std::cout << "\n=== 智能内存管理器基础测试 ===" << std::endl;
    
    // 配置内存监控
    MemoryMonitorConfig config;
    config.warning_threshold_mb = 50;   // 50MB警告
    config.critical_threshold_mb = 70;  // 70MB严重
    config.emergency_threshold_mb = 90; // 90MB紧急
    config.monitor_interval = std::chrono::milliseconds(100); // 100ms间隔
    
    TXSmartMemoryManager manager(*allocator_, config);
    
    // 设置事件回调
    std::vector<MemoryEvent> captured_events;
    manager.setEventCallback([&captured_events](const MemoryEvent& event) {
        captured_events.push_back(event);
    });
    
    // 启动监控
    manager.startMonitoring();
    
    // 分配内存触发警告
    std::vector<void*> ptrs;
    for (int i = 0; i < 60; ++i) { // 分配60MB
        ptrs.push_back(allocator_->allocate(1024 * 1024));
        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    }
    
    // 等待监控检测
    std::this_thread::sleep_for(std::chrono::milliseconds(500));
    
    // 停止监控
    manager.stopMonitoring();
    
    // 检查事件
    std::cout << "捕获的事件数: " << captured_events.size() << std::endl;
    
    bool has_warning = false;
    for (const auto& event : captured_events) {
        if (event.type == MemoryEventType::WARNING) {
            has_warning = true;
            std::cout << "警告事件: " << event.message << std::endl;
        }
    }
    
    EXPECT_TRUE(has_warning) << "应该触发内存警告事件";
    
    // 获取统计信息
    auto stats = manager.getStats();
    std::cout << "监控统计:" << std::endl;
    std::cout << "  总事件: " << stats.total_events << std::endl;
    std::cout << "  警告事件: " << stats.warning_events << std::endl;
    std::cout << "  当前内存: " << stats.current_memory_usage << " MB" << std::endl;
    std::cout << "  峰值内存: " << stats.peak_memory_usage << " MB" << std::endl;
}

TEST_F(MemoryManagementTest, SmartMemoryManagerAutoCleanup) {
    std::cout << "\n=== 自动清理测试 ===" << std::endl;
    
    // 配置自动清理
    MemoryMonitorConfig config;
    config.warning_threshold_mb = 30;
    config.critical_threshold_mb = 50;
    config.emergency_threshold_mb = 70;
    config.enable_auto_cleanup = true;
    config.monitor_interval = std::chrono::milliseconds(50);
    
    TXSmartMemoryManager manager(*allocator_, config);
    
    // 记录清理事件
    std::vector<MemoryEvent> cleanup_events;
    manager.setEventCallback([&cleanup_events](const MemoryEvent& event) {
        if (event.type == MemoryEventType::CLEANUP_START || 
            event.type == MemoryEventType::CLEANUP_END) {
            cleanup_events.push_back(event);
        }
    });
    
    manager.startMonitoring();
    
    // 分配大量内存触发自动清理
    std::vector<void*> ptrs;
    for (int i = 0; i < 80; ++i) { // 分配80MB
        ptrs.push_back(allocator_->allocate(1024 * 1024));
        std::this_thread::sleep_for(std::chrono::milliseconds(5));
    }
    
    // 等待自动清理
    std::this_thread::sleep_for(std::chrono::milliseconds(1000));
    
    manager.stopMonitoring();
    
    std::cout << "清理事件数: " << cleanup_events.size() << std::endl;
    
    bool has_cleanup = false;
    for (const auto& event : cleanup_events) {
        if (event.type == MemoryEventType::CLEANUP_START) {
            has_cleanup = true;
            std::cout << "清理开始: " << event.message << std::endl;
        } else if (event.type == MemoryEventType::CLEANUP_END) {
            std::cout << "清理结束: " << event.message << std::endl;
        }
    }
    
    EXPECT_TRUE(has_cleanup) << "应该触发自动清理";
}

TEST_F(MemoryManagementTest, MemoryTrendPrediction) {
    std::cout << "\n=== 内存趋势预测测试 ===" << std::endl;
    
    MemoryMonitorConfig config;
    config.monitor_interval = std::chrono::milliseconds(50);
    
    TXSmartMemoryManager manager(*allocator_, config);
    manager.startMonitoring();
    
    // 模拟稳定增长的内存使用
    for (int i = 0; i < 20; ++i) {
        allocator_->allocate(2 * 1024 * 1024); // 每次2MB
        std::this_thread::sleep_for(std::chrono::milliseconds(100));
    }
    
    // 获取趋势预测
    auto trend = manager.predictMemoryTrend();
    
    std::cout << "内存趋势分析:" << std::endl;
    std::cout << "  增长率: " << trend.growth_rate_mb_per_sec << " MB/秒" << std::endl;
    std::cout << "  是否增长: " << (trend.is_growing ? "是" : "否") << std::endl;
    
    if (trend.is_growing) {
        std::cout << "  到达警告时间: " << trend.time_to_warning.count() << " 秒" << std::endl;
        std::cout << "  到达严重时间: " << trend.time_to_critical.count() << " 秒" << std::endl;
    }
    
    manager.stopMonitoring();
    
    EXPECT_TRUE(trend.is_growing) << "应该检测到内存增长趋势";
    EXPECT_GT(trend.growth_rate_mb_per_sec, 0) << "增长率应该大于0";
}

// ==================== 集成测试 ====================

TEST_F(MemoryManagementTest, IntegratedMemoryManagement) {
    std::cout << "\n=== 集成内存管理测试 ===" << std::endl;
    
    // 创建完整的内存管理系统
    MemoryMonitorConfig config;
    config.warning_threshold_mb = 40;
    config.critical_threshold_mb = 60;
    config.emergency_threshold_mb = 80;
    config.enable_auto_cleanup = true;
    
    TXSmartMemoryManager manager(*allocator_, config);
    
    // 记录所有事件
    std::vector<MemoryEvent> all_events;
    manager.setEventCallback([&all_events](const MemoryEvent& event) {
        all_events.push_back(event);
    });
    
    manager.startMonitoring();
    
    // 使用优化的SIMD处理器进行大量数据处理
    const size_t DATA_SIZE = 50000;
    std::vector<double> test_data(DATA_SIZE);
    for (size_t i = 0; i < DATA_SIZE; ++i) {
        test_data[i] = static_cast<double>(i) * 3.14159;
    }
    
    // 执行多轮处理，模拟实际工作负载 - 修复：分配更多内存触发监控
    for (int round = 0; round < 15; ++round) {
        std::cout << "处理轮次 " << (round + 1) << std::endl;

        // 分配更大的内存块触发监控事件
        size_t large_allocation = 5 * 1024 * 1024; // 5MB per round
        auto large_ptr = allocator_->allocate(large_allocation);

        // 分配内存用于处理
        auto cells_ptr = allocator_->allocate<UltraCompactCell>(DATA_SIZE);
        if (cells_ptr) {
            // 使用优化的SIMD处理器
            TXOptimizedSIMDProcessor::ultraFastConvertDoublesToCells(
                test_data.data(), cells_ptr, DATA_SIZE);

            // 执行一些计算
            double sum = TXOptimizedSIMDProcessor::ultraFastSumNumbers(cells_ptr, DATA_SIZE);
            EXPECT_GT(sum, 0);
        }

        // 检查当前内存使用
        size_t current_usage = allocator_->getTotalMemoryUsage() / (1024 * 1024);
        std::cout << "  当前内存使用: " << current_usage << " MB" << std::endl;

        std::this_thread::sleep_for(std::chrono::milliseconds(200));
    }
    
    // 等待监控和清理
    std::this_thread::sleep_for(std::chrono::milliseconds(1000));
    
    manager.stopMonitoring();
    
    // 分析结果
    std::cout << "\n处理完成，事件分析:" << std::endl;
    
    size_t warning_count = 0, critical_count = 0, cleanup_count = 0;
    for (const auto& event : all_events) {
        switch (event.type) {
            case MemoryEventType::WARNING:
                warning_count++;
                break;
            case MemoryEventType::CRITICAL:
                critical_count++;
                break;
            case MemoryEventType::CLEANUP_START:
                cleanup_count++;
                break;
            default:
                break;
        }
    }
    
    std::cout << "  警告事件: " << warning_count << std::endl;
    std::cout << "  严重事件: " << critical_count << std::endl;
    std::cout << "  清理事件: " << cleanup_count << std::endl;
    
    // 生成最终报告
    std::string report = manager.generateMonitoringReport();
    std::cout << "\n" << report << std::endl;
    
    // 验证系统正常工作
    EXPECT_GT(all_events.size(), 0) << "应该有监控事件";
    
    auto final_stats = manager.getStats();
    EXPECT_GT(final_stats.total_events, 0);
    EXPECT_GT(final_stats.peak_memory_usage, 0);
}

// ==================== 性能测试 ====================

TEST_F(MemoryManagementTest, MemoryManagementPerformance) {
    std::cout << "\n=== 内存管理性能测试 ===" << std::endl;
    
    const size_t ALLOCATION_COUNT = 10000;
    const size_t ALLOCATION_SIZE = 1024;
    
    // 测试分配性能
    auto start = std::chrono::high_resolution_clock::now();
    
    std::vector<void*> ptrs;
    ptrs.reserve(ALLOCATION_COUNT);
    
    for (size_t i = 0; i < ALLOCATION_COUNT; ++i) {
        void* ptr = allocator_->allocate(ALLOCATION_SIZE);
        if (ptr) {
            ptrs.push_back(ptr);
        }
    }
    
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    
    std::cout << "分配性能测试:" << std::endl;
    std::cout << "  分配数量: " << ptrs.size() << "/" << ALLOCATION_COUNT << std::endl;
    std::cout << "  总时间: " << duration.count() << " 微秒" << std::endl;
    std::cout << "  平均时间: " << (static_cast<double>(duration.count()) / ptrs.size()) << " 微秒/分配" << std::endl;
    std::cout << "  分配速率: " << (ptrs.size() * 1000000.0 / duration.count()) << " 分配/秒" << std::endl;
    
    // 性能要求验证
    double avg_time_per_allocation = static_cast<double>(duration.count()) / ptrs.size();
    EXPECT_LT(avg_time_per_allocation, 10.0) << "平均分配时间应该小于10微秒";
    
    // 内存使用效率
    size_t total_requested = ptrs.size() * ALLOCATION_SIZE;
    size_t actual_usage = allocator_->getTotalMemoryUsage();
    double efficiency = static_cast<double>(total_requested) / actual_usage;
    
    std::cout << "  内存效率: " << (efficiency * 100) << "%" << std::endl;
    EXPECT_GT(efficiency, 0.5) << "内存效率应该大于50%";
}
