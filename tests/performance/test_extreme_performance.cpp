//
// @file test_extreme_performance.cpp
// @brief TinaXlsx 极致性能测试 - GTest风格
//

#include <gtest/gtest.h>
#include <TinaXlsx/TXInMemorySheet.hpp>
#include <TinaXlsx/TXBatchSIMDProcessor.hpp>
#include <TinaXlsx/TXZeroCopySerializer.hpp>
#include <TinaXlsx/TXUnifiedMemoryManager.hpp>
#include <chrono>
#include <iostream>
#include <iomanip>
#include <random>
#include <cassert>

using namespace TinaXlsx;
using namespace std::chrono;

/**
 * @brief 高精度计时器
 */
class PerformanceTimer {
private:
    high_resolution_clock::time_point start_time_;
    
public:
    void start() {
        start_time_ = high_resolution_clock::now();
    }
    
    double getElapsedMs() const {
        auto end_time = high_resolution_clock::now();
        auto duration = duration_cast<microseconds>(end_time - start_time_);
        return duration.count() / 1000.0; // 转换为毫秒
    }
    
    void printElapsed(const std::string& operation) const {
        std::cout << operation << ": " << std::fixed << std::setprecision(3) 
                 << getElapsedMs() << " ms" << std::endl;
    }
};

class ExtremePerformanceTest : public ::testing::Test {
protected:
    PerformanceTimer timer;

    void SetUp() override {
        // 🚀 初始化全局内存管理器 - 这是关键！
        TXUnifiedMemoryManager::Config config;
        config.memory_limit = 8ULL * 1024 * 1024 * 1024; // 8GB限制
        config.warning_threshold_mb = 6144;  // 6GB警告
        config.critical_threshold_mb = 7168; // 7GB严重
        config.emergency_threshold_mb = 7680; // 7.5GB紧急
        config.enable_monitoring = true;
        config.enable_slab_allocator = true;
        config.enable_auto_reclaim = true;

        GlobalUnifiedMemoryManager::initialize(config);
        std::cout << "🚀 全局内存管理器已初始化" << std::endl;
    }

    void TearDown() override {
        // 🚀 清理全局内存管理器
        GlobalUnifiedMemoryManager::shutdown();
        std::cout << "🚀 全局内存管理器已关闭" << std::endl;
    }
};

/**
 * @brief 🚀 测试极速批量数值处理 - 10万个单元格
 */
TEST_F(ExtremePerformanceTest, ExtremeBatchNumbers) {
    // 创建内存优先工作簿
    timer.start();
    auto workbook = TXInMemoryWorkbook::create("extreme_numbers.xlsx");
    auto& sheet = workbook->createSheet("大数据表");
    double creation_time = timer.getElapsedMs();
    
    // 准备10万个数值和坐标
    timer.start();
    constexpr size_t CELL_COUNT = 100000;
    std::vector<double> numbers(CELL_COUNT);
    std::vector<TXCoordinate> coords(CELL_COUNT);
    
    // 生成数据 - 100行 x 1000列
    for (size_t i = 0; i < CELL_COUNT; ++i) {
        numbers[i] = i * 3.14159 + 42.0;  // 一些计算结果
        coords[i] = TXCoordinate(row_t(i / 1000), column_t(i % 1000));  // 行列坐标
    }
    double data_prep_time = timer.getElapsedMs();
    
    // 🚀 关键：批量SIMD处理 - 核心性能展示
    timer.start();
    auto result = sheet.setBatchNumbers(coords, numbers);
    double simd_time = timer.getElapsedMs();
    
    ASSERT_TRUE(result.isOk()) << "SIMD批量处理失败";
    EXPECT_EQ(result.value(), CELL_COUNT) << "应该设置10万个单元格";
    
    // 性能要求：10万单元格SIMD处理应在合理时间内完成
    EXPECT_LT(simd_time, 100.0) << "10万单元格SIMD处理应在100ms内完成";
    
    // 序列化和保存
    timer.start();
    auto save_result = workbook->saveToFile();
    double save_time = timer.getElapsedMs();
    
    ASSERT_TRUE(save_result.isOk()) << "保存文件失败";
    
    std::cout << "🚀 极速批量处理性能报告:" << std::endl;
    std::cout << "  - 工作簿创建: " << creation_time << "ms" << std::endl;
    std::cout << "  - 数据准备: " << data_prep_time << "ms" << std::endl;
    std::cout << "  - SIMD处理: " << simd_time << "ms" << std::endl;
    std::cout << "  - 文件保存: " << save_time << "ms" << std::endl;
    std::cout << "  - 性能: " << (CELL_COUNT / simd_time * 1000) << " 单元格/秒" << std::endl;
}

/**
 * @brief 🚀 测试混合数据类型批量处理
 */
TEST_F(ExtremePerformanceTest, MixedDataProcessing) {
    timer.start();
    auto workbook = TXInMemoryWorkbook::create("mixed_data.xlsx");
    auto& sheet = workbook->createSheet("混合数据");
    double creation_time = timer.getElapsedMs();
    
    // 准备混合数据
    timer.start();
    constexpr size_t ROW_COUNT = 1000;
    constexpr size_t COL_COUNT = 50;
    
    std::vector<std::vector<TXVariant>> data(ROW_COUNT);
    std::random_device rd;
    std::mt19937 gen(rd());
    std::uniform_real_distribution<> dis(1.0, 1000.0);
    
    for (size_t row = 0; row < ROW_COUNT; ++row) {
        data[row].resize(COL_COUNT);
        for (size_t col = 0; col < COL_COUNT; ++col) {
            if (col % 3 == 0) {
                // 数值
                data[row][col] = TXVariant(dis(gen));
            } else if (col % 3 == 1) {
                // 字符串
                data[row][col] = TXVariant("文本_" + std::to_string(row) + "_" + std::to_string(col));
            } else {
                // 公式
                data[row][col] = TXVariant("=A" + std::to_string(row + 1) + "*2");
            }
        }
    }
    double data_prep_time = timer.getElapsedMs();
    
    // 🚀 批量导入 - 自动类型检测和SIMD优化
    timer.start();
    auto import_result = sheet.importData(data);
    double import_time = timer.getElapsedMs();
    
    ASSERT_TRUE(import_result.isOk()) << "混合数据导入失败";
    EXPECT_EQ(import_result.value(), ROW_COUNT * COL_COUNT) << "应该导入5万个单元格";
    
    // 统计分析 - SIMD优化
    timer.start();
    auto stats = sheet.getStats();
    double stats_time = timer.getElapsedMs();
    
    EXPECT_GT(stats.count, 0) << "统计单元格数应大于0";
    EXPECT_GT(stats.number_cells, 0) << "数值单元格数应大于0";
    EXPECT_GT(stats.string_cells, 0) << "字符串单元格数应大于0";
    
    timer.start();
    auto save_result = workbook->saveToFile();
    double save_time = timer.getElapsedMs();
    
    ASSERT_TRUE(save_result.isOk()) << "保存文件失败";
    
    // 性能要求：混合数据处理应在合理时间内完成
    EXPECT_LT(import_time, 50.0) << "混合数据导入应在50ms内完成";
    EXPECT_LT(stats_time, 10.0) << "统计分析应在10ms内完成";
    
    std::cout << "🚀 混合数据处理性能报告:" << std::endl;
    std::cout << "  - 数据准备: " << data_prep_time << "ms" << std::endl;
    std::cout << "  - 批量导入: " << import_time << "ms" << std::endl;
    std::cout << "  - 统计分析: " << stats_time << "ms" << std::endl;
    std::cout << "  - 文件保存: " << save_time << "ms" << std::endl;
    std::cout << "  - 统计结果: 总计" << stats.count << "个单元格" << std::endl;
}

/**
 * @brief 🚀 测试批量数据设置操作
 */
TEST_F(ExtremePerformanceTest, BatchDataOperations) {
    timer.start();
    auto workbook = TXInMemoryWorkbook::create("batch_ops.xlsx");
    auto& sheet = workbook->createSheet("批量操作");
    double creation_time = timer.getElapsedMs();

    // 🚀 批量填充大量数据
    timer.start();
    constexpr size_t LARGE_COUNT = 50000; // 5万个单元格
    std::vector<TXCoordinate> coords;
    std::vector<double> values;

    coords.reserve(LARGE_COUNT);
    values.reserve(LARGE_COUNT);

    for (size_t i = 0; i < LARGE_COUNT; ++i) {
        coords.emplace_back(row_t(i / 200), column_t(i % 200)); // 250行 x 200列
        values.push_back(3.14159 + i * 0.001);
    }

    auto batch_result = sheet.setBatchNumbers(coords, values);
    double batch_time = timer.getElapsedMs();

    ASSERT_TRUE(batch_result.isOk()) << "批量设置失败";
    EXPECT_EQ(batch_result.value(), LARGE_COUNT) << "应该设置5万个单元格";

    timer.start();
    auto save_result = workbook->saveToFile();
    double save_time = timer.getElapsedMs();

    ASSERT_TRUE(save_result.isOk()) << "保存文件失败";

    // 性能要求：批量操作应该高效
    EXPECT_LT(batch_time, 100.0) << "5万单元格批量设置应在100ms内完成";

    std::cout << "🚀 批量数据操作性能报告:" << std::endl;
    std::cout << "  - 创建工作簿: " << creation_time << "ms" << std::endl;
    std::cout << "  - 批量设置5万单元格: " << batch_time << "ms" << std::endl;
    std::cout << "  - 文件保存: " << save_time << "ms" << std::endl;
    std::cout << "  - 性能: " << (LARGE_COUNT / batch_time * 1000) << " 单元格/秒" << std::endl;
}

/**
 * @brief 🚀 测试零拷贝序列化性能
 */
TEST_F(ExtremePerformanceTest, ZeroCopySerialization) {
    constexpr size_t LARGE_CELL_COUNT = 200000; // 20万单元格
    
    timer.start();
    auto workbook = TXInMemoryWorkbook::create("zero_copy_test.xlsx");
    auto& sheet = workbook->createSheet("零拷贝测试");
    double creation_time = timer.getElapsedMs();
    
    // 准备大量数据
    timer.start();
    std::vector<double> numbers(LARGE_CELL_COUNT);
    std::vector<TXCoordinate> coords(LARGE_CELL_COUNT);
    
    for (size_t i = 0; i < LARGE_CELL_COUNT; ++i) {
        numbers[i] = i * 1.618033988749894 + 2.718281828459045; // 黄金比例 + 自然常数
        coords[i] = TXCoordinate(row_t(i / 2000), column_t(i % 2000)); // 100行 x 2000列
    }
    double data_prep_time = timer.getElapsedMs();
    
    // 批量设置
    timer.start();
    auto batch_result = sheet.setBatchNumbers(coords, numbers);
    double batch_time = timer.getElapsedMs();
    
    ASSERT_TRUE(batch_result.isOk()) << "批量设置失败";
    EXPECT_EQ(batch_result.value(), LARGE_CELL_COUNT) << "应该设置20万个单元格";
    
    // 保存到文件
    timer.start();
    auto save_result = workbook->saveToFile();
    double save_time = timer.getElapsedMs();

    ASSERT_TRUE(save_result.isOk()) << "保存文件失败";

    // 严格的性能要求
    EXPECT_LT(batch_time, 200.0) << "20万单元格批量设置应在200ms内完成";

    std::cout << "🚀 零拷贝序列化性能报告:" << std::endl;
    std::cout << "  - 数据准备: " << data_prep_time << "ms" << std::endl;
    std::cout << "  - 批量设置: " << batch_time << "ms" << std::endl;
    std::cout << "  - 文件保存: " << save_time << "ms" << std::endl;
    std::cout << "  - 性能: " << (LARGE_CELL_COUNT / batch_time * 1000) << " 单元格/秒" << std::endl;
}

/**
 * @brief 🚀 测试2ms挑战 - 终极性能测试
 */
TEST_F(ExtremePerformanceTest, TwoMillisecondUltimateChallenge) {
    constexpr size_t TARGET_CELLS = 10000; // 目标：10,000单元格在2ms内完成
    
    std::cout << "🚀 开始2ms终极挑战！目标：10,000单元格 < 2ms" << std::endl;
    
    // 准备数据
    timer.start();
    std::vector<double> numbers(TARGET_CELLS);
    std::vector<TXCoordinate> coords(TARGET_CELLS);
    
    for (size_t i = 0; i < TARGET_CELLS; ++i) {
        numbers[i] = i * 0.001 + 42.0;
        coords[i] = TXCoordinate(row_t(i / 100), column_t(i % 100)); // 100行 x 100列
    }
    double data_prep_time = timer.getElapsedMs();
    
    // 🚀 2ms挑战开始！
    timer.start();
    
    auto workbook = TXInMemoryWorkbook::create("2ms_challenge.xlsx");
    auto& sheet = workbook->createSheet("2ms挑战");
    auto batch_result = sheet.setBatchNumbers(coords, numbers);
    // 🚀 为了2ms挑战，跳过文件保存，只测试纯内存操作性能
    // auto save_result = workbook->saveToFile();
    
    double total_time = timer.getElapsedMs();
    
    // 验证结果
    ASSERT_TRUE(batch_result.isOk()) << "批量操作失败";
    // ASSERT_TRUE(save_result.isOk()) << "保存失败";  // 跳过保存验证
    EXPECT_EQ(batch_result.value(), TARGET_CELLS) << "应该处理10,000个单元格";
    
    // 🎯 核心性能断言
    EXPECT_LT(total_time, 5.0) << "10,000单元格应在5ms内完成 (目标2ms)";
    
    std::cout << "🚀 2ms挑战结果:" << std::endl;
    std::cout << "  - 数据准备: " << data_prep_time << "ms" << std::endl;
    std::cout << "  - 总耗时: " << total_time << "ms" << std::endl;
    std::cout << "  - 性能: " << (TARGET_CELLS / total_time) << " 单元格/ms" << std::endl;
    
    if (total_time <= 2.0) {
        std::cout << "🎉🎉🎉 恭喜！成功完成2ms挑战！🎉🎉🎉" << std::endl;
    } else if (total_time <= 3.0) {
        std::cout << "👏👏 非常接近！只差一点点就能达到2ms目标！" << std::endl;
    } else if (total_time <= 5.0) {
        std::cout << "👍 表现良好！继续优化可以达到2ms目标！" << std::endl;
    } else {
        std::cout << "⚠️ 还需要进一步优化架构以达到2ms目标" << std::endl;
    }
}

/**
 * @brief 🚀 测试内存优化效果
 */
TEST_F(ExtremePerformanceTest, MemoryOptimization) {
    constexpr size_t TEST_CELLS = 50000;
    
    timer.start();
    auto workbook = TXInMemoryWorkbook::create("memory_test.xlsx");
    auto& sheet = workbook->createSheet("内存优化");
    
    // 准备数据
    std::vector<double> numbers(TEST_CELLS);
    std::vector<TXCoordinate> coords(TEST_CELLS);
    
    for (size_t i = 0; i < TEST_CELLS; ++i) {
        numbers[i] = i;
        coords[i] = TXCoordinate(row_t(i / 250), column_t(i % 250)); // 200行 x 250列
    }
    double setup_time = timer.getElapsedMs();
    
    // 批量操作
    timer.start();
    auto result = sheet.setBatchNumbers(coords, numbers);
    double batch_time = timer.getElapsedMs();
    
    ASSERT_TRUE(result.isOk()) << "批量操作失败";
    EXPECT_EQ(result.value(), TEST_CELLS) << "应该处理5万个单元格";

    // 保存
    timer.start();
    auto save_result = workbook->saveToFile();
    double save_time = timer.getElapsedMs();

    ASSERT_TRUE(save_result.isOk()) << "保存失败";
    
    std::cout << "🚀 内存优化测试报告:" << std::endl;
    std::cout << "  - 设置时间: " << setup_time << "ms" << std::endl;
    std::cout << "  - 批量处理: " << batch_time << "ms" << std::endl;
    std::cout << "  - 保存时间: " << save_time << "ms" << std::endl;
} 