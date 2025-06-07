//
// @file test_tx_high_performance_xlsx_reader.cpp
// @brief 🚀 高性能XLSX读取器测试
//

#include <gtest/gtest.h>
#include "TinaXlsx/io/TXHighPerformanceXLSXReader.hpp"
#include "TinaXlsx/TXGlobalMemoryManager.hpp"
#include "TinaXlsx/TXHighPerformanceLogger.hpp"
#include <filesystem>
#include <fstream>

using namespace TinaXlsx;

class TXHighPerformanceXLSXReaderTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 初始化日志系统
        TXGlobalLogger::initialize(GlobalUnifiedMemoryManager::getInstance());
        TXGlobalLogger::setOutputMode(TXLogOutputMode::CONSOLE_ONLY);
        
        // 创建测试目录
        test_dir_ = "test_hp_xlsx_reader";
        std::filesystem::create_directories(test_dir_);
        
        // 创建高性能读取器
        reader_ = std::make_unique<TXHighPerformanceXLSXReader>(
            GlobalUnifiedMemoryManager::getInstance()
        );
    }
    
    void TearDown() override {
        reader_.reset();
        
        // 清理测试文件
        if (std::filesystem::exists(test_dir_)) {
            std::filesystem::remove_all(test_dir_);
        }
    }
    
    std::string getTestFilePath(const std::string& filename) {
        return test_dir_ + "/" + filename;
    }
    
    void createTestXLSXFile(const std::string& path) {
        // 创建一个简单的ZIP文件模拟XLSX
        std::ofstream file(path, std::ios::binary);
        
        // 写入ZIP文件头
        file.write("PK", 2);  // ZIP文件标识
        file.write("\x03\x04", 2);  // 本地文件头标识
        
        // 写入一些模拟数据
        std::string mock_data = "Mock XLSX content for testing";
        file.write(mock_data.c_str(), mock_data.size());
        
        file.close();
    }

protected:
    std::string test_dir_;
    std::unique_ptr<TXHighPerformanceXLSXReader> reader_;
};

// ==================== 基础功能测试 ====================

TEST_F(TXHighPerformanceXLSXReaderTest, ConstructorAndConfig) {
    TX_LOG_INFO("🚀 测试高性能XLSX读取器构造和配置");
    
    // 测试默认配置
    auto default_config = reader_->getConfig();
    EXPECT_TRUE(default_config.enable_simd_processing);
    EXPECT_TRUE(default_config.enable_memory_optimization);
    EXPECT_TRUE(default_config.enable_parallel_parsing);
    EXPECT_EQ(default_config.buffer_initial_capacity, 10000);
    
    // 测试自定义配置
    TXHighPerformanceXLSXReader::Config custom_config;
    custom_config.enable_simd_processing = false;
    custom_config.buffer_initial_capacity = 5000;
    
    reader_->updateConfig(custom_config);
    auto updated_config = reader_->getConfig();
    EXPECT_FALSE(updated_config.enable_simd_processing);
    EXPECT_EQ(updated_config.buffer_initial_capacity, 5000);
    
    TX_LOG_INFO("配置测试通过");
}

TEST_F(TXHighPerformanceXLSXReaderTest, FileValidation) {
    TX_LOG_INFO("🚀 测试文件验证功能");
    
    std::string valid_xlsx = getTestFilePath("valid.xlsx");
    std::string invalid_file = getTestFilePath("invalid.txt");
    std::string nonexistent_file = getTestFilePath("nonexistent.xlsx");
    
    // 创建有效的XLSX文件
    createTestXLSXFile(valid_xlsx);
    
    // 创建无效文件
    std::ofstream invalid(invalid_file);
    invalid << "This is not a ZIP file";
    invalid.close();
    
    // 测试文件验证
    EXPECT_TRUE(TXHighPerformanceXLSXReader::isValidXLSXFile(valid_xlsx));
    EXPECT_FALSE(TXHighPerformanceXLSXReader::isValidXLSXFile(invalid_file));
    EXPECT_FALSE(TXHighPerformanceXLSXReader::isValidXLSXFile(nonexistent_file));
    
    TX_LOG_INFO("文件验证测试通过");
}

TEST_F(TXHighPerformanceXLSXReaderTest, MemoryEstimation) {
    TX_LOG_INFO("🚀 测试内存需求估算");
    
    std::string test_file = getTestFilePath("test.xlsx");
    createTestXLSXFile(test_file);
    
    auto memory_estimate = TXHighPerformanceXLSXReader::estimateMemoryRequirement(test_file);
    EXPECT_TRUE(memory_estimate.isOk());
    
    size_t estimated_memory = memory_estimate.value();
    EXPECT_GT(estimated_memory, 0);
    
    TX_LOG_INFO("预估内存需求: {} 字节", estimated_memory);
    TX_LOG_INFO("内存估算测试通过");
}

// ==================== 核心功能测试 ====================

TEST_F(TXHighPerformanceXLSXReaderTest, LoadXLSXFile) {
    TX_LOG_INFO("🚀 测试XLSX文件读取");
    
    std::string test_file = getTestFilePath("workbook.xlsx");
    createTestXLSXFile(test_file);
    
    // 测试文件读取
    auto result = reader_->loadXLSX(test_file);
    EXPECT_TRUE(result.isOk()) << "XLSX读取失败: " << result.error().getMessage();
    
    if (result.isOk()) {
        auto workbook = result.value();
        EXPECT_NE(workbook, nullptr);
        EXPECT_GT(workbook->getSheetCount(), 0);
        
        // 检查统计信息
        auto stats = reader_->getLastReadStats();
        EXPECT_GT(stats.total_time_ms, 0);
        EXPECT_GT(stats.total_sheets_read, 0);
        
        TX_LOG_INFO("读取统计: {:.3f}ms, {} 个工作表, {} 个单元格", 
                   stats.total_time_ms, stats.total_sheets_read, stats.total_cells_read);
    }
    
    TX_LOG_INFO("XLSX文件读取测试通过");
}

TEST_F(TXHighPerformanceXLSXReaderTest, LoadXLSXFromMemory) {
    TX_LOG_INFO("🚀 测试从内存读取XLSX");
    
    // 创建内存数据
    std::string mock_xlsx_data = "PK\x03\x04Mock XLSX data in memory";
    
    auto result = reader_->loadXLSXFromMemory(mock_xlsx_data.data(), mock_xlsx_data.size());
    EXPECT_TRUE(result.isOk()) << "内存XLSX读取失败: " << result.error().getMessage();
    
    if (result.isOk()) {
        auto workbook = result.value();
        EXPECT_NE(workbook, nullptr);
        EXPECT_EQ(workbook->getName(), "XLSX_Memory_Loaded");
    }
    
    TX_LOG_INFO("内存XLSX读取测试通过");
}

TEST_F(TXHighPerformanceXLSXReaderTest, LoadSheetToBuffer) {
    TX_LOG_INFO("🚀 测试工作表读取到高性能缓冲区");
    
    std::string test_file = getTestFilePath("sheet_test.xlsx");
    createTestXLSXFile(test_file);
    
    auto result = reader_->loadSheetToBuffer(test_file, "Sheet1");
    EXPECT_TRUE(result.isOk()) << "工作表缓冲区读取失败: " << result.error().getMessage();
    
    if (result.isOk()) {
        auto buffer = result.value();
        EXPECT_GE(buffer.capacity, 0);
        TX_LOG_INFO("缓冲区容量: {}, 大小: {}", buffer.capacity, buffer.size);
    }
    
    TX_LOG_INFO("工作表缓冲区读取测试通过");
}

// ==================== 高性能处理测试 ====================

TEST_F(TXHighPerformanceXLSXReaderTest, SIMDProcessing) {
    TX_LOG_INFO("🚀 测试SIMD批量处理");
    
    // 创建测试缓冲区
    TXCompactCellBuffer buffer(GlobalUnifiedMemoryManager::getInstance(), 1000);
    
    // 添加一些测试数据
    buffer.reserve(100);
    for (size_t i = 0; i < 10; ++i) {
        size_t index = buffer.size++;
        buffer.coordinates[index] = ((i + 1) << 16) | (i + 1);
        buffer.number_values[index] = static_cast<double>(i * 10);
        buffer.cell_types[index] = static_cast<uint8_t>(TXCellType::Number);
        buffer.style_indices[index] = 0;
        buffer.string_indices[index] = 0;
    }
    
    // 测试SIMD处理
    auto result = reader_->processWithSIMD(buffer);
    EXPECT_TRUE(result.isOk()) << "SIMD处理失败: " << result.error().getMessage();
    
    TX_LOG_INFO("SIMD处理测试通过");
}

TEST_F(TXHighPerformanceXLSXReaderTest, MemoryOptimization) {
    TX_LOG_INFO("🚀 测试内存布局优化");
    
    // 创建测试缓冲区
    TXCompactCellBuffer buffer(GlobalUnifiedMemoryManager::getInstance(), 1000);
    
    // 测试内存优化
    auto result = reader_->optimizeMemoryLayout(buffer);
    EXPECT_TRUE(result.isOk()) << "内存优化失败: " << result.error().getMessage();
    
    TX_LOG_INFO("内存优化测试通过");
}

TEST_F(TXHighPerformanceXLSXReaderTest, StatisticsCalculation) {
    TX_LOG_INFO("🚀 测试统计信息计算");
    
    // 创建测试缓冲区
    TXCompactCellBuffer buffer(GlobalUnifiedMemoryManager::getInstance(), 1000);
    
    // 添加混合类型的测试数据
    buffer.reserve(20);
    for (size_t i = 0; i < 15; ++i) {
        size_t index = buffer.size++;
        buffer.coordinates[index] = ((i + 1) << 16) | (i + 1);
        
        if (i < 10) {
            buffer.number_values[index] = static_cast<double>(i);
            buffer.cell_types[index] = static_cast<uint8_t>(TXCellType::Number);
        } else {
            buffer.string_indices[index] = i - 10;
            buffer.cell_types[index] = static_cast<uint8_t>(TXCellType::String);
        }
        
        buffer.style_indices[index] = 0;
    }
    
    // 计算统计信息
    auto result = reader_->calculateStatistics(buffer);
    EXPECT_TRUE(result.isOk()) << "统计计算失败: " << result.error().getMessage();
    
    if (result.isOk()) {
        auto stats = result.value();
        EXPECT_EQ(stats.total_cells, 15);
        EXPECT_EQ(stats.number_cells, 10);
        EXPECT_EQ(stats.string_cells, 5);
        
        TX_LOG_INFO("统计结果: 总计={}, 数字={}, 字符串={}, 空={}", 
                   stats.total_cells, stats.number_cells, 
                   stats.string_cells, stats.empty_cells);
    }
    
    TX_LOG_INFO("统计计算测试通过");
}

// ==================== 错误处理测试 ====================

TEST_F(TXHighPerformanceXLSXReaderTest, ErrorHandling) {
    TX_LOG_INFO("🚀 测试错误处理");
    
    // 测试不存在的文件
    auto result1 = reader_->loadXLSX("nonexistent_file.xlsx");
    EXPECT_FALSE(result1.isOk());
    EXPECT_EQ(result1.error().getCode(), TXErrorCode::FileNotFound);
    
    // 测试无效的内存数据
    auto result2 = reader_->loadXLSXFromMemory(nullptr, 0);
    EXPECT_FALSE(result2.isOk());
    EXPECT_EQ(result2.error().getCode(), TXErrorCode::InvalidArgument);
    
    TX_LOG_INFO("错误处理测试通过");
}

// ==================== 性能基准测试 ====================

TEST_F(TXHighPerformanceXLSXReaderTest, PerformanceBenchmark) {
    TX_LOG_INFO("🚀 性能基准测试");
    
    std::string test_file = getTestFilePath("benchmark.xlsx");
    createTestXLSXFile(test_file);
    
    // 重置统计信息
    reader_->resetStats();
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 执行读取
    auto result = reader_->loadXLSX(test_file);
    EXPECT_TRUE(result.isOk());
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto total_time = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time).count() / 1000.0;
    
    // 获取详细统计
    auto stats = reader_->getLastReadStats();
    
    TX_LOG_INFO("🚀 性能基准结果:");
    TX_LOG_INFO("  总耗时: {:.3f}ms", total_time);
    TX_LOG_INFO("  解析耗时: {:.3f}ms", stats.parsing_time_ms);
    TX_LOG_INFO("  导入耗时: {:.3f}ms", stats.import_time_ms);
    TX_LOG_INFO("  SIMD处理耗时: {:.3f}ms", stats.simd_processing_time_ms);
    TX_LOG_INFO("  内存使用: {:.2f} MB", stats.memory_used_bytes / (1024.0 * 1024.0));
    TX_LOG_INFO("  处理单元格: {}", stats.total_cells_read);
    
    // 性能断言（根据实际情况调整）
    EXPECT_LT(total_time, 1000.0);  // 应该在1秒内完成
    
    TX_LOG_INFO("性能基准测试通过");
}
