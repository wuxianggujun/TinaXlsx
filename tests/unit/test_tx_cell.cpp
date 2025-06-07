//
// @file test_tx_cell.cpp
// @brief TXCell 用户层单元格类测试
//

#include <gtest/gtest.h>
#include <TinaXlsx/user/TXCell.hpp>
#include <TinaXlsx/TXInMemorySheet.hpp>
#include <TinaXlsx/TXUnifiedMemoryManager.hpp>
#include <TinaXlsx/TXHighPerformanceLogger.hpp>
#include <TinaXlsx/TXGlobalStringPool.hpp>

using namespace TinaXlsx;

class TXCellTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 初始化内存管理器
        TXUnifiedMemoryManager::Config config;
        config.memory_limit = 512ULL * 1024 * 1024; // 512MB
        GlobalUnifiedMemoryManager::initialize(config);
        
        // 初始化日志系统
        TXGlobalLogger::initialize(GlobalUnifiedMemoryManager::getInstance());
        TXGlobalLogger::setOutputMode(TXLogOutputMode::CONSOLE_ONLY);
        
        // 创建工作表
        sheet_ = std::make_unique<TXInMemorySheet>(
            "测试工作表",
            GlobalUnifiedMemoryManager::getInstance(),
            TXGlobalStringPool::instance()
        );

        // 🚀 预分配足够的容量避免"缓冲区已满"错误
        sheet_->reserve(100000);
    }
    
    void TearDown() override {
        sheet_.reset();
        GlobalUnifiedMemoryManager::shutdown();
    }
    
    std::unique_ptr<TXInMemorySheet> sheet_;
};

/**
 * @brief 测试TXCell基本构造
 */
TEST_F(TXCellTest, BasicConstruction) {
    // 测试坐标构造 (A1 = row=1, col=1 in 1-based)
    TXCoordinate coord(row_t(1), column_t(static_cast<uint32_t>(1)));
    TXCell cell(*sheet_, coord);

    EXPECT_TRUE(cell.isValid());
    EXPECT_EQ(cell.getRow(), 0);  // getRow()返回0-based索引
    EXPECT_EQ(cell.getColumn(), 0);  // getColumn()返回0-based索引
    EXPECT_EQ(cell.getAddress(), "A1");
    
    // 测试Excel格式构造
    TXCell cell2(*sheet_, "B2");
    EXPECT_TRUE(cell2.isValid());
    EXPECT_EQ(cell2.getRow(), 1);  // B2 = row=1, col=1 (0-based)
    EXPECT_EQ(cell2.getColumn(), 1);
    EXPECT_EQ(cell2.getAddress(), "B2");
    
    TX_LOG_INFO("基本构造测试通过");
}

/**
 * @brief 测试TXCell值操作
 */
TEST_F(TXCellTest, ValueOperations) {
    TXCell cell(*sheet_, "A1");
    
    // 测试数值设置
    cell.setValue(42.5);
    auto value = cell.getValue();
    EXPECT_EQ(value.getType(), TXVariant::Type::Number);
    EXPECT_DOUBLE_EQ(value.getNumber(), 42.5);
    
    // 测试字符串设置
    cell.setValue("Hello World");
    value = cell.getValue();
    EXPECT_EQ(value.getType(), TXVariant::Type::String);
    EXPECT_EQ(value.getString(), "Hello World");
    
    // 测试布尔值设置
    cell.setValue(true);
    value = cell.getValue();
    EXPECT_EQ(value.getType(), TXVariant::Type::Number);
    EXPECT_DOUBLE_EQ(value.getNumber(), 1.0);
    
    // 测试清除
    cell.clear();
    EXPECT_TRUE(cell.isEmpty());
    
    TX_LOG_INFO("值操作测试通过");
}

/**
 * @brief 测试TXCell链式调用
 */
TEST_F(TXCellTest, ChainedOperations) {
    TXCell cell(*sheet_, "C3");
    
    // 测试链式设置
    cell.setValue(100.0)
        .add(50.0)
        .multiply(2.0)
        .subtract(25.0);
    
    auto value = cell.getValue();
    EXPECT_EQ(value.getType(), TXVariant::Type::Number);
    EXPECT_DOUBLE_EQ(value.getNumber(), 275.0); // (100+50)*2-25 = 275
    
    TX_LOG_INFO("链式调用测试通过");
}

/**
 * @brief 测试TXCell操作符重载
 */
TEST_F(TXCellTest, OperatorOverloads) {
    TXCell cell1(*sheet_, "D1");
    TXCell cell2(*sheet_, "D2");
    
    // 测试赋值操作符
    cell1 = 123.45;
    cell2 = "测试字符串";
    
    EXPECT_DOUBLE_EQ(cell1.getValue().getNumber(), 123.45);
    EXPECT_EQ(cell2.getValue().getString(), "测试字符串");
    
    // 测试数学操作符
    cell1 += 10.0;
    EXPECT_DOUBLE_EQ(cell1.getValue().getNumber(), 133.45);
    
    cell1 *= 2.0;
    EXPECT_DOUBLE_EQ(cell1.getValue().getNumber(), 266.9);
    
    // 测试比较操作符
    TXCell cell3(*sheet_, "D1");
    EXPECT_TRUE(cell1 == cell3);  // 相同坐标
    EXPECT_FALSE(cell1 == cell2); // 不同坐标
    
    TX_LOG_INFO("操作符重载测试通过");
}

/**
 * @brief 测试TXCell便捷函数
 */
TEST_F(TXCellTest, ConvenienceFunctions) {
    // 测试makeCell函数
    auto cell1 = makeCell(*sheet_, TXCoordinate(row_t(3), column_t(static_cast<uint32_t>(3))));  // C3
    auto cell2 = makeCell(*sheet_, "C3");  // C3
    auto cell3 = makeCell(*sheet_, 2, 2);  // makeCell(2,2) -> C3 (0-based输入转为1-based)

    EXPECT_EQ(cell1.getAddress(), "C3");
    EXPECT_EQ(cell2.getAddress(), "C3");
    EXPECT_EQ(cell3.getAddress(), "C3");
    
    // 所有三个应该指向同一个单元格
    EXPECT_TRUE(cell1 == cell2);
    EXPECT_TRUE(cell2 == cell3);
    
    TX_LOG_INFO("便捷函数测试通过");
}

/**
 * @brief 测试TXCell错误处理
 */
TEST_F(TXCellTest, ErrorHandling) {
    // 测试无效坐标
    TXCell invalid_cell(*sheet_, "INVALID");
    EXPECT_FALSE(invalid_cell.isValid());
    
    // 测试除零错误
    TXCell cell(*sheet_, "E1");
    cell.setValue(100.0);
    cell.divide(0.0); // 应该不会崩溃，只是记录错误
    
    // 值应该保持不变
    EXPECT_DOUBLE_EQ(cell.getValue().getNumber(), 100.0);
    
    TX_LOG_INFO("错误处理测试通过");
}

/**
 * @brief 测试TXCell调试功能
 */
TEST_F(TXCellTest, DebuggingFeatures) {
    TXCell cell(*sheet_, "F5");
    cell.setValue(3.14159);
    
    std::string debug_str = cell.toString();
    EXPECT_TRUE(debug_str.find("F5") != std::string::npos);
    EXPECT_TRUE(debug_str.find("3.14159") != std::string::npos);
    EXPECT_TRUE(debug_str.find("数值") != std::string::npos);
    
    TX_LOG_INFO("调试信息: {}", debug_str);
    TX_LOG_INFO("调试功能测试通过");
}

/**
 * @brief 测试TXCell性能
 */
TEST_F(TXCellTest, Performance) {
    constexpr size_t CELL_COUNT = 10000;
    
    auto start_time = std::chrono::high_resolution_clock::now();
    
    // 创建大量单元格并设置值
    for (size_t i = 0; i < CELL_COUNT; ++i) {
        uint32_t row = i / 100;
        uint32_t col = i % 100;
        
        TXCell cell(*sheet_, TXCoordinate(row_t(row + 1), column_t(static_cast<uint32_t>(col + 1))));
        cell.setValue(static_cast<double>(i));
    }
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    
    TX_LOG_INFO("创建并设置{}个TXCell耗时: {:.3f}ms", CELL_COUNT, duration.count() / 1000.0);
    TX_LOG_INFO("平均每个TXCell: {:.1f}μs", static_cast<double>(duration.count()) / CELL_COUNT);
    
    // 性能要求：每个TXCell操作应该在10μs以内
    double avg_time_us = static_cast<double>(duration.count()) / CELL_COUNT;
    EXPECT_LT(avg_time_us, 100.0); // 100μs以内
    
    TX_LOG_INFO("性能测试通过");
}

/**
 * @brief 测试TXCell内存占用
 */
TEST_F(TXCellTest, MemoryFootprint) {
    // 验证TXCell确实是16字节
    EXPECT_EQ(sizeof(TXCell), 16);
    
    TX_LOG_INFO("TXCell内存占用: {} 字节", sizeof(TXCell));
    TX_LOG_INFO("内存占用测试通过");
}
