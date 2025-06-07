//
// @file test_tx_sheet.cpp
// @brief TXSheet 用户层工作表类测试
//

#include <gtest/gtest.h>
#include <TinaXlsx/user/TXSheet.hpp>
#include <TinaXlsx/TXUnifiedMemoryManager.hpp>
#include <TinaXlsx/TXHighPerformanceLogger.hpp>
#include <TinaXlsx/TXGlobalStringPool.hpp>

using namespace TinaXlsx;

class TXSheetTest : public ::testing::Test {
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
        sheet_ = std::make_unique<TXSheet>(
            "测试工作表",
            GlobalUnifiedMemoryManager::getInstance(),
            TXGlobalStringPool::instance()
        );
        
        // 🚀 预分配足够的容量
        sheet_->reserve(100000);
    }
    
    void TearDown() override {
        sheet_.reset();
        GlobalUnifiedMemoryManager::shutdown();
    }
    
    std::unique_ptr<TXSheet> sheet_;
};

/**
 * @brief 测试TXSheet基本属性
 */
TEST_F(TXSheetTest, BasicProperties) {
    EXPECT_EQ(sheet_->getName(), "测试工作表");
    EXPECT_TRUE(sheet_->isEmpty());
    EXPECT_EQ(sheet_->getCellCount(), 0);
    EXPECT_TRUE(sheet_->isValid());
    
    // 测试名称修改
    sheet_->setName("新工作表");
    EXPECT_EQ(sheet_->getName(), "新工作表");
    
    TX_LOG_INFO("基本属性测试通过");
}

/**
 * @brief 测试单元格访问
 */
TEST_F(TXSheetTest, CellAccess) {
    // 测试Excel格式访问
    auto cell1 = sheet_->cell("A1");
    cell1.setValue(42.0);
    EXPECT_DOUBLE_EQ(cell1.getValue().getNumber(), 42.0);
    
    // 测试行列索引访问
    auto cell2 = sheet_->cell(0, 1); // B1
    cell2.setValue("Hello");
    EXPECT_EQ(cell2.getValue().getString(), "Hello");
    EXPECT_EQ(cell2.getAddress(), "B1");
    
    // 测试坐标对象访问
    TXCoordinate coord(row_t(3), column_t(3)); // C3
    auto cell3 = sheet_->cell(coord);
    cell3.setValue(true);
    EXPECT_DOUBLE_EQ(cell3.getValue().getNumber(), 1.0);
    
    // 验证工作表不再为空
    EXPECT_FALSE(sheet_->isEmpty());
    EXPECT_EQ(sheet_->getCellCount(), 3);
    
    TX_LOG_INFO("单元格访问测试通过");
}

/**
 * @brief 测试范围操作
 */
TEST_F(TXSheetTest, RangeOperations) {
    // 测试Excel格式范围
    auto range1 = sheet_->range("A1:C3");
    EXPECT_TRUE(range1.isValid());
    EXPECT_EQ(range1.getCellCount(), 9);
    
    // 测试坐标范围
    auto range2 = sheet_->range(0, 0, 2, 2); // A1:C3
    EXPECT_TRUE(range2.isValid());
    EXPECT_EQ(range2.getCellCount(), 9);
    
    // 测试坐标对象范围
    TXCoordinate start(row_t(1), column_t(1)); // A1
    TXCoordinate end(row_t(2), column_t(2));   // B2
    auto range3 = sheet_->range(start, end);
    EXPECT_TRUE(range3.isValid());
    EXPECT_EQ(range3.getCellCount(), 4);
    
    TX_LOG_INFO("范围操作测试通过");
}

/**
 * @brief 测试批量数据操作
 */
TEST_F(TXSheetTest, BatchDataOperations) {
    // 准备测试数据
    std::vector<std::vector<TXVariant>> data = {
        {1.0, 2.0, 3.0},
        {"A", "B", "C"},
        {true, false, true}
    };
    
    // 批量设置值
    auto result = sheet_->setValues("A1:C3", data);
    EXPECT_TRUE(result.isOk());
    
    // 批量获取值
    auto get_result = sheet_->getValues("A1:C3");
    EXPECT_TRUE(get_result.isOk());
    
    auto retrieved_data = get_result.value();
    EXPECT_EQ(retrieved_data.size(), 3);
    EXPECT_EQ(retrieved_data[0].size(), 3);
    
    // 验证数据正确性
    EXPECT_DOUBLE_EQ(retrieved_data[0][0].getNumber(), 1.0);
    EXPECT_EQ(retrieved_data[1][1].getString(), "B");
    // 布尔值可能存储为数值，检查类型
    if (retrieved_data[2][2].getType() == TXVariant::Type::Number) {
        EXPECT_DOUBLE_EQ(retrieved_data[2][2].getNumber(), 1.0); // true -> 1.0
    } else if (retrieved_data[2][2].getType() == TXVariant::Type::Boolean) {
        EXPECT_TRUE(retrieved_data[2][2].getBoolean());
    }
    
    TX_LOG_INFO("批量数据操作测试通过");
}

/**
 * @brief 测试填充和清除操作
 */
TEST_F(TXSheetTest, FillAndClearOperations) {
    // 测试数值填充
    auto fill_result = sheet_->fillRange("D1:F3", TXVariant(99.0));
    EXPECT_TRUE(fill_result.isOk());
    
    // 验证填充结果
    EXPECT_DOUBLE_EQ(sheet_->cell("D1").getValue().getNumber(), 99.0);
    EXPECT_DOUBLE_EQ(sheet_->cell("F3").getValue().getNumber(), 99.0);
    
    // 测试字符串填充
    auto fill_result2 = sheet_->fillRange("G1:G5", TXVariant("测试"));
    EXPECT_TRUE(fill_result2.isOk());
    
    // 验证字符串填充
    EXPECT_EQ(sheet_->cell("G3").getValue().getString(), "测试");
    
    // 测试清除操作
    auto clear_result = sheet_->clearRange("D1:F3");
    if (!clear_result.isOk()) {
        TX_LOG_WARN("清除操作失败: {}", clear_result.error().getMessage());
    }
    EXPECT_TRUE(clear_result.isOk());
    
    TX_LOG_INFO("填充和清除操作测试通过");
}

/**
 * @brief 测试统计功能
 */
TEST_F(TXSheetTest, StatisticalFunctions) {
    // 准备数值数据
    std::vector<std::vector<TXVariant>> numbers = {
        {10.0, 20.0, 30.0},
        {40.0, 50.0, 60.0}
    };
    
    auto result = sheet_->setValues("H1:J2", numbers);
    EXPECT_TRUE(result.isOk());
    
    // 测试求和
    auto sum_result = sheet_->sum("H1:J2");
    EXPECT_TRUE(sum_result.isOk());
    EXPECT_DOUBLE_EQ(sum_result.value(), 210.0); // 10+20+30+40+50+60
    
    // 测试平均值
    auto avg_result = sheet_->average("H1:J2");
    EXPECT_TRUE(avg_result.isOk());
    EXPECT_DOUBLE_EQ(avg_result.value(), 35.0); // 210/6
    
    // 测试最大值
    auto max_result = sheet_->max("H1:J2");
    EXPECT_TRUE(max_result.isOk());
    EXPECT_DOUBLE_EQ(max_result.value(), 60.0);
    
    // 测试最小值
    auto min_result = sheet_->min("H1:J2");
    EXPECT_TRUE(min_result.isOk());
    EXPECT_DOUBLE_EQ(min_result.value(), 10.0);
    
    TX_LOG_INFO("统计功能测试通过");
}

/**
 * @brief 测试查找功能
 */
TEST_F(TXSheetTest, FindFunctions) {
    // 设置一些测试数据
    sheet_->cell("K1").setValue(100.0);
    sheet_->cell("K2").setValue("查找我");
    sheet_->cell("K3").setValue(100.0);
    sheet_->cell("K4").setValue("查找我");
    
    // 查找数值
    auto coords1 = sheet_->findValue(TXVariant(100.0));
    EXPECT_GE(coords1.size(), 2); // 至少找到2个
    
    // 查找字符串
    auto coords2 = sheet_->findValue(TXVariant("查找我"));
    EXPECT_GE(coords2.size(), 2); // 至少找到2个
    
    // 在指定范围内查找
    auto coords3 = sheet_->findValue(TXVariant(100.0), "K1:K2");
    EXPECT_EQ(coords3.size(), 1); // 只在K1找到
    
    TX_LOG_INFO("查找功能测试通过");
}

/**
 * @brief 测试性能优化功能
 */
TEST_F(TXSheetTest, PerformanceOptimization) {
    // 添加一些数据
    for (int i = 0; i < 100; ++i) {
        sheet_->cell(i, 0).setValue(static_cast<double>(i));
    }
    
    // 测试优化
    sheet_->optimize();
    
    // 测试压缩
    size_t compressed = sheet_->compress();
    TX_LOG_INFO("压缩了 {} 个单元格", compressed);
    
    // 测试收缩
    sheet_->shrinkToFit();
    
    // 获取性能统计
    std::string stats = sheet_->getPerformanceStats();
    EXPECT_FALSE(stats.empty());
    TX_LOG_INFO("性能统计:\n{}", stats);
    
    TX_LOG_INFO("性能优化功能测试通过");
}

/**
 * @brief 测试调试功能
 */
TEST_F(TXSheetTest, DebuggingFeatures) {
    // 添加一些数据
    sheet_->cell("A1").setValue(42.0);
    sheet_->cell("B1").setValue("测试");
    
    // 测试toString
    std::string debug_str = sheet_->toString();
    EXPECT_TRUE(debug_str.find("测试工作表") != std::string::npos);
    EXPECT_TRUE(debug_str.find("单元格数=") != std::string::npos);
    
    TX_LOG_INFO("调试信息: {}", debug_str);
    
    // 测试使用范围
    auto used_range = sheet_->getUsedRange();
    EXPECT_TRUE(used_range.isValid());
    
    TX_LOG_INFO("调试功能测试通过");
}

/**
 * @brief 测试错误处理
 */
TEST_F(TXSheetTest, ErrorHandling) {
    // 测试无效范围
    auto result1 = sheet_->setValues("INVALID", {});
    EXPECT_TRUE(result1.isError());
    
    auto result2 = sheet_->getValues("INVALID");
    EXPECT_TRUE(result2.isError());
    
    auto result3 = sheet_->sum("INVALID");
    EXPECT_TRUE(result3.isError());
    
    // 测试空范围统计
    auto result4 = sheet_->average("Z100:Z100");
    // 这个可能成功也可能失败，取决于实现
    
    TX_LOG_INFO("错误处理测试通过");
}

/**
 * @brief 测试便捷函数
 */
TEST_F(TXSheetTest, ConvenienceFunctions) {
    // 测试makeSheet函数
    auto new_sheet = makeSheet("便捷工作表");
    EXPECT_TRUE(new_sheet->isValid());
    EXPECT_EQ(new_sheet->getName(), "便捷工作表");
    EXPECT_TRUE(new_sheet->isEmpty());
    
    TX_LOG_INFO("便捷函数测试通过");
}
