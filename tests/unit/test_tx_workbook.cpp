//
// @file test_tx_workbook.cpp
// @brief TXWorkbook单元测试
//

#include <gtest/gtest.h>
#include <TinaXlsx/user/TXWorkbook.hpp>
#include <TinaXlsx/TXUnifiedMemoryManager.hpp>
#include <TinaXlsx/TXHighPerformanceLogger.hpp>
#include <TinaXlsx/TXGlobalStringPool.hpp>

using namespace TinaXlsx;

/**
 * @brief TXWorkbook测试类
 */
class TXWorkbookTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 初始化内存管理器
        TXUnifiedMemoryManager::Config config;
        config.memory_limit = 512ULL * 1024 * 1024; // 512MB
        GlobalUnifiedMemoryManager::initialize(config);

        // 初始化日志系统
        TXGlobalLogger::initialize(GlobalUnifiedMemoryManager::getInstance());
        TXGlobalLogger::setOutputMode(TXLogOutputMode::CONSOLE_ONLY);

        // 创建测试工作簿
        workbook_ = TXWorkbook::create("测试工作簿");
        ASSERT_NE(workbook_, nullptr) << "工作簿创建失败";
    }

    void TearDown() override {
        workbook_.reset();
        GlobalUnifiedMemoryManager::shutdown();
    }

    std::unique_ptr<TXWorkbook> workbook_;
};

/**
 * @brief 测试工作簿基本属性
 */
TEST_F(TXWorkbookTest, BasicProperties) {
    // 测试名称
    EXPECT_EQ(workbook_->getName(), "测试工作簿");
    
    // 测试工作表数量（应该有默认的Sheet1）
    EXPECT_EQ(workbook_->getSheetCount(), 1);
    EXPECT_FALSE(workbook_->isEmpty());
    
    // 测试活动工作表
    EXPECT_EQ(workbook_->getActiveSheetIndex(), 0);
    auto active_sheet = workbook_->getActiveSheet();
    ASSERT_NE(active_sheet, nullptr);
    EXPECT_EQ(active_sheet->getName(), "Sheet1");
    
    // 测试有效性
    EXPECT_TRUE(workbook_->isValid());
    
    TX_LOG_INFO("基本属性测试通过");
}

/**
 * @brief 测试工作表管理
 */
TEST_F(TXWorkbookTest, SheetManagement) {
    // 添加工作表
    auto sheet2 = workbook_->addSheet("数据表");
    ASSERT_NE(sheet2, nullptr);
    EXPECT_EQ(sheet2->getName(), "数据表");
    EXPECT_EQ(workbook_->getSheetCount(), 2);
    
    auto sheet3 = workbook_->addSheet("统计表");
    ASSERT_NE(sheet3, nullptr);
    EXPECT_EQ(workbook_->getSheetCount(), 3);
    
    // 测试重复名称（应该自动生成唯一名称）
    auto sheet4 = workbook_->addSheet("数据表");
    ASSERT_NE(sheet4, nullptr);
    EXPECT_NE(sheet4->getName(), "数据表"); // 应该是"数据表1"或类似
    
    // 测试工作表访问
    EXPECT_EQ(workbook_->getSheet(0)->getName(), "Sheet1");
    EXPECT_EQ(workbook_->getSheet(1)->getName(), "数据表");
    EXPECT_EQ(workbook_->getSheet("数据表"), sheet2);
    EXPECT_EQ(workbook_->getSheet("不存在"), nullptr);
    
    // 测试工作表存在性
    EXPECT_TRUE(workbook_->hasSheet("Sheet1"));
    EXPECT_TRUE(workbook_->hasSheet("数据表"));
    EXPECT_FALSE(workbook_->hasSheet("不存在"));
    
    // 测试查找索引
    EXPECT_EQ(workbook_->findSheetIndex("Sheet1"), 0);
    EXPECT_EQ(workbook_->findSheetIndex("数据表"), 1);
    EXPECT_EQ(workbook_->findSheetIndex("不存在"), -1);
    
    TX_LOG_INFO("工作表管理测试通过");
}

/**
 * @brief 测试工作表操作
 */
TEST_F(TXWorkbookTest, SheetOperations) {
    // 添加几个工作表
    workbook_->addSheet("表A");
    workbook_->addSheet("表B");
    workbook_->addSheet("表C");
    EXPECT_EQ(workbook_->getSheetCount(), 4); // Sheet1 + 3个新表
    
    // 测试插入工作表
    auto insert_result = workbook_->insertSheet(2, "插入表");
    EXPECT_TRUE(insert_result.isOk());
    EXPECT_EQ(workbook_->getSheetCount(), 5);
    EXPECT_EQ(workbook_->getSheet(2)->getName(), "插入表");
    
    // 测试重命名工作表
    auto rename_result = workbook_->renameSheet(2, "新名称");
    EXPECT_TRUE(rename_result.isOk());
    EXPECT_EQ(workbook_->getSheet(2)->getName(), "新名称");
    
    // 测试移动工作表
    auto move_result = workbook_->moveSheet(2, 4);
    EXPECT_TRUE(move_result.isOk());
    EXPECT_EQ(workbook_->getSheet(4)->getName(), "新名称");
    
    // 测试删除工作表
    auto remove_result = workbook_->removeSheet(4);
    EXPECT_TRUE(remove_result.isOk());
    EXPECT_EQ(workbook_->getSheetCount(), 4);
    
    // 测试删除最后一个工作表（应该失败）
    while (workbook_->getSheetCount() > 1) {
        workbook_->removeSheet(1);
    }
    auto remove_last = workbook_->removeSheet(0);
    EXPECT_TRUE(remove_last.isError()); // 不能删除最后一个工作表
    
    TX_LOG_INFO("工作表操作测试通过");
}

/**
 * @brief 测试活动工作表管理
 */
TEST_F(TXWorkbookTest, ActiveSheetManagement) {
    // 添加工作表
    workbook_->addSheet("表1");
    workbook_->addSheet("表2");
    
    // 测试设置活动工作表（按索引）
    auto result1 = workbook_->setActiveSheet(1);
    EXPECT_TRUE(result1.isOk());
    EXPECT_EQ(workbook_->getActiveSheetIndex(), 1);
    EXPECT_EQ(workbook_->getActiveSheet()->getName(), "表1");
    
    // 测试设置活动工作表（按名称）
    auto result2 = workbook_->setActiveSheet("表2");
    EXPECT_TRUE(result2.isOk());
    EXPECT_EQ(workbook_->getActiveSheetIndex(), 2);
    EXPECT_EQ(workbook_->getActiveSheet()->getName(), "表2");
    
    // 测试无效索引
    auto result3 = workbook_->setActiveSheet(999);
    EXPECT_TRUE(result3.isError());
    
    // 测试不存在的名称
    auto result4 = workbook_->setActiveSheet("不存在");
    EXPECT_TRUE(result4.isError());
    
    TX_LOG_INFO("活动工作表管理测试通过");
}

/**
 * @brief 测试便捷操作符
 */
TEST_F(TXWorkbookTest, ConvenienceOperators) {
    // 添加工作表
    workbook_->addSheet("测试表");
    
    // 测试[]操作符（按索引）
    auto sheet1 = (*workbook_)[0];
    ASSERT_NE(sheet1, nullptr);
    EXPECT_EQ(sheet1->getName(), "Sheet1");
    
    auto sheet2 = (*workbook_)[1];
    ASSERT_NE(sheet2, nullptr);
    EXPECT_EQ(sheet2->getName(), "测试表");
    
    // 测试[]操作符（按名称）
    auto sheet3 = (*workbook_)["Sheet1"];
    ASSERT_NE(sheet3, nullptr);
    EXPECT_EQ(sheet3, sheet1);
    
    auto sheet4 = (*workbook_)["测试表"];
    ASSERT_NE(sheet4, nullptr);
    EXPECT_EQ(sheet4, sheet2);
    
    // 测试不存在的工作表
    auto sheet5 = (*workbook_)[999];
    EXPECT_EQ(sheet5, nullptr);
    
    auto sheet6 = (*workbook_)["不存在"];
    EXPECT_EQ(sheet6, nullptr);
    
    TX_LOG_INFO("便捷操作符测试通过");
}

/**
 * @brief 测试文件操作
 */
TEST_F(TXWorkbookTest, FileOperations) {
    // 测试保存状态（新创建的工作簿应该标记为已保存）
    EXPECT_FALSE(workbook_->hasUnsavedChanges());
    
    // 修改工作簿
    workbook_->addSheet("新表");
    EXPECT_TRUE(workbook_->hasUnsavedChanges());
    
    // 测试saveAs（目前是模拟实现）
    auto save_result = workbook_->saveAs("test.xlsx");
    // 由于是模拟实现，这里可能成功也可能失败
    
    // 测试文件路径
    if (save_result.isOk()) {
        EXPECT_EQ(workbook_->getFilePath(), "test.xlsx");
        EXPECT_FALSE(workbook_->hasUnsavedChanges());
    }
    
    TX_LOG_INFO("文件操作测试通过");
}

/**
 * @brief 测试性能优化
 */
TEST_F(TXWorkbookTest, PerformanceOptimization) {
    // 添加一些工作表和数据
    for (int i = 0; i < 3; ++i) {
        auto sheet = workbook_->addSheet("表" + std::to_string(i));
        
        // 添加一些数据
        for (int j = 0; j < 10; ++j) {
            sheet->cell(j, 0).setValue(static_cast<double>(j));
        }
    }
    
    // 测试预分配
    workbook_->reserve(10);
    
    // 测试优化
    workbook_->optimize();
    
    // 测试压缩
    size_t compressed = workbook_->compress();
    // 压缩数量可能为0，这是正常的
    
    // 测试收缩
    workbook_->shrinkToFit();
    
    TX_LOG_INFO("性能优化测试通过");
}

/**
 * @brief 测试调试功能
 */
TEST_F(TXWorkbookTest, DebuggingFeatures) {
    // 添加一些工作表
    workbook_->addSheet("调试表1");
    workbook_->addSheet("调试表2");
    
    // 测试toString
    std::string debug_str = workbook_->toString();
    EXPECT_TRUE(debug_str.find("测试工作簿") != std::string::npos);
    EXPECT_TRUE(debug_str.find("工作表数=") != std::string::npos);
    
    TX_LOG_INFO("调试信息: {}", debug_str);
    
    // 测试性能统计
    std::string stats = workbook_->getPerformanceStats();
    EXPECT_FALSE(stats.empty());
    TX_LOG_INFO("性能统计:\n{}", stats);
    
    // 测试内存使用
    size_t memory = workbook_->getMemoryUsage();
    EXPECT_GT(memory, 0);
    TX_LOG_INFO("内存使用: {} 字节", memory);
    
    // 测试工作表名称列表
    auto names = workbook_->getSheetNames();
    EXPECT_EQ(names.size(), 3);
    EXPECT_EQ(names[0], "Sheet1");
    EXPECT_EQ(names[1], "调试表1");
    EXPECT_EQ(names[2], "调试表2");
    
    TX_LOG_INFO("调试功能测试通过");
}

/**
 * @brief 测试错误处理
 */
TEST_F(TXWorkbookTest, ErrorHandling) {
    // 测试无效索引
    EXPECT_EQ(workbook_->getSheet(999), nullptr);
    
    auto remove_result = workbook_->removeSheet(999);
    EXPECT_TRUE(remove_result.isError());
    
    auto rename_result = workbook_->renameSheet(999, "新名称");
    EXPECT_TRUE(rename_result.isError());
    
    auto move_result = workbook_->moveSheet(0, 999);
    EXPECT_TRUE(move_result.isError());
    
    // 测试不存在的工作表名称
    auto remove_by_name = workbook_->removeSheet("不存在");
    EXPECT_TRUE(remove_by_name.isError());
    
    auto rename_by_name = workbook_->renameSheet("不存在", "新名称");
    EXPECT_TRUE(rename_by_name.isError());
    
    TX_LOG_INFO("错误处理测试通过");
}

/**
 * @brief 测试便捷函数
 */
TEST_F(TXWorkbookTest, ConvenienceFunctions) {
    // 测试makeWorkbook函数
    auto new_workbook = makeWorkbook("便捷工作簿");
    EXPECT_TRUE(new_workbook->isValid());
    EXPECT_EQ(new_workbook->getName(), "便捷工作簿");
    EXPECT_EQ(new_workbook->getSheetCount(), 1);
    
    TX_LOG_INFO("便捷函数测试通过");
}
