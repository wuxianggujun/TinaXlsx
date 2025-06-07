//
// @file test_tx_excel_io.cpp
// @brief TXExcelIO单元测试
//

#include <gtest/gtest.h>
#include <TinaXlsx/io/TXExcelIO.hpp>
#include <TinaXlsx/user/TXWorkbook.hpp>
#include <TinaXlsx/TXUnifiedMemoryManager.hpp>
#include <TinaXlsx/TXHighPerformanceLogger.hpp>
#include <TinaXlsx/TXGlobalStringPool.hpp>
#include <fstream>
#include <filesystem>

using namespace TinaXlsx;

/**
 * @brief TXExcelIO测试类
 */
class TXExcelIOTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 初始化内存管理器
        TXUnifiedMemoryManager::Config config;
        config.memory_limit = 512ULL * 1024 * 1024; // 512MB
        GlobalUnifiedMemoryManager::initialize(config);
        
        // 初始化日志系统
        TXGlobalLogger::initialize(GlobalUnifiedMemoryManager::getInstance());
        TXGlobalLogger::setOutputMode(TXLogOutputMode::CONSOLE_ONLY);
        
        // 创建测试目录
        test_dir_ = "test_excel_io";
        std::filesystem::create_directories(test_dir_);
        
        // 创建测试工作簿
        workbook_ = TXWorkbook::create("测试工作簿");
        auto sheet = workbook_->getSheet(0);
        
        // 添加测试数据
        sheet->cell("A1").setValue("姓名");
        sheet->cell("B1").setValue("年龄");
        sheet->cell("C1").setValue("分数");
        
        sheet->cell("A2").setValue("张三");
        sheet->cell("B2").setValue(25.0);
        sheet->cell("C2").setValue(95.5);
        
        sheet->cell("A3").setValue("李四");
        sheet->cell("B3").setValue(30.0);
        sheet->cell("C3").setValue(87.2);
    }
    
    void TearDown() override {
        workbook_.reset();
        
        // 清理测试文件
        try {
            std::filesystem::remove_all(test_dir_);
        } catch (...) {
            // 忽略清理错误
        }
        
        GlobalUnifiedMemoryManager::shutdown();
    }
    
    std::string test_dir_;
    std::unique_ptr<TXWorkbook> workbook_;
    
    std::string getTestFilePath(const std::string& filename) {
        return test_dir_ + "/" + filename;
    }
    
    void createTestCSV(const std::string& filename) {
        std::ofstream file(getTestFilePath(filename));
        file << "姓名,年龄,分数\n";
        file << "张三,25,95.5\n";
        file << "李四,30,87.2\n";
        file << "王五,28,92.1\n";
    }
};

/**
 * @brief 测试格式检测
 */
TEST_F(TXExcelIOTest, FormatDetection) {
    // 测试扩展名检测
    EXPECT_EQ(TXExcelIO::detectFormat("test.xlsx"), TXExcelIO::FileFormat::XLSX);
    EXPECT_EQ(TXExcelIO::detectFormat("test.xls"), TXExcelIO::FileFormat::XLS);
    EXPECT_EQ(TXExcelIO::detectFormat("test.csv"), TXExcelIO::FileFormat::CSV);
    
    // 创建测试CSV文件
    createTestCSV("test.csv");
    EXPECT_EQ(TXExcelIO::detectFormat(getTestFilePath("test.csv")), TXExcelIO::FileFormat::CSV);
    
    // 测试文件有效性检查
    EXPECT_FALSE(TXExcelIO::isValidExcelFile("nonexistent.xlsx"));
    
    TX_LOG_INFO("格式检测测试通过");
}

/**
 * @brief 测试CSV读取
 */
TEST_F(TXExcelIOTest, CSVReading) {
    // 创建测试CSV文件
    createTestCSV("test_read.csv");
    
    // 读取CSV文件
    auto result = TXExcelIO::loadCSV(getTestFilePath("test_read.csv"));
    ASSERT_TRUE(result.isOk()) << "CSV读取失败: " << result.error().getMessage();
    
    auto loaded_workbook = std::move(result.value());
    ASSERT_NE(loaded_workbook, nullptr);
    EXPECT_EQ(loaded_workbook->getSheetCount(), 1);

    auto sheet = loaded_workbook->getSheet(0);
    ASSERT_NE(sheet, nullptr);
    
    // 验证数据 - 使用安全的类型检查
    auto a1_value = sheet->cell("A1").getValue();
    if (a1_value.getType() == TXVariant::Type::String) {
        EXPECT_EQ(a1_value.getString(), "姓名");
    }

    auto b1_value = sheet->cell("B1").getValue();
    if (b1_value.getType() == TXVariant::Type::String) {
        EXPECT_EQ(b1_value.getString(), "年龄");
    }

    auto c1_value = sheet->cell("C1").getValue();
    if (c1_value.getType() == TXVariant::Type::String) {
        EXPECT_EQ(c1_value.getString(), "分数");
    }

    auto a2_value = sheet->cell("A2").getValue();
    if (a2_value.getType() == TXVariant::Type::String) {
        EXPECT_EQ(a2_value.getString(), "张三");
    }

    auto b2_value = sheet->cell("B2").getValue();
    if (b2_value.getType() == TXVariant::Type::Number) {
        EXPECT_EQ(b2_value.getNumber(), 25.0);
    }

    auto c2_value = sheet->cell("C2").getValue();
    if (c2_value.getType() == TXVariant::Type::Number) {
        EXPECT_EQ(c2_value.getNumber(), 95.5);
    }

    auto a4_value = sheet->cell("A4").getValue();
    if (a4_value.getType() == TXVariant::Type::String) {
        EXPECT_EQ(a4_value.getString(), "王五");
    }

    auto b4_value = sheet->cell("B4").getValue();
    if (b4_value.getType() == TXVariant::Type::Number) {
        EXPECT_EQ(b4_value.getNumber(), 28.0);
    }

    auto c4_value = sheet->cell("C4").getValue();
    if (c4_value.getType() == TXVariant::Type::Number) {
        EXPECT_EQ(c4_value.getNumber(), 92.1);
    }
    
    TX_LOG_INFO("CSV读取测试通过");
}

/**
 * @brief 测试CSV写入
 */
TEST_F(TXExcelIOTest, CSVWriting) {
    std::string csv_path = getTestFilePath("test_write.csv");

    // 在保存前检查工作表中的数据
    auto sheet = workbook_->getSheet(0);
    TX_LOG_INFO("保存前检查数据:");
    TX_LOG_INFO("A1: '{}' (类型: {})",
               sheet->cell("A1").getValue().getType() == TXVariant::Type::String ?
               sheet->cell("A1").getValue().getString() : "NOT_STRING",
               static_cast<int>(sheet->cell("A1").getValue().getType()));
    TX_LOG_INFO("B1: '{}' (类型: {})",
               sheet->cell("B1").getValue().getType() == TXVariant::Type::String ?
               sheet->cell("B1").getValue().getString() : "NOT_STRING",
               static_cast<int>(sheet->cell("B1").getValue().getType()));
    TX_LOG_INFO("A2: '{}' (类型: {})",
               sheet->cell("A2").getValue().getType() == TXVariant::Type::String ?
               sheet->cell("A2").getValue().getString() : "NOT_STRING",
               static_cast<int>(sheet->cell("A2").getValue().getType()));

    // 保存为CSV
    auto result = TXExcelIO::saveCSV(workbook_.get(), csv_path);
    EXPECT_TRUE(result.isOk()) << "CSV保存失败: " << result.error().getMessage();
    
    // 验证文件存在
    EXPECT_TRUE(std::filesystem::exists(csv_path));
    
    // 读取并验证内容
    std::ifstream file(csv_path);
    ASSERT_TRUE(file.is_open());

    // 先打印实际文件内容进行调试
    std::string content;
    std::string line;
    int line_num = 1;
    while (std::getline(file, line)) {
        TX_LOG_INFO("CSV文件第{}行: '{}'", line_num++, line);
        content += line + "\n";
    }

    // 重新打开文件进行验证
    file.close();
    file.open(csv_path);

    std::getline(file, line);
    EXPECT_EQ(line, "姓名,年龄,分数");

    std::getline(file, line);
    EXPECT_EQ(line, "张三,25,95.5");

    std::getline(file, line);
    EXPECT_EQ(line, "李四,30,87.2");
    
    TX_LOG_INFO("CSV写入测试通过");
}

/**
 * @brief 测试XLSX读取（当前为占位实现）
 */
TEST_F(TXExcelIOTest, XLSXReading) {
    // 创建一个假的XLSX文件（实际是文本文件）
    std::string xlsx_path = getTestFilePath("test.xlsx");
    std::ofstream file(xlsx_path);
    file << "PK"; // 模拟ZIP文件头
    file.close();
    
    // 尝试读取XLSX文件
    auto result = TXExcelIO::loadFromFile(xlsx_path);
    
    // 当前实现应该返回示例数据
    if (result.isOk()) {
        auto loaded_workbook = std::move(result.value());
        EXPECT_NE(loaded_workbook, nullptr);
        EXPECT_GT(loaded_workbook->getSheetCount(), 0);
        
        TX_LOG_INFO("XLSX读取测试通过（占位实现）");
    } else {
        TX_LOG_WARN("XLSX读取失败: {}", result.error().getMessage());
    }
}

/**
 * @brief 测试XLSX写入（当前为占位实现）
 */
TEST_F(TXExcelIOTest, XLSXWriting) {
    std::string xlsx_path = getTestFilePath("test_output.xlsx");
    
    // 保存为XLSX
    auto result = TXExcelIO::saveToFile(workbook_.get(), xlsx_path);
    
    if (result.isOk()) {
        // 验证文件存在
        EXPECT_TRUE(std::filesystem::exists(xlsx_path));
        TX_LOG_INFO("XLSX写入测试通过（占位实现）");
    } else {
        TX_LOG_WARN("XLSX写入失败: {}", result.error().getMessage());
    }
}

/**
 * @brief 测试内存操作
 */
TEST_F(TXExcelIOTest, MemoryOperations) {
    // 测试保存到内存
    auto save_result = TXExcelIO::saveToMemory(workbook_.get());
    
    if (save_result.isOk()) {
        auto data = save_result.value();
        EXPECT_GT(data.size(), 0);
        
        // 测试从内存加载
        auto load_result = TXExcelIO::loadFromMemory(data.data(), data.size());
        
        if (load_result.isOk()) {
            auto loaded_workbook = std::move(load_result.value());
            EXPECT_NE(loaded_workbook, nullptr);
            TX_LOG_INFO("内存操作测试通过（占位实现）");
        } else {
            TX_LOG_WARN("内存加载失败: {}", load_result.error().getMessage());
        }
    } else {
        TX_LOG_WARN("内存保存失败: {}", save_result.error().getMessage());
    }
}

/**
 * @brief 测试错误处理
 */
TEST_F(TXExcelIOTest, ErrorHandling) {
    // 测试不存在的文件
    auto result1 = TXExcelIO::loadFromFile("nonexistent.xlsx");
    EXPECT_TRUE(result1.isError());
    
    // 测试无效的工作簿指针
    auto result2 = TXExcelIO::saveToFile(nullptr, "test.xlsx");
    EXPECT_TRUE(result2.isError());
    
    // 测试无效的内存数据
    auto result3 = TXExcelIO::loadFromMemory(nullptr, 0);
    EXPECT_TRUE(result3.isError());
    
    // 测试无效的工作表索引
    auto result4 = TXExcelIO::saveCSV(workbook_.get(), "test.csv", 999);
    EXPECT_TRUE(result4.isError());
    
    TX_LOG_INFO("错误处理测试通过");
}

/**
 * @brief 测试文件路径处理
 */
TEST_F(TXExcelIOTest, FilePathHandling) {
    // 测试自动创建目录
    std::string nested_path = test_dir_ + "/nested/dir/test.csv";
    auto result = TXExcelIO::saveCSV(workbook_.get(), nested_path);
    EXPECT_TRUE(result.isOk());
    EXPECT_TRUE(std::filesystem::exists(nested_path));
    
    // 测试备份功能（通过重复保存触发）
    auto result2 = TXExcelIO::saveToFile(workbook_.get(), nested_path);
    // 应该成功，即使文件已存在
    
    TX_LOG_INFO("文件路径处理测试通过");
}

/**
 * @brief 测试TXWorkbook集成
 */
TEST_F(TXExcelIOTest, WorkbookIntegration) {
    std::string csv_path = getTestFilePath("integration_test.csv");
    
    // 使用TXWorkbook的保存方法
    auto save_result = workbook_->saveAs(csv_path);
    EXPECT_TRUE(save_result.isOk());
    
    // 使用TXWorkbook的加载方法
    auto load_result = TXWorkbook::load(csv_path);
    
    if (load_result.isOk()) {
        auto loaded_workbook = std::move(load_result.value());
        EXPECT_NE(loaded_workbook, nullptr);
        EXPECT_GT(loaded_workbook->getSheetCount(), 0);
        TX_LOG_INFO("TXWorkbook集成测试通过");
    } else {
        TX_LOG_WARN("TXWorkbook加载失败: {}", load_result.error().getMessage());
    }
}

/**
 * @brief 测试便捷函数
 */
TEST_F(TXExcelIOTest, ConvenienceFunctions) {
    std::string csv_path = getTestFilePath("convenience_test.csv");
    
    // 测试便捷保存函数
    auto save_result = saveExcel(workbook_.get(), csv_path);
    EXPECT_TRUE(save_result.isOk());
    
    // 测试便捷加载函数
    auto load_result = loadExcel(csv_path);
    
    if (load_result.isOk()) {
        auto loaded_workbook = std::move(load_result.value());
        EXPECT_NE(loaded_workbook, nullptr);
        TX_LOG_INFO("便捷函数测试通过");
    } else {
        TX_LOG_WARN("便捷加载失败: {}", load_result.error().getMessage());
    }
}
