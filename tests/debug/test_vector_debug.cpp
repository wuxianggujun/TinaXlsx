//
// @file test_vector_debug.cpp
// @brief Vector越界问题调试测试
//

#include <gtest/gtest.h>
#include <TinaXlsx/TinaXlsx.hpp>
#include <iostream>
#include <vector>
#include <filesystem>
#include <fstream>

using namespace TinaXlsx;

/**
 * @brief Vector越界问题调试测试类
 */
class VectorDebugTest : public ::testing::Test {
protected:
    void SetUp() override {
        std::cout << "=== 开始Vector调试测试 ===" << std::endl;

        // 显示当前工作目录
        auto current_path = std::filesystem::current_path();
        std::cout << "当前工作目录: " << current_path << std::endl;

        // 初始化库
        ASSERT_TRUE(TinaXlsx::initialize()) << "库初始化失败";
        std::cout << "✓ 库初始化成功" << std::endl;
    }

    void TearDown() override {
        // 清理资源
        TinaXlsx::cleanup();
        std::cout << "=== Vector调试测试结束 ===" << std::endl;
    }

    // 辅助函数：验证文件是否存在且有效
    bool verifyExcelFile(const std::string& filename) {
        if (!std::filesystem::exists(filename)) {
            std::cout << "❌ 文件不存在: " << filename << std::endl;
            return false;
        }

        auto file_size = std::filesystem::file_size(filename);
        std::cout << "✓ 文件存在: " << filename << " (大小: " << file_size << " 字节)" << std::endl;

        // 检查文件是否为空
        if (file_size == 0) {
            std::cout << "❌ 文件为空" << std::endl;
            return false;
        }

        // 检查文件是否以PK开头（ZIP文件标识）
        std::ifstream file(filename, std::ios::binary);
        if (!file) {
            std::cout << "❌ 无法打开文件" << std::endl;
            return false;
        }

        char header[2];
        file.read(header, 2);
        if (header[0] == 'P' && header[1] == 'K') {
            std::cout << "✓ 文件格式正确 (ZIP/XLSX)" << std::endl;
            return true;
        } else {
            std::cout << "❌ 文件格式错误，不是有效的XLSX文件" << std::endl;
            return false;
        }
    }
};

/**
 * @brief 测试最基本的工作簿创建
 */
TEST_F(VectorDebugTest, BasicWorkbookCreation) {
    std::cout << "\n--- 测试基本工作簿创建 ---" << std::endl;
    
    try {
        std::cout << "1. 创建工作簿..." << std::endl;
        auto workbook = TXInMemoryWorkbook::create("debug_basic.xlsx");
        ASSERT_NE(workbook, nullptr) << "工作簿创建失败";
        std::cout << "✓ 工作簿创建成功" << std::endl;
        
        std::cout << "2. 创建工作表..." << std::endl;
        auto& sheet = workbook->createSheet("调试测试");
        std::cout << "✓ 工作表创建成功: " << sheet.getName() << std::endl;
        
    } catch (const std::exception& e) {
        FAIL() << "异常: " << e.what();
    }
}

/**
 * @brief 测试单个单元格设置
 */
TEST_F(VectorDebugTest, SingleCellSet) {
    std::cout << "\n--- 测试单个单元格设置 ---" << std::endl;
    
    try {
        auto workbook = TXInMemoryWorkbook::create("debug_single.xlsx");
        auto& sheet = workbook->createSheet("单元格测试");
        
        std::cout << "1. 准备坐标和数值..." << std::endl;
        std::vector<TXCoordinate> coords = {
            TXCoordinate(row_t(0u), column_t(0u))
        };
        std::vector<double> values = {42.0};
        
        std::cout << "2. 设置单元格..." << std::endl;
        auto result = sheet.setBatchNumbers(coords, values);
        
        ASSERT_TRUE(result.isOk()) << "设置单元格失败: " << result.error().getMessage();
        EXPECT_EQ(result.value(), 1) << "应该设置1个单元格";
        std::cout << "✓ 单元格设置成功，设置了 " << result.value() << " 个单元格" << std::endl;
        
    } catch (const std::exception& e) {
        FAIL() << "异常: " << e.what();
    }
}

/**
 * @brief 测试批量单元格设置
 */
TEST_F(VectorDebugTest, BatchCellSet) {
    std::cout << "\n--- 测试批量单元格设置 ---" << std::endl;
    
    try {
        auto workbook = TXInMemoryWorkbook::create("debug_batch.xlsx");
        auto& sheet = workbook->createSheet("批量测试");
        
        std::cout << "1. 准备批量数据..." << std::endl;
        constexpr size_t CELL_COUNT = 10;
        std::vector<TXCoordinate> coords;
        std::vector<double> values;
        
        coords.reserve(CELL_COUNT);
        values.reserve(CELL_COUNT);
        
        for (size_t i = 0; i < CELL_COUNT; ++i) {
            coords.emplace_back(row_t(static_cast<uint32_t>(i)), column_t(0u));
            values.push_back(i * 10.0);
        }
        
        std::cout << "2. 批量设置单元格..." << std::endl;
        auto result = sheet.setBatchNumbers(coords, values);
        
        ASSERT_TRUE(result.isOk()) << "批量设置失败: " << result.error().getMessage();
        EXPECT_EQ(result.value(), CELL_COUNT) << "应该设置" << CELL_COUNT << "个单元格";
        std::cout << "✓ 批量设置成功，设置了 " << result.value() << " 个单元格" << std::endl;
        
    } catch (const std::exception& e) {
        FAIL() << "异常: " << e.what();
    }
}

/**
 * @brief 测试文件保存
 */
TEST_F(VectorDebugTest, FileSave) {
    std::cout << "\n--- 测试文件保存 ---" << std::endl;
    
    try {
        auto workbook = TXInMemoryWorkbook::create("debug_save.xlsx");
        auto& sheet = workbook->createSheet("保存测试");
        
        // 设置一些数据
        std::vector<TXCoordinate> coords = {
            TXCoordinate(row_t(0u), column_t(0u)),
            TXCoordinate(row_t(0u), column_t(1u)),
            TXCoordinate(row_t(1u), column_t(0u))
        };
        std::vector<double> values = {1.0, 2.0, 3.0};
        
        auto set_result = sheet.setBatchNumbers(coords, values);
        ASSERT_TRUE(set_result.isOk()) << "设置数据失败";
        
        std::cout << "1. 保存文件..." << std::endl;
        auto save_result = workbook->saveToFile();

        ASSERT_TRUE(save_result.isOk()) << "保存文件失败: " << save_result.error().getMessage();
        std::cout << "✓ 文件保存成功" << std::endl;

        // 验证文件
        std::cout << "2. 验证文件..." << std::endl;
        EXPECT_TRUE(verifyExcelFile("debug_save.xlsx")) << "文件验证失败";
        
    } catch (const std::exception& e) {
        FAIL() << "异常: " << e.what();
    }
}

/**
 * @brief 测试字符串单元格
 */
TEST_F(VectorDebugTest, StringCells) {
    std::cout << "\n--- 测试字符串单元格 ---" << std::endl;
    
    try {
        auto workbook = TXInMemoryWorkbook::create("debug_strings.xlsx");
        auto& sheet = workbook->createSheet("字符串测试");
        
        std::cout << "1. 准备字符串数据..." << std::endl;
        std::vector<TXCoordinate> coords = {
            TXCoordinate(row_t(0u), column_t(0u)),
            TXCoordinate(row_t(0u), column_t(1u))
        };
        std::vector<std::string> strings = {"Hello", "World"};
        
        std::cout << "2. 设置字符串单元格..." << std::endl;
        auto result = sheet.setBatchStrings(coords, strings);
        
        ASSERT_TRUE(result.isOk()) << "设置字符串失败: " << result.error().getMessage();
        EXPECT_EQ(result.value(), 2) << "应该设置2个字符串单元格";
        std::cout << "✓ 字符串设置成功，设置了 " << result.value() << " 个单元格" << std::endl;
        
    } catch (const std::exception& e) {
        FAIL() << "异常: " << e.what();
    }
}

/**
 * @brief 压力测试 - 大量数据
 */
TEST_F(VectorDebugTest, StressTest) {
    std::cout << "\n--- 压力测试 ---" << std::endl;
    
    try {
        auto workbook = TXInMemoryWorkbook::create("debug_stress.xlsx");
        auto& sheet = workbook->createSheet("压力测试");
        
        std::cout << "1. 准备大量数据..." << std::endl;
        constexpr size_t LARGE_COUNT = 1000;
        std::vector<TXCoordinate> coords;
        std::vector<double> values;
        
        coords.reserve(LARGE_COUNT);
        values.reserve(LARGE_COUNT);
        
        for (size_t i = 0; i < LARGE_COUNT; ++i) {
            coords.emplace_back(row_t(static_cast<uint32_t>(i / 10)), column_t(static_cast<uint32_t>(i % 10)));
            values.push_back(i * 0.1);
        }
        
        std::cout << "2. 批量设置大量数据..." << std::endl;
        auto result = sheet.setBatchNumbers(coords, values);
        
        ASSERT_TRUE(result.isOk()) << "大量数据设置失败: " << result.error().getMessage();
        EXPECT_EQ(result.value(), LARGE_COUNT) << "应该设置" << LARGE_COUNT << "个单元格";
        std::cout << "✓ 大量数据设置成功，设置了 " << result.value() << " 个单元格" << std::endl;
        
        std::cout << "3. 保存大文件..." << std::endl;
        auto save_result = workbook->saveToFile();
        ASSERT_TRUE(save_result.isOk()) << "大文件保存失败";
        std::cout << "✓ 大文件保存成功" << std::endl;

        // 验证大文件
        std::cout << "4. 验证大文件..." << std::endl;
        EXPECT_TRUE(verifyExcelFile("debug_stress.xlsx")) << "大文件验证失败";
        
    } catch (const std::exception& e) {
        FAIL() << "异常: " << e.what();
    }
}

/**
 * @brief 测试文件验证和列出生成的文件
 */
TEST_F(VectorDebugTest, FileVerification) {
    std::cout << "\n--- 文件验证测试 ---" << std::endl;

    try {
        // 创建一个简单的测试文件
        auto workbook = TXInMemoryWorkbook::create("verification_test.xlsx");
        auto& sheet = workbook->createSheet("验证测试");

        // 添加一些测试数据
        std::vector<TXCoordinate> coords = {
            TXCoordinate(row_t(0u), column_t(0u)),  // A1
            TXCoordinate(row_t(0u), column_t(1u)),  // B1
            TXCoordinate(row_t(1u), column_t(0u)),  // A2
        };

        std::vector<double> numbers = {1.0, 2.0, 3.0};
        auto result = sheet.setBatchNumbers(coords, numbers);
        ASSERT_TRUE(result.isOk()) << "设置数据失败";

        // 添加字符串数据
        std::vector<TXCoordinate> str_coords = {
            TXCoordinate(row_t(1u), column_t(1u))  // B2
        };
        std::vector<std::string> strings = {"测试字符串"};
        auto str_result = sheet.setBatchStrings(str_coords, strings);
        ASSERT_TRUE(str_result.isOk()) << "设置字符串失败";

        // 保存文件
        std::cout << "1. 保存验证测试文件..." << std::endl;
        auto save_result = workbook->saveToFile();
        ASSERT_TRUE(save_result.isOk()) << "保存失败";

        // 验证文件
        std::cout << "2. 验证文件..." << std::endl;
        EXPECT_TRUE(verifyExcelFile("verification_test.xlsx")) << "文件验证失败";

        // 列出当前目录中的所有xlsx文件
        std::cout << "3. 当前目录中的XLSX文件:" << std::endl;
        for (const auto& entry : std::filesystem::directory_iterator(".")) {
            if (entry.is_regular_file() && entry.path().extension() == ".xlsx") {
                auto file_size = std::filesystem::file_size(entry.path());
                std::cout << "  📄 " << entry.path().filename()
                         << " (大小: " << file_size << " 字节)" << std::endl;
            }
        }

    } catch (const std::exception& e) {
        FAIL() << "异常: " << e.what();
    }
}

// 主函数 - 支持独立运行
int main(int argc, char** argv) {
    ::testing::InitGoogleTest(&argc, argv);
    
    std::cout << "🔍 Vector越界问题调试测试" << std::endl;
    std::cout << "目标：找到并修复vector subscript out of range错误" << std::endl;
    
    return RUN_ALL_TESTS();
}
