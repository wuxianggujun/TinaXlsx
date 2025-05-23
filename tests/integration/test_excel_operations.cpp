#include <gtest/gtest.h>
#include <gmock/gmock.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "TinaXlsx/Workbook.hpp"
#include "TinaXlsx/Reader.hpp"
#include "TinaXlsx/Writer.hpp"
#include <filesystem>
#include <vector>

namespace TinaXlsx {
namespace Test {

/**
 * @brief Excel操作的集成测试
 */
class ExcelOperationsTest : public ::testing::Test {
protected:
    void SetUp() override {
        test_write_file = "test_write.xlsx";
        test_read_file = "test_read.xlsx";
        
        // 清理可能存在的测试文件
        cleanupTestFiles();
    }

    void TearDown() override {
        cleanupTestFiles();
    }

    void cleanupTestFiles() {
        if (std::filesystem::exists(test_write_file)) {
            std::filesystem::remove(test_write_file);
        }
        if (std::filesystem::exists(test_read_file)) {
            std::filesystem::remove(test_read_file);
        }
    }

    std::string test_write_file;
    std::string test_read_file;
};

/**
 * @brief 测试基本的Workbook创建和关闭
 */
TEST_F(ExcelOperationsTest, BasicWorkbookOperations) {
    // 创建写入模式的工作簿
    EXPECT_NO_THROW({
        Workbook workbook(test_write_file, Workbook::Mode::Write);
        
        // 获取Writer
        auto& writer = workbook.getWriter();
        (void)writer; // 避免未使用变量警告
        
        // 关闭工作簿
        bool result = workbook.close();
        EXPECT_TRUE(result);
    });
    
    // 验证文件已创建
    EXPECT_TRUE(std::filesystem::exists(test_write_file));
    EXPECT_GT(std::filesystem::file_size(test_write_file), 0);
}

/**
 * @brief 测试Workbook模式检查
 */
TEST_F(ExcelOperationsTest, WorkbookModeChecking) {
    // 测试写入模式
    {
        Workbook write_workbook(test_write_file, Workbook::Mode::Write);
        EXPECT_TRUE(write_workbook.canWrite());
        EXPECT_FALSE(write_workbook.canRead());
        EXPECT_EQ(write_workbook.getMode(), Workbook::Mode::Write);
        write_workbook.close();
    }
    
    // 创建一个临时文件用于读取模式测试
    if (std::filesystem::exists(test_write_file)) {
        Workbook read_workbook(test_write_file, Workbook::Mode::Read);
        EXPECT_FALSE(read_workbook.canWrite());
        EXPECT_TRUE(read_workbook.canRead());
        EXPECT_EQ(read_workbook.getMode(), Workbook::Mode::Read);
    }
}

/**
 * @brief 测试静态工厂方法
 */
TEST_F(ExcelOperationsTest, StaticFactoryMethods) {
    // 测试createForWrite
    EXPECT_NO_THROW({
        auto write_workbook = Workbook::createForWrite(test_write_file);
        ASSERT_NE(write_workbook, nullptr);
        EXPECT_EQ(write_workbook->getMode(), Workbook::Mode::Write);
        EXPECT_TRUE(write_workbook->canWrite());
        write_workbook->close();
    });
    
    // 验证文件已创建
    EXPECT_TRUE(std::filesystem::exists(test_write_file));
    
    // 测试openForRead
    if (std::filesystem::exists(test_write_file)) {
        EXPECT_NO_THROW({
            auto read_workbook = Workbook::openForRead(test_write_file);
            ASSERT_NE(read_workbook, nullptr);
            EXPECT_EQ(read_workbook->getMode(), Workbook::Mode::Read);
            EXPECT_TRUE(read_workbook->canRead());
        });
    }
}

/**
 * @brief 测试独立的Writer类
 */
TEST_F(ExcelOperationsTest, StandaloneWriter) {
    EXPECT_NO_THROW({
        Writer writer(test_write_file);
        // 基本的writer操作应该不抛出异常
        // 具体的写入测试需要依赖实际的Writer接口
        (void)writer; // 避免未使用变量警告
    });
}

/**
 * @brief 测试独立的Reader类
 */
TEST_F(ExcelOperationsTest, StandaloneReader) {
    // 首先创建一个文件
    {
        Writer writer(test_read_file);
        // 假设Writer有某种方式创建有效的Excel文件
    }
    
    if (std::filesystem::exists(test_read_file)) {
        EXPECT_NO_THROW({
            Reader reader(test_read_file);
            // 基本的reader操作应该不抛出异常
            (void)reader; // 避免未使用变量警告
        });
    }
}

/**
 * @brief 测试多个Workbook实例
 */
TEST_F(ExcelOperationsTest, MultipleWorkbooks) {
    std::string file1 = "test_multi1.xlsx";
    std::string file2 = "test_multi2.xlsx";
    
    // 确保清理
    auto cleanup = [&]() {
        if (std::filesystem::exists(file1)) {
            std::filesystem::remove(file1);
        }
        if (std::filesystem::exists(file2)) {
            std::filesystem::remove(file2);
        }
    };
    
    cleanup();
    
    EXPECT_NO_THROW({
        // 创建多个工作簿
        auto workbook1 = Workbook::createForWrite(file1);
        auto workbook2 = Workbook::createForWrite(file2);
        
        ASSERT_NE(workbook1, nullptr);
        ASSERT_NE(workbook2, nullptr);
        
        // 关闭工作簿
        workbook1->close();
        workbook2->close();
    });
    
    // 验证文件已创建
    EXPECT_TRUE(std::filesystem::exists(file1));
    EXPECT_TRUE(std::filesystem::exists(file2));
    
    cleanup();
}

/**
 * @brief 测试文件路径处理
 */
TEST_F(ExcelOperationsTest, FilePathHandling) {
    // 测试相对路径
    EXPECT_NO_THROW({
        Workbook workbook("./relative_path.xlsx", Workbook::Mode::Write);
        workbook.close();
    });
    
    // 清理相对路径文件
    if (std::filesystem::exists("./relative_path.xlsx")) {
        std::filesystem::remove("./relative_path.xlsx");
    }
    
    // 测试空路径（应该抛出异常）
    EXPECT_THROW({
        Workbook workbook("", Workbook::Mode::Write);
    }, Exception);
}

} // namespace Test
} // namespace TinaXlsx 