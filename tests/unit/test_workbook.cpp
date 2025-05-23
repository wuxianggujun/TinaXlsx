#include <gtest/gtest.h>
#include <gmock/gmock.h>
#include "TinaXlsx/Workbook.hpp"
#include "TinaXlsx/Exception.hpp"
#include <filesystem>
#include <fstream>

namespace TinaXlsx {
namespace Test {

/**
 * @brief Workbook类的单元测试
 */
class WorkbookTest : public ::testing::Test {
protected:
    void SetUp() override {
        test_filename = "test_workbook.xlsx";
        // 确保测试文件不存在
        if (std::filesystem::exists(test_filename)) {
            std::filesystem::remove(test_filename);
        }
    }

    void TearDown() override {
        // 清理测试文件
        if (std::filesystem::exists(test_filename)) {
            std::filesystem::remove(test_filename);
        }
    }

    std::string test_filename;
};

/**
 * @brief 测试Workbook的写入模式构造函数
 */
TEST_F(WorkbookTest, WriteConstructor) {
    EXPECT_NO_THROW({
        Workbook workbook(test_filename, Workbook::Mode::Write);
        EXPECT_EQ(workbook.getFilePath(), test_filename);
        EXPECT_EQ(workbook.getMode(), Workbook::Mode::Write);
    });
}

/**
 * @brief 测试Workbook的静态工厂方法
 */
TEST_F(WorkbookTest, StaticFactoryMethods) {
    EXPECT_NO_THROW({
        auto write_workbook = Workbook::createForWrite(test_filename);
        ASSERT_NE(write_workbook, nullptr);
        EXPECT_EQ(write_workbook->getFilePath(), test_filename);
        EXPECT_EQ(write_workbook->getMode(), Workbook::Mode::Write);
    });
}

/**
 * @brief 测试获取Writer
 */
TEST_F(WorkbookTest, GetWriter) {
    Workbook workbook(test_filename, Workbook::Mode::Write);
    
    EXPECT_NO_THROW({
        auto& writer = workbook.getWriter();
        // Writer应该是有效的
        (void)writer; // 避免未使用变量警告
    });
}

/**
 * @brief 测试关闭Workbook
 */
TEST_F(WorkbookTest, CloseWorkbook) {
    Workbook workbook(test_filename, Workbook::Mode::Write);
    
    EXPECT_FALSE(workbook.isClosed());
    
    EXPECT_NO_THROW({
        bool result = workbook.close();
        EXPECT_TRUE(result);
    });
    
    EXPECT_TRUE(workbook.isClosed());
    
    // 检查文件是否已创建
    EXPECT_TRUE(std::filesystem::exists(test_filename));
}

/**
 * @brief 测试保存Workbook
 */
TEST_F(WorkbookTest, SaveWorkbook) {
    Workbook workbook(test_filename, Workbook::Mode::Write);
    
    EXPECT_NO_THROW({
        bool result = workbook.save();
        EXPECT_TRUE(result);
    });
    
    // 检查文件是否已创建
    EXPECT_TRUE(std::filesystem::exists(test_filename));
}

/**
 * @brief 测试空文件名的处理
 */
TEST_F(WorkbookTest, EmptyFilename) {
    EXPECT_THROW({
        Workbook workbook("", Workbook::Mode::Write);
    }, Exception);
}

/**
 * @brief 测试模式检查方法
 */
TEST_F(WorkbookTest, ModeChecking) {
    // 测试写入模式
    Workbook write_workbook(test_filename, Workbook::Mode::Write);
    EXPECT_TRUE(write_workbook.canWrite());
    EXPECT_FALSE(write_workbook.canRead());
    
    // 关闭写入工作簿，这应该创建一个真正的Excel文件
    write_workbook.close();
    
    // 测试读取模式 - 使用Writer创建的真实Excel文件
    if (std::filesystem::exists(test_filename)) {
        EXPECT_NO_THROW({
            Workbook read_workbook(test_filename, Workbook::Mode::Read);
            EXPECT_FALSE(read_workbook.canWrite());
            EXPECT_TRUE(read_workbook.canRead());
        });
    }
}

/**
 * @brief 测试移动语义
 */
TEST_F(WorkbookTest, MoveSemantics) {
    EXPECT_NO_THROW({
        Workbook original(test_filename, Workbook::Mode::Write);
        std::string original_path = original.getFilePath();
        
        // 移动构造
        Workbook moved = std::move(original);
        EXPECT_EQ(moved.getFilePath(), original_path);
        EXPECT_EQ(moved.getMode(), Workbook::Mode::Write);
    });
}

/**
 * @brief 测试重复关闭
 */
TEST_F(WorkbookTest, DoubleClose) {
    Workbook workbook(test_filename, Workbook::Mode::Write);
    
    // 第一次关闭
    EXPECT_TRUE(workbook.close());
    EXPECT_TRUE(workbook.isClosed());
    
    // 第二次关闭应该是安全的
    EXPECT_NO_THROW({
        bool result = workbook.close();
        // 结果可能为false，但不应该抛出异常
        (void)result;
    });
}

} // namespace Test
} // namespace TinaXlsx 