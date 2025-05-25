#include <gtest/gtest.h>
#include "TinaXlsx/TXZipHandler.hpp"
#include <fstream>
#include <filesystem>

class TXZipHandlerTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 创建测试目录
        std::filesystem::create_directories("test_files");
        
        // 创建一个测试文件
        std::ofstream test_file("test_files/test.txt");
        test_file << "Hello, World!";
        test_file.close();
    }

    void TearDown() override {
        // 清理测试文件
        std::filesystem::remove_all("test_files");
        std::filesystem::remove("test.zip");
        std::filesystem::remove("test_output.zip");
    }
};

TEST_F(TXZipHandlerTest, CreateAndWriteZip) {
    TinaXlsx::TXZipHandler zip;
    
    // 测试创建新的ZIP文件
    EXPECT_TRUE(zip.open("test.zip", TinaXlsx::TXZipHandler::OpenMode::Write));
    EXPECT_TRUE(zip.isOpen());
    
    // 写入文件
    std::string content = "This is a test file content.";
    EXPECT_TRUE(zip.writeFile("test.txt", content));
    
    // 写入多个文件
    std::unordered_map<std::string, std::string> files;
    files["file1.txt"] = "Content of file 1";
    files["file2.txt"] = "Content of file 2";
    files["dir/file3.txt"] = "Content of file 3 in subdirectory";
    
    std::size_t written = zip.writeMultipleFiles(files);
    EXPECT_EQ(written, 3);
    
    zip.close();
    EXPECT_FALSE(zip.isOpen());
    
    // 验证文件是否存在
    EXPECT_TRUE(std::filesystem::exists("test.zip"));
}

TEST_F(TXZipHandlerTest, ReadZip) {
    // 首先创建一个ZIP文件
    TinaXlsx::TXZipHandler write_zip;
    ASSERT_TRUE(write_zip.open("test.zip", TinaXlsx::TXZipHandler::OpenMode::Write));
    
    std::string test_content = "Test content for reading";
    ASSERT_TRUE(write_zip.writeFile("read_test.txt", test_content));
    
    std::vector<uint8_t> binary_data = {0x48, 0x65, 0x6C, 0x6C, 0x6F}; // "Hello"
    ASSERT_TRUE(write_zip.writeFile("binary_test.bin", binary_data));
    
    write_zip.close();
    
    // 现在读取ZIP文件
    TinaXlsx::TXZipHandler read_zip;
    ASSERT_TRUE(read_zip.open("test.zip", TinaXlsx::TXZipHandler::OpenMode::Read));
    
    // 测试文件存在性检查
    EXPECT_TRUE(read_zip.hasFile("read_test.txt"));
    EXPECT_TRUE(read_zip.hasFile("binary_test.bin"));
    EXPECT_FALSE(read_zip.hasFile("nonexistent.txt"));
    
    // 读取文本文件
    std::string read_content = read_zip.readFileToString("read_test.txt");
    EXPECT_EQ(read_content, test_content);
    
    // 读取二进制文件
    std::vector<uint8_t> read_binary = read_zip.readFileToBytes("binary_test.bin");
    EXPECT_EQ(read_binary, binary_data);
    
    // 获取ZIP条目列表
    auto entries = read_zip.getEntries();
    EXPECT_EQ(entries.size(), 2);
    
    // 测试批量读取
    std::vector<std::string> filenames = {"read_test.txt", "binary_test.bin"};
    std::vector<std::pair<std::string, std::string>> results;
    
    std::size_t read_count = read_zip.readMultipleFiles(filenames, 
        [&results](const std::string& filename, const std::string& content) {
            results.emplace_back(filename, content);
        });
    
    EXPECT_EQ(read_count, 2);
    EXPECT_EQ(results.size(), 2);
    
    read_zip.close();
}

TEST_F(TXZipHandlerTest, ErrorHandling) {
    TinaXlsx::TXZipHandler zip;
    
    // 测试打开不存在的文件
    EXPECT_FALSE(zip.open("nonexistent.zip", TinaXlsx::TXZipHandler::OpenMode::Read));
    EXPECT_FALSE(zip.getLastError().empty());
    
    // 测试在未打开的情况下操作
    EXPECT_FALSE(zip.hasFile("test.txt"));
    EXPECT_TRUE(zip.readFileToString("test.txt").empty());
    
    // 测试在只读模式下写入
    ASSERT_TRUE(zip.open("test.zip", TinaXlsx::TXZipHandler::OpenMode::Write));
    ASSERT_TRUE(zip.writeFile("test.txt", "test"));
    zip.close();
    
    ASSERT_TRUE(zip.open("test.zip", TinaXlsx::TXZipHandler::OpenMode::Read));
    EXPECT_FALSE(zip.writeFile("new_file.txt", "content"));
    EXPECT_FALSE(zip.getLastError().empty());
    
    zip.close();
}

TEST_F(TXZipHandlerTest, MoveSemantics) {
    TinaXlsx::TXZipHandler zip1;
    ASSERT_TRUE(zip1.open("test.zip", TinaXlsx::TXZipHandler::OpenMode::Write));
    ASSERT_TRUE(zip1.writeFile("test.txt", "test content"));
    
    // 测试移动构造
    TinaXlsx::TXZipHandler zip2 = std::move(zip1);
    EXPECT_TRUE(zip2.isOpen());
    EXPECT_FALSE(zip1.isOpen()); // zip1 应该被移动清空
    
    // 测试移动赋值
    TinaXlsx::TXZipHandler zip3;
    zip3 = std::move(zip2);
    EXPECT_TRUE(zip3.isOpen());
    EXPECT_FALSE(zip2.isOpen()); // zip2 应该被移动清空
    
    zip3.close();
}

TEST_F(TXZipHandlerTest, CompressionLevels) {
    TinaXlsx::TXZipHandler zip;
    ASSERT_TRUE(zip.open("test.zip", TinaXlsx::TXZipHandler::OpenMode::Write));
    
    std::string large_content(10000, 'A'); // 创建一个较大的字符串用于测试压缩
    
    // 测试不同的压缩级别
    EXPECT_TRUE(zip.writeFile("no_compression.txt", large_content, 0));
    EXPECT_TRUE(zip.writeFile("fast_compression.txt", large_content, 1));
    EXPECT_TRUE(zip.writeFile("default_compression.txt", large_content, 6));
    EXPECT_TRUE(zip.writeFile("best_compression.txt", large_content, 9));
    
    zip.close();
    
    // 验证文件可以被正确读取
    ASSERT_TRUE(zip.open("test.zip", TinaXlsx::TXZipHandler::OpenMode::Read));
    
    EXPECT_EQ(zip.readFileToString("no_compression.txt"), large_content);
    EXPECT_EQ(zip.readFileToString("fast_compression.txt"), large_content);
    EXPECT_EQ(zip.readFileToString("default_compression.txt"), large_content);
    EXPECT_EQ(zip.readFileToString("best_compression.txt"), large_content);
    
    zip.close();
} 