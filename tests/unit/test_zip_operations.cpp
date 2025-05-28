#include <gtest/gtest.h>
#include "TinaXlsx/TXZipArchive.hpp"
#include <fstream>
#include <filesystem>

class TXZipArchiveTest : public ::testing::Test {
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

TEST_F(TXZipArchiveTest, CreateAndWriteZip) {
    TinaXlsx::TXZipArchiveWriter zip;
    
    // 测试创建新的ZIP文件
    EXPECT_TRUE(zip.open("test.zip", false));
    EXPECT_TRUE(zip.isOpen());
    
    // 写入文件
    std::string content = "This is a test file content.";
    std::vector<uint8_t> content_data(content.begin(), content.end());
    EXPECT_TRUE(zip.write("test.txt", content_data));
    
    // 写入多个文件
    std::unordered_map<std::string, std::string> files;
    files["file1.txt"] = "Content of file 1";
    files["file2.txt"] = "Content of file 2";
    files["dir/file3.txt"] = "Content of file 3 in subdirectory";
    
    std::size_t written = 0;
    for (const auto& [filename, file_content] : files) {
        std::vector<uint8_t> file_data(file_content.begin(), file_content.end());
        if (zip.write(filename, file_data)) {
            written++;
        }
    }
    EXPECT_EQ(written, 3);
    
    zip.close();
    EXPECT_FALSE(zip.isOpen());
    
    // 验证文件是否存在
    EXPECT_TRUE(std::filesystem::exists("test.zip"));
}

TEST_F(TXZipArchiveTest, ReadZip) {
    // 首先创建一个ZIP文件
    TinaXlsx::TXZipArchiveWriter write_zip;
    ASSERT_TRUE(write_zip.open("test.zip", false));
    
    std::string test_content = "Test content for reading";
    std::vector<uint8_t> test_data(test_content.begin(), test_content.end());
    ASSERT_TRUE(write_zip.write("read_test.txt", test_data));
    
    std::vector<uint8_t> binary_data = {0x48, 0x65, 0x6C, 0x6C, 0x6F}; // "Hello"
    ASSERT_TRUE(write_zip.write("binary_test.bin", binary_data));
    
    write_zip.close();
    
    // 现在读取ZIP文件
    TinaXlsx::TXZipArchiveReader read_zip;
    ASSERT_TRUE(read_zip.open("test.zip"));
    
    // 测试文件存在性检查
    EXPECT_TRUE(read_zip.has("read_test.txt"));
    EXPECT_TRUE(read_zip.has("binary_test.bin"));
    EXPECT_FALSE(read_zip.has("nonexistent.txt"));
    
    // 读取文本文件
    std::string read_content = read_zip.readString("read_test.txt");
    EXPECT_EQ(read_content, test_content);
    
    // 读取二进制文件
    std::vector<uint8_t> read_binary = read_zip.read("binary_test.bin");
    EXPECT_EQ(read_binary, binary_data);
    
    // 获取ZIP条目列表
    auto entries = read_zip.entries();
    EXPECT_EQ(entries.size(), 2);
    
    // 测试批量读取（简化版本）
    std::vector<std::string> filenames = {"read_test.txt", "binary_test.bin"};
    std::vector<std::pair<std::string, std::string>> results;
    
    for (const auto& filename : filenames) {
        if (read_zip.has(filename)) {
            std::string content = read_zip.readString(filename);
            results.emplace_back(filename, content);
        }
    }
    
    EXPECT_EQ(results.size(), 2);
    
    read_zip.close();
}

TEST_F(TXZipArchiveTest, ErrorHandling) {
    TinaXlsx::TXZipArchiveReader zip;
    
    // 测试打开不存在的文件
    EXPECT_FALSE(zip.open("nonexistent.zip"));
    EXPECT_FALSE(zip.lastError().empty());
    
    // 测试在未打开的情况下操作
    EXPECT_FALSE(zip.has("test.txt"));
    EXPECT_TRUE(zip.readString("test.txt").empty());
    
    // 创建一个测试文件用于读取测试
    {
        TinaXlsx::TXZipArchiveWriter writer;
        ASSERT_TRUE(writer.open("test.zip", false));
        std::string test_content = "test";
        std::vector<uint8_t> test_data(test_content.begin(), test_content.end());
        ASSERT_TRUE(writer.write("test.txt", test_data));
        writer.close();
    }
    
    ASSERT_TRUE(zip.open("test.zip"));
    // 读取器不支持写入操作，这是预期的行为
    
    zip.close();
}

TEST_F(TXZipArchiveTest, MoveSemantics) {
    TinaXlsx::TXZipArchiveWriter zip1;
    ASSERT_TRUE(zip1.open("test.zip", false));
    std::string test_content = "test content";
    std::vector<uint8_t> test_data(test_content.begin(), test_content.end());
    ASSERT_TRUE(zip1.write("test.txt", test_data));
    
    // 测试移动构造
    TinaXlsx::TXZipArchiveWriter zip2 = std::move(zip1);
    EXPECT_TRUE(zip2.isOpen());
    EXPECT_FALSE(zip1.isOpen()); // zip1 应该被移动清空
    
    // 测试移动赋值
    TinaXlsx::TXZipArchiveWriter zip3;
    zip3 = std::move(zip2);
    EXPECT_TRUE(zip3.isOpen());
    EXPECT_FALSE(zip2.isOpen()); // zip2 应该被移动清空
    
    zip3.close();
}

TEST_F(TXZipArchiveTest, CompressionLevels) {
    TinaXlsx::TXZipArchiveWriter zip;
    ASSERT_TRUE(zip.open("test.zip", false));
    
    std::string large_content(10000, 'A'); // 创建一个较大的字符串用于测试压缩
    std::vector<uint8_t> large_data(large_content.begin(), large_content.end());
    
    // 测试写入文件（新API不直接暴露压缩级别设置）
    EXPECT_TRUE(zip.write("test_file.txt", large_data));
    
    zip.close();
    
    // 验证文件可以被正确读取
    TinaXlsx::TXZipArchiveReader reader;
    ASSERT_TRUE(reader.open("test.zip"));
    
    EXPECT_EQ(reader.readString("test_file.txt"), large_content);
    
    reader.close();
} 