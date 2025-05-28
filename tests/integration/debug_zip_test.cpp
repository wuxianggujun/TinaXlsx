//
// ZIP Handler Debug Test
// 专门用于诊断ZIP写入问题
//

#include <gtest/gtest.h>
#include "TinaXlsx/TXZipArchive.hpp"
#include <iostream>
#include <filesystem>

class DebugZipTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 清理可能存在的测试文件
        std::filesystem::remove("debug_test.zip");
    }

    void TearDown() override {
        // 清理测试文件
        std::filesystem::remove("debug_test.zip");
    }
};

TEST_F(DebugZipTest, DiagnoseWriteFailure) {
    TinaXlsx::TXZipArchiveWriter zip;
    
    std::cout << "=== ZIP Archive Debug Test ===" << std::endl;
    
    // 步骤1: 打开ZIP文件
    std::cout << "1. Opening ZIP file for writing..." << std::endl;
    bool opened = zip.open("debug_test.zip", false);
    std::cout << "   Open result: " << (opened ? "SUCCESS" : "FAILED") << std::endl;
    
    if (!opened) {
        std::cout << "   Error: " << zip.lastError() << std::endl;
        FAIL() << "Failed to open ZIP file: " << zip.lastError();
    }
    
    // 步骤2: 检查是否真的打开了
    std::cout << "2. Checking if file is open..." << std::endl;
    bool isOpen = zip.isOpen();
    std::cout << "   Is open: " << (isOpen ? "YES" : "NO") << std::endl;
    EXPECT_TRUE(isOpen);
    
    // 步骤3: 尝试写入文件
    std::cout << "3. Writing test file..." << std::endl;
    std::string content = "Hello, World!";
    std::cout << "   Content: \"" << content << "\"" << std::endl;
    std::cout << "   Content size: " << content.size() << " bytes" << std::endl;
    
    std::vector<uint8_t> content_data(content.begin(), content.end());
    bool written = zip.write("test.txt", content_data);
    std::cout << "   Write result: " << (written ? "SUCCESS" : "FAILED") << std::endl;
    
    if (!written) {
        std::cout << "   Error: " << zip.lastError() << std::endl;
        // 不要立即失败，继续收集信息
    }
    
    // 步骤4: 关闭文件
    std::cout << "4. Closing ZIP file..." << std::endl;
    zip.close();
    std::cout << "   Closed successfully" << std::endl;
    
    // 步骤5: 检查文件是否存在
    std::cout << "5. Checking if ZIP file exists..." << std::endl;
    bool fileExists = std::filesystem::exists("debug_test.zip");
    std::cout << "   File exists: " << (fileExists ? "YES" : "NO") << std::endl;
    
    if (fileExists) {
        auto fileSize = std::filesystem::file_size("debug_test.zip");
        std::cout << "   File size: " << fileSize << " bytes" << std::endl;
    }
    
    std::cout << "=== End Debug Test ===" << std::endl;
    
    // 只有在写入失败时才标记测试失败
    if (!written) {
        FAIL() << "ZIP write operation failed: " << zip.lastError();
    }
}

TEST_F(DebugZipTest, BasicFunctionality) {
    TinaXlsx::TXZipArchiveWriter zip;
    
    // 简单的功能测试
    EXPECT_TRUE(zip.open("debug_test.zip", false));
    EXPECT_TRUE(zip.isOpen());
    
    std::string content = "Test content";
    std::vector<uint8_t> content_data(content.begin(), content.end());
    bool writeResult = zip.write("simple.txt", content_data);
    
    if (!writeResult) {
        std::cout << "Write failed with error: " << zip.lastError() << std::endl;
    }
    
    zip.close();
    EXPECT_FALSE(zip.isOpen());
    
    // 如果写入失败，输出详细信息但不一定失败测试
    // 这样我们可以看到具体的错误信息
    EXPECT_TRUE(writeResult) << "Write failed: " << zip.lastError();
} 