#pragma once

#include <gtest/gtest.h>
#include <filesystem>
#include <string>
#include <memory>
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"

namespace TinaXlsx {

/**
 * @brief 测试文件生成器基类
 * 
 * 为每个测试类提供生成xlsx文件的功能，用于在WPS/Excel中验证兼容性
 */
class TestFileGenerator {
public:
    /**
     * @brief 构造函数
     * @param testSuiteName 测试套件名称，用作目录名
     */
    explicit TestFileGenerator(const std::string& testSuiteName);
    
    /**
     * @brief 析构函数
     */
    virtual ~TestFileGenerator() = default;

protected:
    /**
     * @brief 创建新的工作簿
     * @param filename 文件名（不包含扩展名）
     * @return 工作簿指针
     */
    std::unique_ptr<TXWorkbook> createWorkbook(const std::string& filename);
    
    /**
     * @brief 保存工作簿到文件
     * @param workbook 工作簿指针
     * @param filename 文件名（不包含扩展名）
     * @return 是否保存成功
     */
    bool saveWorkbook(const std::unique_ptr<TXWorkbook>& workbook, const std::string& filename);

    /**
     * @brief 保存工作簿到文件（移动语义版本）
     * @param workbook 工作簿指针（移动）
     * @param filename 文件名（不包含扩展名）
     * @return 是否保存成功
     */
    bool saveWorkbook(std::unique_ptr<TXWorkbook>&& workbook, const std::string& filename);
    
    /**
     * @brief 获取完整的文件路径
     * @param filename 文件名（不包含扩展名）
     * @return 完整路径
     */
    std::string getFilePath(const std::string& filename) const;
    
    /**
     * @brief 获取输出目录路径
     * @return 目录路径
     */
    std::string getOutputDir() const { return outputDir_; }
    
    /**
     * @brief 添加测试信息到工作表
     * @param sheet 工作表指针
     * @param testName 测试名称
     * @param description 测试描述
     */
    void addTestInfo(TXSheet* sheet, const std::string& testName, const std::string& description);

private:
    std::string testSuiteName_;  ///< 测试套件名称
    std::string outputDir_;      ///< 输出目录路径
    
    /**
     * @brief 创建输出目录
     */
    void createOutputDirectory();
};

/**
 * @brief 带文件生成功能的测试基类
 */
template<typename TestClass>
class TestWithFileGeneration : public ::testing::Test, public TestFileGenerator {
public:
    TestWithFileGeneration() : TestFileGenerator(getTestSuiteName()) {}

protected:
    void SetUp() override {
        // 子类可以重写此方法进行额外的设置
    }
    
    void TearDown() override {
        // 子类可以重写此方法进行清理
    }

private:
    /**
     * @brief 获取测试套件名称
     * @return 测试套件名称
     */
    std::string getTestSuiteName() const {
        // 从类型名中提取测试套件名称
        std::string typeName = typeid(TestClass).name();

        // 处理MSVC的"class ClassName"格式
        size_t classPos = typeName.find("class ");
        if (classPos == 0) {
            typeName = typeName.substr(6); // 去掉"class "前缀
        }

        // 处理其他编译器可能的前缀
        size_t structPos = typeName.find("struct ");
        if (structPos == 0) {
            typeName = typeName.substr(7); // 去掉"struct "前缀
        }

        // 查找"Test"并截取到Test结尾
        size_t testPos = typeName.find("Test");
        if (testPos != std::string::npos) {
            return typeName.substr(0, testPos + 4); // 包含"Test"
        }

        return typeName;
    }
};

} // namespace TinaXlsx
