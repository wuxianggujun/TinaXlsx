#pragma once

// 测试基础头文件
// 这个文件提供了所有测试需要的基础功能

#include <gtest/gtest.h>
#include <filesystem>
#include <string>
#include <memory>
#include <iostream>

// TinaXlsx 核心头文件
#include "TinaXlsx/TinaXlsx.hpp"
#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"

// 包含现有的文件生成器
#include "test_file_generator.hpp"

// 为了向后兼容，在全局命名空间中也提供这个类
using TinaXlsx::TestWithFileGeneration;
using TinaXlsx::TestFileGenerator;

/**
 * @brief 测试辅助函数
 */
namespace TestUtils {

/**
 * @brief 保存工作簿的便捷函数
 * @param workbook 工作簿对象
 * @param filename 文件名（不包含扩展名）
 * @param outputDir 输出目录
 * @return 是否保存成功
 */
inline bool saveWorkbook(TinaXlsx::TXWorkbook& workbook, const std::string& filename, const std::string& outputDir = "test_output") {
    // 创建输出目录
    std::filesystem::path dir = std::filesystem::current_path() / outputDir;
    std::filesystem::create_directories(dir);

    // 构建完整路径
    std::string fullPath = (dir / (filename + ".xlsx")).string();

    bool result = workbook.saveToFile(fullPath);

    if (result) {
        std::cout << "📁 文件已保存: " << fullPath << std::endl;
    } else {
        std::cout << "❌ 文件保存失败: " << fullPath << std::endl;
    }

    return result;
}

/**
 * @brief 创建工作簿的便捷函数
 * @param filename 文件名（用于日志）
 * @return 工作簿对象
 */
inline TinaXlsx::TXWorkbook createWorkbook(const std::string& filename = "") {
    TinaXlsx::TXWorkbook workbook;
    if (!filename.empty()) {
        std::cout << "📝 创建工作簿: " << filename << std::endl;
    }
    return workbook;
}

} // namespace TestUtils
