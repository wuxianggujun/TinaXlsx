/**
 * @file TinaXlsx.hpp
 * @brief TinaXlsx Excel处理库 - 主头文件
 * @author TinaXlsx团队
 * @version 1.0.0
 * @date 2025
 * 
 * 基于xlsxio和xlsxwriter的高性能Excel处理库
 * 特性：
 * - C++17语法支持
 * - 零开销抽象
 * - RAII资源管理
 * - 类型安全
 * - 高性能读写
 */

#pragma once

// 核心类
#include "Reader.hpp"
#include "Writer.hpp"
#include "Workbook.hpp"
#include "Worksheet.hpp"
#include "Cell.hpp"
#include "Format.hpp"

// 工具类
#include "Exception.hpp"
#include "Types.hpp"
#include "Utils.hpp"
#include "DiffTool.hpp"

/**
 * @namespace TinaXlsx
 * @brief TinaXlsx库的命名空间
 */
namespace TinaXlsx {

/**
 * @brief 库版本信息
 */
struct Version {
    static constexpr int major = 1;
    static constexpr int minor = 0;
    static constexpr int patch = 0;
    static constexpr const char* string = "1.0.0";
};

/**
 * @brief 便捷创建读取器 - 基于minizip-ng和expat实现
 * @param filePath Excel文件路径
 * @return std::unique_ptr<Reader> 读取器智能指针
 */
[[nodiscard]] std::unique_ptr<Reader> createReader(const std::string& filePath);

/**
 * @brief 便捷创建写入器
 * @param filePath Excel文件路径
 * @return std::unique_ptr<Writer> 写入器智能指针
 */
[[nodiscard]] std::unique_ptr<Writer> createWriter(const std::string& filePath);

/**
 * @brief 便捷创建工作簿
 * @param filePath Excel文件路径
 * @return std::unique_ptr<Workbook> 工作簿智能指针
 */
[[nodiscard]] std::unique_ptr<Workbook> createWorkbook(const std::string& filePath);

} // namespace TinaXlsx 