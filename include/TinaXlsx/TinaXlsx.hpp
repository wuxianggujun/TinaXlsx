/**
 * @file TinaXlsx.hpp
 * @brief TinaXlsx Excel processing library - Main header file
 * @author TinaXlsx Team
 * @version 1.0.0
 * @date 2025
 * 
 * High performance Excel processing library based on xlsxio and xlsxwriter
 * Features:
 * - C++17 syntax support
 * - Zero overhead abstractions
 * - RAII resource management
 * - Type safety
 * - High performance read/write
 */

#pragma once

// Core classes
#include "Reader.hpp"
#include "Writer.hpp"
#include "Workbook.hpp"
#include "Worksheet.hpp"
#include "Cell.hpp"
#include "Format.hpp"

// Utility classes
#include "Exception.hpp"
#include "Types.hpp"
#include "Utils.hpp"
#include "DiffTool.hpp"

/**
 * @namespace TinaXlsx
 * @brief TinaXlsx library namespace
 */
namespace TinaXlsx {

/**
 * @brief Library version information
 */
struct Version {
    static constexpr int major = 1;
    static constexpr int minor = 0;
    static constexpr int patch = 0;
    static constexpr const char* string = "1.0.0";
};

/**
 * @brief Convenient reader factory - based on minizip-ng and expat
 * @param filePath Excel file path
 * @return std::unique_ptr<Reader> Reader smart pointer
 */
[[nodiscard]] std::unique_ptr<Reader> createReader(const std::string& filePath);

/**
 * @brief Convenient writer factory
 * @param filePath Excel file path
 * @return std::unique_ptr<Writer> Writer smart pointer
 */
[[nodiscard]] std::unique_ptr<Writer> createWriter(const std::string& filePath);

/**
 * @brief Convenient workbook factory
 * @param filePath Excel file path
 * @return std::unique_ptr<Workbook> Workbook smart pointer
 */
[[nodiscard]] std::unique_ptr<Workbook> createWorkbook(const std::string& filePath);

} // namespace TinaXlsx 