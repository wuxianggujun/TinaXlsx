/**
 * @file TinaXlsx.cpp
 * @brief TinaXlsx Excel处理库 - 主实现文件
 */

#include "TinaXlsx/TinaXlsx.hpp"
#include "TinaXlsx/Reader.hpp"
#include "TinaXlsx/Writer.hpp"
#include "TinaXlsx/Workbook.hpp"

namespace TinaXlsx {

std::unique_ptr<Reader> createReader(const std::string& filePath) {
    return std::make_unique<Reader>(filePath);
}

std::unique_ptr<Writer> createWriter(const std::string& filePath) {
    return std::make_unique<Writer>(filePath);
}

std::unique_ptr<Workbook> createWorkbook(const std::string& filePath) {
    return std::make_unique<Workbook>(filePath);
}

} // namespace TinaXlsx 
