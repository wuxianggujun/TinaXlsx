/**
 * @file TinaXlsx.hpp
 * @brief TinaXlsx 主要包含文件
 */

#pragma once

// 核心类型定义
#include "Types.hpp"

// 异常处理
#include "Exception.hpp"

// 核心功能类
#include "Reader.hpp"
#include "Writer.hpp"
#include "Format.hpp"

// 基础设施组件（用于高级定制）
#include "DataCache.hpp"
#include "ExcelStructureManager.hpp"
#include "WorksheetDataParser.hpp"
#include "XmlParser.hpp"
#include "ZipReader.hpp"

/**
 * @namespace TinaXlsx
 * @brief TinaXlsx 命名空间，包含所有库的功能
 */
namespace TinaXlsx {

/**
 * @brief 获取库版本信息
 * @return const char* 版本字符串
 */
inline const char* getVersion() {
    return "1.0.0";
}

/**
 * @brief 获取库的详细信息
 * @return const char* 详细信息字符串
 */
inline const char* getLibraryInfo() {
    return "TinaXlsx v1.0.0 - 高性能Excel读写库\n"
           "基于: libxlsxwriter (写入), minizip-ng (ZIP), expat (XML)\n"
           "特性: 快速读取, 流式支持, Unicode支持, 线程安全, 智能缓存\n"
           "GitHub: https://github.com/wuxianggujun/TinaXlsx";
}

} // namespace TinaXlsx 