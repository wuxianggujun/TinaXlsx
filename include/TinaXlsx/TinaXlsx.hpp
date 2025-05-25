#pragma once

/**
 * @file TinaXlsx.hpp
 * @brief TinaXlsx 主头文件
 * @author TinaXlsx Team
 * @version 2.0.0
 * @date 2024-12
 * 
 * 包含TinaXlsx库的所有核心功能
 */

// ==================== 核心类型系统 ====================
#include "TXTypes.hpp"         ///< 统一类型定义
#include "TXColor.hpp"         ///< 颜色处理类
#include "TXCoordinate.hpp"    ///< 坐标和范围类

// ==================== 样式系统 ====================
#include "TXStyle.hpp"         ///< 完整样式系统

// ==================== 核心业务类 ====================
#include "TXCell.hpp"          ///< 单元格类
#include "TXSheet.hpp"         ///< 工作表类
#include "TXWorkbook.hpp"      ///< 工作簿类

// ==================== 辅助工具类 ====================
#include "TXUtils.hpp"         ///< 工具函数
#include "TXZipHandler.hpp"    ///< ZIP处理
#include "TXXmlHandler.hpp"    ///< XML处理

/**
 * @namespace TinaXlsx
 * @brief TinaXlsx 库命名空间
 * 
 * 包含所有TinaXlsx相关的类和函数
 */
namespace TinaXlsx {

/**
 * @brief 库版本信息
 */
namespace Version {
    constexpr int MAJOR = 2;
    constexpr int MINOR = 0;
    constexpr int PATCH = 0;
    constexpr const char* STRING = "2.0.0";
    constexpr const char* BUILD_DATE = __DATE__;
}

/**
 * @brief 库特性标志
 */
namespace Features {
    constexpr bool HAS_STYLES = true;        ///< 是否支持样式
    constexpr bool HAS_COLORS = true;        ///< 是否支持颜色
    constexpr bool HAS_COORDINATES = true;   ///< 是否支持坐标系统
    constexpr bool HAS_ZIP_SUPPORT = true;   ///< 是否支持ZIP
    constexpr bool HAS_XML_SUPPORT = true;   ///< 是否支持XML
}

/**
 * @brief 快速使用别名
 */
using Workbook = TXWorkbook;
using Sheet = TXSheet;
using Cell = TXCell;
using Style = TXCellStyle;
using Color = TXColor;
using Coordinate = TXCoordinate;
using Range = TXRange;

// 常用类型别名
using RowIndex = TXTypes::RowIndex;
using ColIndex = TXTypes::ColIndex;
using ColorValue = TXTypes::ColorValue;
using FontSize = TXTypes::FontSize;

/**
 * @brief 初始化库
 * 
 * 在使用库的功能之前调用此函数进行初始化。
 * 这个函数是线程安全的，可以多次调用。
 * 
 * @return 成功返回true，失败返回false
 */
bool initialize();

/**
 * @brief 清理库资源
 * 
 * 在程序结束前调用此函数清理库使用的资源。
 * 这个函数是线程安全的，可以多次调用。
 */
void cleanup();

/**
 * @brief 获取库的版本信息
 * @return 版本字符串
 */
std::string getVersion();

/**
 * @brief 获取构建信息
 * @return 构建信息字符串
 */
std::string getBuildInfo();

} // namespace TinaXlsx 