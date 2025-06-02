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
#include "TXCoordinate.hpp"    ///< 坐标类
#include "TXRange.hpp"         ///< 范围类

// ==================== 样式系统 ====================
#include "TXStyle.hpp"         ///< 完整样式系统

// ==================== 业务功能模块 ====================
#include "TXFormula.hpp"       ///< 公式处理类
#include "TXMergedCells.hpp"   ///< 合并单元格管理类
#include "TXNumberFormat.hpp"  ///< 数字格式化类
#include "TXStyleTemplate.hpp" ///< 样式模板系统（预设主题）

// ==================== 核心业务类 ====================
#include "TXCompactCell.hpp"   ///< 紧凑单元格类（内存优化版）
#include "TXSheet.hpp"         ///< 工作表类
#include "TXWorkbook.hpp"      ///< 工作簿类

// ==================== 工具类 ====================
#include "TXComponentManager.hpp" ///< 组件管理器

// XML处理器
#include "TXWorksheetXmlHandler.hpp"
#include "TXWorkbookXmlHandler.hpp"
#include "TXStylesXmlHandler.hpp"
#include "TXUnifiedXmlHandler.hpp"  // 统一的XML处理器

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
    constexpr bool HAS_STYLES = true;           ///< 是否支持样式
    constexpr bool HAS_COLORS = true;           ///< 是否支持颜色
    constexpr bool HAS_COORDINATES = true;      ///< 是否支持坐标系统
    constexpr bool HAS_ZIP_SUPPORT = true;      ///< 是否支持ZIP
    constexpr bool HAS_XML_SUPPORT = true;      ///< 是否支持XML
    constexpr bool HAS_FORMULAS = true;         ///< 是否支持公式
    constexpr bool HAS_MERGED_CELLS = true;     ///< 是否支持合并单元格
    constexpr bool HAS_NUMBER_FORMAT = true;    ///< 是否支持数字格式化
}

/**
 * @brief 快速使用别名
 */
using Workbook = TXWorkbook;
using Sheet = TXSheet;
using Cell = TXCompactCell;  // 使用内存优化的紧凑单元格
using Style = TXCellStyle;
using Color = TXColor;
using Coordinate = TXCoordinate;
using Range = TXRange;
using Formula = TXFormula;
using MergedCells = TXMergedCells;
using NumberFormat = TXNumberFormat;

// 常用类型别名
using RowIndex = row_t;
using ColIndex = column_t;
using ColorValue = color_value_t;
using FontSize = font_size_t;

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

/**
 * @brief 获取支持的特性列表
 * @return 特性描述字符串
 */
std::string getSupportedFeatures();

} // namespace TinaXlsx 
