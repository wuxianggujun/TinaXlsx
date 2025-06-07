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
#include "TXError.hpp"         ///< 错误处理
#include "TXResult.hpp"        ///< 结果包装
#include "TXVariant.hpp"       ///< 通用数据类型
#include "TXColor.hpp"         ///< 颜色处理类
#include "TXCoordinate.hpp"    ///< 坐标类

// ==================== 内存优先架构 (新) ====================
#include "TXInMemorySheet.hpp"      ///< 内存优先工作表
#include "TXBatchSIMDProcessor.hpp" ///< SIMD批量处理器
#include "TXZeroCopySerializer.hpp" ///< 零拷贝序列化器
#include "TXXMLTemplates.hpp"       ///< XML模板系统

// ==================== 内存管理 ====================
#include "TXSlabAllocator.hpp"      ///< Slab分配器
#include "TXUnifiedMemoryManager.hpp" ///< 统一内存管理
#include "TXGlobalStringPool.hpp"   ///< 全局字符串池

// ==================== 样式系统 ====================
#include "TXStyle.hpp"         ///< 完整样式系统

// ==================== 业务功能模块 ====================
#include "TXFormula.hpp"       ///< 公式处理类
#include "TXMergedCells.hpp"   ///< 合并单元格管理类
#include "TXNumberFormat.hpp"  ///< 数字格式化类
#include "TXStyleTemplate.hpp" ///< 样式模板系统（预设主题）


// ==================== 工具类 ====================
#include "TXComponentManager.hpp" ///< 组件管理器
#include "TXRange.hpp"         ///< 范围类

// 注意：传统XML处理器已被TXZeroCopySerializer替代

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

// 🚀 不使用别名！TX前缀已经足够防止命名冲突
// 直接使用完整的类名：TXWorkbook, TXSheet, TXCell, TXCoordinate 等

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

// ==================== 内存优先架构 快速API ====================

/**
 * @brief 内存优先架构配置
 */
struct MemoryFirstConfig {
    bool enable_simd = true;             ///< 启用SIMD优化
    bool enable_parallel = true;         ///< 启用并行处理  
    size_t batch_size = 10000;          ///< 批处理大小
    size_t memory_limit_gb = 4;         ///< 内存限制(GB)
    bool enable_compression = true;      ///< 启用压缩
    size_t parallel_threshold = 1000;   ///< 并行处理阈值
};

/**
 * @brief 快速创建Excel文件 - 主要API (目标: <2ms)
 */
class QuickExcel {
public:
    /**
     * @brief 快速创建数值表格
     * @param data 二维数值数组
     * @param filename 输出文件名
     * @return 操作结果
     */
    static TXResult<void> createFromNumbers(
        const std::vector<std::vector<double>>& data,
        const std::string& filename
    );
    
    /**
     * @brief 快速创建混合数据表格  
     * @param data 二维混合数据数组
     * @param filename 输出文件名
     * @return 操作结果
     */
    static TXResult<void> createFromData(
        const std::vector<std::vector<TXVariant>>& data,
        const std::string& filename
    );
    
    /**
     * @brief 从CSV快速创建Excel
     * @param csv_content CSV内容
     * @param filename 输出文件名
     * @return 操作结果
     */
    static TXResult<void> createFromCSV(
        const std::string& csv_content,
        const std::string& filename
    );
    
    // 🚀 直接使用完整类名，无需别名
};

// 🚀 无别名！直接使用完整类名：
// - TXInMemoryWorkbook
// - TXInMemorySheet
// - TXBatchSIMDProcessor
// - TXZeroCopySerializer

} // namespace TinaXlsx 
