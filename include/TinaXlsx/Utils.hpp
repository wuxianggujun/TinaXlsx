/**
 * @file Utils.hpp
 * @brief TinaXlsx工具函数
 */

#pragma once

#include "Types.hpp"
#include "Exception.hpp"
#include <string>
#include <vector>
#include <chrono>
#include <optional>

namespace TinaXlsx {
namespace Utils {

/**
 * @brief 日期时间工具类
 */
class DateTime {
public:
    /**
     * @brief 将C++时间点转换为Excel日期序列号
     * @param timePoint 时间点
     * @return double Excel日期序列号
     */
    [[nodiscard]] static double toExcelDate(const std::chrono::system_clock::time_point& timePoint);
    
    /**
     * @brief 将Excel日期序列号转换为C++时间点
     * @param excelDate Excel日期序列号
     * @return std::chrono::system_clock::time_point 时间点
     */
    [[nodiscard]] static std::chrono::system_clock::time_point fromExcelDate(double excelDate);
    
    /**
     * @brief 获取当前时间的Excel日期序列号
     * @return double Excel日期序列号
     */
    [[nodiscard]] static double now();
    
    /**
     * @brief 将日期字符串转换为Excel日期序列号
     * @param dateStr 日期字符串（支持多种格式）
     * @return std::optional<double> Excel日期序列号，解析失败则返回空
     */
    [[nodiscard]] static std::optional<double> parseDate(const std::string& dateStr);
    
    /**
     * @brief 将Excel日期序列号格式化为字符串
     * @param excelDate Excel日期序列号
     * @param format 格式字符串
     * @return std::string 格式化后的日期字符串
     */
    [[nodiscard]] static std::string formatDate(double excelDate, const std::string& format = "%Y-%m-%d");
};

/**
 * @brief 字符串工具类
 */
class String {
public:
    /**
     * @brief 去除字符串两端的空白字符
     * @param str 输入字符串
     * @return std::string 处理后的字符串
     */
    [[nodiscard]] static std::string trim(const std::string& str);
    
    /**
     * @brief 转换为小写
     * @param str 输入字符串
     * @return std::string 小写字符串
     */
    [[nodiscard]] static std::string toLower(const std::string& str);
    
    /**
     * @brief 转换为大写
     * @param str 输入字符串
     * @return std::string 大写字符串
     */
    [[nodiscard]] static std::string toUpper(const std::string& str);
    
    /**
     * @brief 分割字符串
     * @param str 输入字符串
     * @param delimiter 分隔符
     * @return std::vector<std::string> 分割后的字符串列表
     */
    [[nodiscard]] static std::vector<std::string> split(const std::string& str, const std::string& delimiter);
    
    /**
     * @brief 连接字符串
     * @param strings 字符串列表
     * @param delimiter 分隔符
     * @return std::string 连接后的字符串
     */
    [[nodiscard]] static std::string join(const std::vector<std::string>& strings, const std::string& delimiter);
    
    /**
     * @brief 检查字符串是否为数字
     * @param str 输入字符串
     * @return bool 是否为数字
     */
    [[nodiscard]] static bool isNumber(const std::string& str);
    
    /**
     * @brief 检查字符串是否为整数
     * @param str 输入字符串
     * @return bool 是否为整数
     */
    [[nodiscard]] static bool isInteger(const std::string& str);
    
    /**
     * @brief 替换字符串中的所有匹配项
     * @param str 输入字符串
     * @param from 要替换的子串
     * @param to 替换为的子串
     * @return std::string 替换后的字符串
     */
    [[nodiscard]] static std::string replace(const std::string& str, const std::string& from, const std::string& to);
};

/**
 * @brief 数据验证工具类
 */
class Validation {
public:
    /**
     * @brief 验证单元格位置是否有效
     * @param position 单元格位置
     * @return bool 是否有效
     */
    [[nodiscard]] static bool isValidPosition(const CellPosition& position);
    
    /**
     * @brief 验证单元格范围是否有效
     * @param range 单元格范围
     * @return bool 是否有效
     */
    [[nodiscard]] static bool isValidRange(const CellRange& range);
    
    /**
     * @brief 验证工作表名称是否有效
     * @param name 工作表名称
     * @return bool 是否有效
     */
    [[nodiscard]] static bool isValidSheetName(const std::string& name);
    
    /**
     * @brief 验证文件路径是否有效
     * @param filePath 文件路径
     * @return bool 是否有效
     */
    [[nodiscard]] static bool isValidFilePath(const std::string& filePath);
};

/**
 * @brief 数据转换工具类
 */
class Convert {
public:
    /**
     * @brief 尝试将字符串转换为CellValue
     * @param str 输入字符串
     * @param autoDetectType 是否自动检测类型
     * @return CellValue 转换后的值
     */
    [[nodiscard]] static CellValue stringToCellValue(const std::string& str, bool autoDetectType = true);
    
    /**
     * @brief 将CellValue转换为字符串
     * @param value 单元格值
     * @param format 格式字符串（可选）
     * @return std::string 转换后的字符串
     */
    [[nodiscard]] static std::string cellValueToString(const CellValue& value, const std::string& format = "");
    
    /**
     * @brief 将RowData转换为字符串数组
     * @param rowData 行数据
     * @return std::vector<std::string> 字符串数组
     */
    [[nodiscard]] static std::vector<std::string> rowDataToStrings(const RowData& rowData);
    
    /**
     * @brief 将字符串数组转换为RowData
     * @param strings 字符串数组
     * @param autoDetectType 是否自动检测类型
     * @return RowData 行数据
     */
    [[nodiscard]] static RowData stringsToRowData(const std::vector<std::string>& strings, bool autoDetectType = true);
    
    /**
     * @brief 将TableData转换为CSV格式字符串
     * @param tableData 表格数据
     * @param delimiter 分隔符
     * @return std::string CSV字符串
     */
    [[nodiscard]] static std::string tableToCsv(const TableData& tableData, const std::string& delimiter = ",");
    
    /**
     * @brief 从CSV格式字符串解析TableData
     * @param csvData CSV字符串
     * @param delimiter 分隔符
     * @param autoDetectType 是否自动检测类型
     * @return TableData 表格数据
     */
    [[nodiscard]] static TableData csvToTable(const std::string& csvData, const std::string& delimiter = ",", bool autoDetectType = true);
};

/**
 * @brief 性能分析工具类
 */
class Performance {
public:
    /**
     * @brief 简单的性能计时器
     */
    class Timer {
    private:
        std::chrono::high_resolution_clock::time_point start_;
        std::string name_;
        
    public:
        explicit Timer(const std::string& name = "Timer");
        ~Timer();
        
        /**
         * @brief 重置计时器
         */
        void reset();
        
        /**
         * @brief 获取已消耗的时间（毫秒）
         * @return double 消耗的时间
         */
        [[nodiscard]] double elapsed() const;
    };
    
    /**
     * @brief 估算内存使用量
     * @param tableData 表格数据
     * @return size_t 估算的内存使用量（字节）
     */
    [[nodiscard]] static size_t estimateMemoryUsage(const TableData& tableData);
    
    /**
     * @brief 获取推荐的批处理大小
     * @param totalRows 总行数
     * @param avgColumnsPerRow 平均每行列数
     * @param availableMemoryMB 可用内存（MB）
     * @return size_t 推荐的批处理大小
     */
    [[nodiscard]] static size_t getRecommendedBatchSize(size_t totalRows, size_t avgColumnsPerRow, size_t availableMemoryMB = 512);
};

/**
 * @brief 颜色工具类
 */
class ColorUtils {
public:
    /**
     * @brief 从RGB分量创建颜色
     * @param r 红色分量 (0-255)
     * @param g 绿色分量 (0-255) 
     * @param b 蓝色分量 (0-255)
     * @return Color 颜色值
     */
    [[nodiscard]] static constexpr Color rgb(UInt8 r, UInt8 g, UInt8 b) {
        return (static_cast<Color>(r) << 16) | (static_cast<Color>(g) << 8) | static_cast<Color>(b);
    }
    
    /**
     * @brief 从十六进制字符串创建颜色
     * @param hexColor 十六进制颜色字符串（如"#FF0000"或"FF0000"）
     * @return std::optional<Color> 颜色值，解析失败则返回空
     */
    [[nodiscard]] static std::optional<Color> fromHex(const std::string& hexColor);
    
    /**
     * @brief 将颜色转换为十六进制字符串
     * @param color 颜色值
     * @return std::string 十六进制颜色字符串
     */
    [[nodiscard]] static std::string toHex(Color color);
    
    /**
     * @brief 将颜色转换为RGB分量
     * @param color 颜色值
     * @return std::tuple<UInt8, UInt8, UInt8> RGB分量
     */
    [[nodiscard]] static std::tuple<UInt8, UInt8, UInt8> toRgb(Color color);
};

} // namespace Utils
} // namespace TinaXlsx 