#pragma once

#include "TXTypes.hpp"
#include <string>
#include <vector>
#include <fstream>
#include <random>

namespace TinaXlsx {

/**
 * @brief 工具函数类
 * 
 * 提供通用的辅助工具函数，包括字符串处理、文件操作、数值转换等
 */
class TXUtils {
public:
    // ==================== 字符串操作 ====================
    
    /**
     * @brief 去除字符串两端的空白字符
     * @param str 输入字符串
     * @return 去除空白字符后的字符串
     */
    static std::string trim(const std::string& str);
    
    /**
     * @brief 转换为小写
     * @param str 输入字符串
     * @return 小写字符串
     */
    static std::string toLower(const std::string& str);
    
    /**
     * @brief 转换为大写
     * @param str 输入字符串
     * @return 大写字符串
     */
    static std::string toUpper(const std::string& str);
    
    /**
     * @brief 分割字符串
     * @param str 输入字符串
     * @param delimiter 分隔符
     * @return 分割后的字符串向量
     */
    static std::vector<std::string> split(const std::string& str, char delimiter);
    
    /**
     * @brief 连接字符串
     * @param strings 字符串向量
     * @param delimiter 分隔符
     * @return 连接后的字符串
     */
    static std::string join(const std::vector<std::string>& strings, const std::string& delimiter);
    
    /**
     * @brief 检查字符串是否以指定前缀开始
     * @param str 字符串
     * @param prefix 前缀
     * @return 匹配返回true，否则返回false
     */
    static bool startsWith(const std::string& str, const std::string& prefix);
    
    /**
     * @brief 检查字符串是否以指定后缀结束
     * @param str 字符串
     * @param suffix 后缀
     * @return 匹配返回true，否则返回false
     */
    static bool endsWith(const std::string& str, const std::string& suffix);
    
    /**
     * @brief 替换字符串中的子串
     * @param str 原字符串
     * @param from 被替换的子串
     * @param to 替换成的子串
     * @return 替换后的字符串
     */
    static std::string replace(const std::string& str, const std::string& from, const std::string& to);
    
    // ==================== 数值转换 ====================
    
    /**
     * @brief 检查字符串是否为数字
     * @param str 字符串
     * @return 是数字返回true，否则返回false
     */
    static bool isNumeric(const std::string& str);
    
    /**
     * @brief 字符串转换为double
     * @param str 字符串
     * @param defaultValue 转换失败时的默认值
     * @return 转换结果
     */
    static double stringToDouble(const std::string& str, double defaultValue = 0.0);
    
    /**
     * @brief 字符串转换为int64_t
     * @param str 字符串
     * @param defaultValue 转换失败时的默认值
     * @return 转换结果
     */
    static int64_t stringToInt64(const std::string& str, int64_t defaultValue = 0);
    
    /**
     * @brief double转换为字符串
     * @param value 数值
     * @param precision 精度，负数表示不限制
     * @return 字符串表示
     */
    static std::string doubleToString(double value, int precision = -1);
    
    // ==================== 时间工具 ====================
    
    /**
     * @brief 获取当前时间戳字符串
     * @return 时间戳字符串 (YYYY-MM-DD HH:MM:SS 格式)
     */
    static std::string getCurrentTimestamp();
    
    // ==================== XML 工具 ====================
    
    /**
     * @brief XML转义
     * @param str 输入字符串
     * @return 转义后的字符串
     */
    static std::string escapeXml(const std::string& str);
    
    /**
     * @brief XML反转义
     * @param str 转义的字符串
     * @return 反转义后的字符串
     */
    static std::string unescapeXml(const std::string& str);
    
    // ==================== UUID 生成 ====================
    
    /**
     * @brief 生成UUID
     * @return UUID字符串
     */
    static std::string generateUUID();
    
    // ==================== 文件操作 ====================
    
    /**
     * @brief 获取文件大小
     * @param filename 文件名
     * @return 文件大小（字节），失败返回0
     */
    static size_t getFileSize(const std::string& filename);
    
    /**
     * @brief 检查文件是否存在
     * @param filename 文件名
     * @return 存在返回true，否则返回false
     */
    static bool fileExists(const std::string& filename);
    
    /**
     * @brief 读取文本文件
     * @param filename 文件名
     * @return 文件内容，失败返回空字符串
     */
    static std::string readTextFile(const std::string& filename);
    
    /**
     * @brief 写入文本文件
     * @param filename 文件名
     * @param content 文件内容
     * @return 成功返回true，失败返回false
     */
    static bool writeTextFile(const std::string& filename, const std::string& content);
    
    /**
     * @brief 获取文件扩展名
     * @param filename 文件名
     * @return 扩展名（不含.），没有扩展名返回空字符串
     */
    static std::string getFileExtension(const std::string& filename);
    
    /**
     * @brief 获取文件基础名（不含扩展名）
     * @param filename 文件名
     * @return 基础名
     */
    static std::string getBaseName(const std::string& filename);
    
    // ==================== 格式化工具 ====================
    
    /**
     * @brief 格式化字节大小
     * @param bytes 字节数
     * @return 格式化的字符串 (如 "1.5 KB", "2.3 MB")
     */
    static std::string formatBytes(size_t bytes);
};

} // namespace TinaXlsx 