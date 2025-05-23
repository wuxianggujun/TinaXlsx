/**
 * @file Exception.hpp
 * @brief TinaXlsx异常处理
 */

#pragma once

#include <stdexcept>
#include <string>
#include <string_view>

namespace TinaXlsx {

/**
 * @brief TinaXlsx基础异常类
 */
class Exception : public std::runtime_error {
public:
    explicit Exception(const std::string& message) 
        : std::runtime_error(message) {}
    
    explicit Exception(std::string_view message) 
        : std::runtime_error(std::string(message)) {}
    
    explicit Exception(const char* message) 
        : std::runtime_error(message) {}
};

/**
 * @brief 文件操作异常
 */
class FileException : public Exception {
public:
    explicit FileException(const std::string& message) 
        : Exception("文件操作错误: " + message) {}
};

/**
 * @brief 工作表操作异常
 */
class WorksheetException : public Exception {
public:
    explicit WorksheetException(const std::string& message) 
        : Exception("工作表操作错误: " + message) {}
};

/**
 * @brief 单元格操作异常
 */
class CellException : public Exception {
public:
    explicit CellException(const std::string& message) 
        : Exception("单元格操作错误: " + message) {}
};

/**
 * @brief 格式操作异常
 */
class FormatException : public Exception {
public:
    explicit FormatException(const std::string& message) 
        : Exception("格式操作错误: " + message) {}
};

/**
 * @brief 数据类型异常
 */
class TypeException : public Exception {
public:
    explicit TypeException(const std::string& message) 
        : Exception("数据类型错误: " + message) {}
};

/**
 * @brief 索引越界异常
 */
class IndexOutOfRangeException : public Exception {
public:
    explicit IndexOutOfRangeException(const std::string& message) 
        : Exception("索引越界: " + message) {}
};

/**
 * @brief 无效参数异常
 */
class InvalidArgumentException : public Exception {
public:
    explicit InvalidArgumentException(const std::string& message) 
        : Exception("无效参数: " + message) {}
};

} // namespace TinaXlsx 