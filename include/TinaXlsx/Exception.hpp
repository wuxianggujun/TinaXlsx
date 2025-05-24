/**
 * @file Exception.hpp
 * @brief TinaXlsx exception handling
 */

#pragma once

#include <stdexcept>
#include <string>
#include <string_view>

namespace TinaXlsx {

/**
 * @brief TinaXlsx base exception class
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
 * @brief File operation exception
 */
class FileException : public Exception {
public:
    explicit FileException(const std::string& message) 
        : Exception("File operation error: " + message) {}
};

/**
 * @brief Worksheet operation exception
 */
class WorksheetException : public Exception {
public:
    explicit WorksheetException(const std::string& message) 
        : Exception("Worksheet operation error: " + message) {}
};

/**
 * @brief Cell operation exception
 */
class CellException : public Exception {
public:
    explicit CellException(const std::string& message) 
        : Exception("Cell operation error: " + message) {}
};

/**
 * @brief Format operation exception
 */
class FormatException : public Exception {
public:
    explicit FormatException(const std::string& message) 
        : Exception("Format operation error: " + message) {}
};

/**
 * @brief Data type exception
 */
class TypeException : public Exception {
public:
    explicit TypeException(const std::string& message) 
        : Exception("Data type error: " + message) {}
};

/**
 * @brief Index out of range exception
 */
class IndexOutOfRangeException : public Exception {
public:
    explicit IndexOutOfRangeException(const std::string& message) 
        : Exception("Index out of range: " + message) {}
};

/**
 * @brief Invalid argument exception
 */
class InvalidArgumentException : public Exception {
public:
    explicit InvalidArgumentException(const std::string& message) 
        : Exception("Invalid argument: " + message) {}
};

} // namespace TinaXlsx 