/**
 * @file Utils.cpp
 * @brief TinaXlsx工具类实现
 */

#include "TinaXlsx/Utils.hpp"
#include <chrono>
#include <ctime>
#include <algorithm>
#include <cctype>
#include <sstream>
#include <iomanip>
#include <charconv>
#include <filesystem>
#include <regex>

namespace TinaXlsx {
namespace Utils {

// DateTime 实现
double DateTime::toExcelDate(const std::chrono::system_clock::time_point& timePoint) {
    // Excel的日期从1900年1月1日开始计算
    // 但Excel错误地认为1900年是闰年，所以我们需要调整
    auto epoch = std::chrono::system_clock::from_time_t(0);
    auto excelEpoch = std::chrono::system_clock::from_time_t(-2208988800); // 1900-01-01 00:00:00 UTC
    
    auto duration = timePoint - excelEpoch;
    auto days = std::chrono::duration_cast<std::chrono::hours>(duration).count() / 24.0;
    
    // 调整Excel的1900年闰年错误
    if (days >= 60) {
        days += 1; // 1900年2月29日不存在，但Excel认为存在
    }
    
    return days;
}

std::chrono::system_clock::time_point DateTime::fromExcelDate(double excelDate) {
    // 转换Excel日期回系统时间点
    auto excelEpoch = std::chrono::system_clock::from_time_t(-2208988800); // 1900-01-01 00:00:00 UTC
    
    // 调整Excel的1900年闰年错误
    if (excelDate >= 60) {
        excelDate -= 1;
    }
    
    auto duration = std::chrono::hours(static_cast<long long>(excelDate * 24));
    return excelEpoch + duration;
}

double DateTime::now() {
    return toExcelDate(std::chrono::system_clock::now());
}

std::optional<double> DateTime::parseDate(const std::string& dateStr) {
    // 简化的日期解析，支持常见格式
    std::regex dateRegex(R"((\d{4})[-/](\d{1,2})[-/](\d{1,2}))");
    std::smatch match;
    
    if (std::regex_match(dateStr, match, dateRegex)) {
        try {
            int year = std::stoi(match[1].str());
            int month = std::stoi(match[2].str());
            int day = std::stoi(match[3].str());
            
            std::tm tm = {};
            tm.tm_year = year - 1900;
            tm.tm_mon = month - 1;
            tm.tm_mday = day;
            
            auto timeT = std::mktime(&tm);
            if (timeT != -1) {
                auto timePoint = std::chrono::system_clock::from_time_t(timeT);
                return toExcelDate(timePoint);
            }
        } catch (const std::exception&) {
            // 解析失败
        }
    }
    
    return std::nullopt;
}

std::string DateTime::formatDate(double excelDate, const std::string& format) {
    auto timePoint = fromExcelDate(excelDate);
    auto timeT = std::chrono::system_clock::to_time_t(timePoint);
    
    std::stringstream ss;
    ss << std::put_time(std::localtime(&timeT), format.c_str());
    return ss.str();
}

// String 实现
std::string String::trim(const std::string& str) {
    auto start = str.find_first_not_of(" \t\n\r\f\v");
    if (start == std::string::npos) {
        return "";
    }
    
    auto end = str.find_last_not_of(" \t\n\r\f\v");
    return str.substr(start, end - start + 1);
}

std::string String::toLower(const std::string& str) {
    std::string result = str;
    std::transform(result.begin(), result.end(), result.begin(),
                   [](unsigned char c) { return std::tolower(c); });
    return result;
}

std::string String::toUpper(const std::string& str) {
    std::string result = str;
    std::transform(result.begin(), result.end(), result.begin(),
                   [](unsigned char c) { return std::toupper(c); });
    return result;
}

std::vector<std::string> String::split(const std::string& str, const std::string& delimiter) {
    std::vector<std::string> result;
    if (delimiter.empty()) {
        result.push_back(str);
        return result;
    }
    
    size_t start = 0;
    size_t end = str.find(delimiter);
    
    while (end != std::string::npos) {
        result.push_back(str.substr(start, end - start));
        start = end + delimiter.length();
        end = str.find(delimiter, start);
    }
    
    result.push_back(str.substr(start));
    return result;
}

std::string String::join(const std::vector<std::string>& strings, const std::string& delimiter) {
    if (strings.empty()) {
        return "";
    }
    
    std::stringstream ss;
    for (size_t i = 0; i < strings.size(); ++i) {
        if (i > 0) {
            ss << delimiter;
        }
        ss << strings[i];
    }
    
    return ss.str();
}

bool String::isNumber(const std::string& str) {
    if (str.empty()) {
        return false;
    }
    
    double value;
    auto result = std::from_chars(str.data(), str.data() + str.size(), value);
    return result.ec == std::errc{} && result.ptr == str.data() + str.size();
}

bool String::isInteger(const std::string& str) {
    if (str.empty()) {
        return false;
    }
    
    long long value;
    auto result = std::from_chars(str.data(), str.data() + str.size(), value);
    return result.ec == std::errc{} && result.ptr == str.data() + str.size();
}

std::string String::replace(const std::string& str, const std::string& from, const std::string& to) {
    if (from.empty()) {
        return str;
    }
    
    std::string result = str;
    size_t pos = 0;
    
    while ((pos = result.find(from, pos)) != std::string::npos) {
        result.replace(pos, from.length(), to);
        pos += to.length();
    }
    
    return result;
}

// Validation 实现
bool Validation::isValidPosition(const CellPosition& position) {
    // Excel的最大行数和列数
    constexpr uint32_t MAX_ROWS = 1048576;
    constexpr uint32_t MAX_COLUMNS = 16384;
    
    return position.row < MAX_ROWS && position.column < MAX_COLUMNS;
}

bool Validation::isValidRange(const CellRange& range) {
    return range.isValid() && 
           isValidPosition(range.start) && 
           isValidPosition(range.end);
}

bool Validation::isValidSheetName(const std::string& name) {
    if (name.empty() || name.length() > 31) {
        return false;
    }
    
    // Excel工作表名称不能包含这些字符: [ ] : \ / ? *
    const std::string invalidChars = "[]:\\/?*";
    return name.find_first_of(invalidChars) == std::string::npos;
}

bool Validation::isValidFilePath(const std::string& filePath) {
    try {
        std::filesystem::path path(filePath);
        // 检查路径是否有效
        return !path.empty() && path.is_absolute() || path.is_relative();
    } catch (const std::exception&) {
        return false;
    }
}

// Convert 实现
CellValue Convert::stringToCellValue(const std::string& str, bool autoDetectType) {
    if (str.empty()) {
        return std::monostate{};
    }
    
    if (!autoDetectType) {
        return str;
    }
    
    // 尝试解析为数字
    if (String::isInteger(str)) {
        try {
            return std::stoll(str);
        } catch (const std::exception&) {
            // 如果整数解析失败，尝试浮点数
        }
    }
    
    if (String::isNumber(str)) {
        try {
            return std::stod(str);
        } catch (const std::exception&) {
            // 如果数字解析失败，返回字符串
        }
    }
    
    // 尝试解析为布尔值
    std::string lower = String::toLower(str);
    if (lower == "true" || lower == "yes" || lower == "1") {
        return true;
    } else if (lower == "false" || lower == "no" || lower == "0") {
        return false;
    }
    
    // 默认返回字符串
    return str;
}

std::string Convert::cellValueToString(const CellValue& value, const std::string& format) {
    if (format.empty()) {
        // 使用基本的类型安全枚举转换
        const CellValueType type = getCellValueType(value);
        
        switch (type) {
            case CellValueType::String:
                return std::get<std::string>(value);
            case CellValueType::Double: {
                const double val = std::get<double>(value);
                // 优化：检查是否为整数以避免不必要的小数点
                if (val == static_cast<long long>(val) && 
                    val >= -9007199254740992.0 && val <= 9007199254740992.0) {
                    return std::to_string(static_cast<long long>(val));
                }
                return std::to_string(val);
            }
            case CellValueType::Integer:
                return std::to_string(std::get<long long>(value));
            case CellValueType::Boolean:
                return std::get<bool>(value) ? "true" : "false";
            case CellValueType::Empty:
                return "";
            default:
                return "";
        }
    } else {
        // 带格式化的转换（保留原有逻辑用于格式化）
        return std::visit([&format](const auto& v) -> std::string {
            using T = std::decay_t<decltype(v)>;
            if constexpr (std::is_same_v<T, std::monostate>) {
                return "";
            } else if constexpr (std::is_same_v<T, std::string>) {
                return v;
            } else if constexpr (std::is_same_v<T, double>) {
                // 简化的格式化，实际实现可能需要更复杂的逻辑
                std::stringstream ss;
                ss << std::fixed << std::setprecision(2) << v;
                return ss.str();
            } else if constexpr (std::is_same_v<T, long long>) {
                return std::to_string(v);
            } else if constexpr (std::is_same_v<T, bool>) {
                return v ? "true" : "false";
            } else {
                return "";
            }
        }, value);
    }
}

std::vector<std::string> Convert::rowDataToStrings(const RowData& rowData) {
    std::vector<std::string> result;
    result.reserve(rowData.size());
    
    for (const auto& cellValue : rowData) {
        result.push_back(cellValueToString(cellValue));
    }
    
    return result;
}

RowData Convert::stringsToRowData(const std::vector<std::string>& strings, bool autoDetectType) {
    RowData result;
    result.reserve(strings.size());
    
    for (const auto& str : strings) {
        result.push_back(stringToCellValue(str, autoDetectType));
    }
    
    return result;
}

std::string Convert::tableToCsv(const TableData& tableData, const std::string& delimiter) {
    std::stringstream ss;
    
    for (size_t row = 0; row < tableData.size(); ++row) {
        if (row > 0) {
            ss << "\n";
        }
        
        for (size_t col = 0; col < tableData[row].size(); ++col) {
            if (col > 0) {
                ss << delimiter;
            }
            
            std::string cellStr = cellValueToString(tableData[row][col]);
            // 简单的CSV转义：如果包含分隔符、换行符或引号，则用引号包围
            if (cellStr.find(delimiter) != std::string::npos || 
                cellStr.find('\n') != std::string::npos || 
                cellStr.find('"') != std::string::npos) {
                // 转义引号
                cellStr = String::replace(cellStr, "\"", "\"\"");
                cellStr = "\"" + cellStr + "\"";
            }
            
            ss << cellStr;
        }
    }
    
    return ss.str();
}

TableData Convert::csvToTable(const std::string& csvData, const std::string& delimiter, bool autoDetectType) {
    TableData result;
    auto lines = String::split(csvData, "\n");
    
    for (const auto& line : lines) {
        if (line.empty()) {
            continue;
        }
        
        auto cells = String::split(line, delimiter);
        RowData rowData;
        
        for (const auto& cell : cells) {
            std::string cleanCell = String::trim(cell);
            
            // 简单的CSV解析：去除引号
            if (cleanCell.length() >= 2 && cleanCell.front() == '"' && cleanCell.back() == '"') {
                cleanCell = cleanCell.substr(1, cleanCell.length() - 2);
                cleanCell = String::replace(cleanCell, "\"\"", "\"");
            }
            
            rowData.push_back(stringToCellValue(cleanCell, autoDetectType));
        }
        
        result.push_back(std::move(rowData));
    }
    
    return result;
}

// Performance 实现
Performance::Timer::Timer(const std::string& name) : name_(name) {
    reset();
}

void Performance::Timer::reset() {
    start_ = std::chrono::high_resolution_clock::now();
}

double Performance::Timer::elapsed() const {
    auto end = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start_);
    return duration.count() / 1000.0; // 返回毫秒
}

size_t Performance::estimateMemoryUsage(const TableData& tableData) {
    size_t totalSize = 0;
    
    for (const auto& row : tableData) {
        for (const auto& cell : row) {
            totalSize += std::visit([](const auto& v) -> size_t {
                using T = std::decay_t<decltype(v)>;
                if constexpr (std::is_same_v<T, std::monostate>) {
                    return 1; // 最小开销
                } else if constexpr (std::is_same_v<T, std::string>) {
                    return v.size() + sizeof(std::string);
                } else if constexpr (std::is_same_v<T, double>) {
                    return sizeof(double);
                } else if constexpr (std::is_same_v<T, long long>) {
                    return sizeof(long long);
                } else if constexpr (std::is_same_v<T, bool>) {
                    return sizeof(bool);
                } else {
                    return sizeof(CellValue);
                }
            }, cell);
        }
        totalSize += sizeof(RowData); // 行的开销
    }
    
    totalSize += sizeof(TableData); // 表的开销
    return totalSize;
}

size_t Performance::getRecommendedBatchSize(size_t totalRows, size_t avgColumnsPerRow, size_t availableMemoryMB) {
    // 估算每行的内存使用量（字节）
    size_t estimatedBytesPerRow = avgColumnsPerRow * (sizeof(CellValue) + 20); // 假设平均每个单元格20字节
    
    // 转换MB到字节
    size_t availableMemoryBytes = availableMemoryMB * 1024 * 1024;
    
    // 计算推荐的批处理大小
    size_t recommendedBatchSize = availableMemoryBytes / estimatedBytesPerRow;
    
    // 设置合理的上下限
    recommendedBatchSize = std::max<size_t>(100, recommendedBatchSize);    // 最小100行
    recommendedBatchSize = std::min<size_t>(10000, recommendedBatchSize);  // 最大10000行
    recommendedBatchSize = std::min(totalRows, recommendedBatchSize);      // 不超过总行数
    
    return recommendedBatchSize;
}

// ColorUtils 实现
std::optional<Color> ColorUtils::fromHex(const std::string& hexColor) {
    std::string hex = hexColor;
    
    // 移除前缀 '#' 如果存在
    if (!hex.empty() && hex[0] == '#') {
        hex = hex.substr(1);
    }
    
    // 检查长度
    if (hex.length() != 6) {
        return std::nullopt;
    }
    
    // 检查是否都是有效的十六进制字符
    if (hex.find_first_not_of("0123456789ABCDEFabcdef") != std::string::npos) {
        return std::nullopt;
    }
    
    try {
        return static_cast<Color>(std::stoul(hex, nullptr, 16));
    } catch (const std::exception&) {
        return std::nullopt;
    }
}

std::string ColorUtils::toHex(Color color) {
    std::stringstream ss;
    ss << "#" << std::hex << std::uppercase << std::setfill('0') << std::setw(6) << (color & 0xFFFFFF);
    return ss.str();
}

std::tuple<unsigned char, unsigned char, unsigned char> ColorUtils::toRgb(Color color) {
    unsigned char r = (color >> 16) & 0xFF;
    unsigned char g = (color >> 8) & 0xFF;
    unsigned char b = color & 0xFF;
    return std::make_tuple(r, g, b);
}

} // namespace Utils
} // namespace TinaXlsx 