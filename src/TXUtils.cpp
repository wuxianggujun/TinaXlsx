#include "TinaXlsx/TXUtils.hpp"
#include <algorithm>
#include <sstream>
#include <iomanip>
#include <ctime>
#include <locale>
#include <codecvt>

namespace TinaXlsx {

// ==================== TXUtils 实现 ====================

std::string TXUtils::trim(const std::string& str) {
    auto start = str.find_first_not_of(" \t\n\r\f\v");
    if (start == std::string::npos) {
        return "";
    }
    
    auto end = str.find_last_not_of(" \t\n\r\f\v");
    return str.substr(start, end - start + 1);
}

std::string TXUtils::toLower(const std::string& str) {
    std::string result = str;
    std::transform(result.begin(), result.end(), result.begin(), ::tolower);
    return result;
}

std::string TXUtils::toUpper(const std::string& str) {
    std::string result = str;
    std::transform(result.begin(), result.end(), result.begin(), ::toupper);
    return result;
}

std::vector<std::string> TXUtils::split(const std::string& str, char delimiter) {
    std::vector<std::string> tokens;
    std::string token;
    std::istringstream tokenStream(str);
    
    while (std::getline(tokenStream, token, delimiter)) {
        tokens.push_back(token);
    }
    
    return tokens;
}

std::string TXUtils::join(const std::vector<std::string>& strings, const std::string& delimiter) {
    if (strings.empty()) {
        return "";
    }
    
    std::ostringstream result;
    for (size_t i = 0; i < strings.size(); ++i) {
        if (i > 0) {
            result << delimiter;
        }
        result << strings[i];
    }
    
    return result.str();
}

bool TXUtils::startsWith(const std::string& str, const std::string& prefix) {
    return str.size() >= prefix.size() && 
           str.compare(0, prefix.size(), prefix) == 0;
}

bool TXUtils::endsWith(const std::string& str, const std::string& suffix) {
    return str.size() >= suffix.size() && 
           str.compare(str.size() - suffix.size(), suffix.size(), suffix) == 0;
}

std::string TXUtils::replace(const std::string& str, const std::string& from, const std::string& to) {
    std::string result = str;
    size_t start_pos = 0;
    
    while ((start_pos = result.find(from, start_pos)) != std::string::npos) {
        result.replace(start_pos, from.length(), to);
        start_pos += to.length();
    }
    
    return result;
}

bool TXUtils::isNumeric(const std::string& str) {
    if (str.empty()) return false;
    
    size_t start = 0;
    if (str[0] == '+' || str[0] == '-') {
        start = 1;
        if (str.length() == 1) return false;
    }
    
    bool hasDecimal = false;
    for (size_t i = start; i < str.length(); ++i) {
        if (str[i] == '.') {
            if (hasDecimal) return false;
            hasDecimal = true;
        } else if (!std::isdigit(str[i])) {
            return false;
        }
    }
    
    return true;
}

double TXUtils::stringToDouble(const std::string& str, double defaultValue) {
    try {
        return std::stod(str);
    } catch (...) {
        return defaultValue;
    }
}

int64_t TXUtils::stringToInt64(const std::string& str, int64_t defaultValue) {
    try {
        return std::stoll(str);
    } catch (...) {
        return defaultValue;
    }
}

std::string TXUtils::doubleToString(double value, int precision) {
    std::ostringstream oss;
    if (precision >= 0) {
        oss << std::fixed << std::setprecision(precision);
    }
    oss << value;
    return oss.str();
}

std::string TXUtils::getCurrentTimestamp() {
    auto now = std::time(nullptr);
    auto tm = *std::localtime(&now);
    
    std::ostringstream oss;
    oss << std::put_time(&tm, "%Y-%m-%d %H:%M:%S");
    return oss.str();
}

std::string TXUtils::escapeXml(const std::string& str) {
    std::string result;
    result.reserve(str.length() * 1.1); // 预留一些空间
    
    for (char c : str) {
        switch (c) {
            case '&':  result += "&amp;";  break;
            case '<':  result += "&lt;";   break;
            case '>':  result += "&gt;";   break;
            case '"':  result += "&quot;"; break;
            case '\'': result += "&apos;"; break;
            default:   result += c;        break;
        }
    }
    
    return result;
}

std::string TXUtils::unescapeXml(const std::string& str) {
    std::string result = str;
    
    result = replace(result, "&amp;", "&");
    result = replace(result, "&lt;", "<");
    result = replace(result, "&gt;", ">");
    result = replace(result, "&quot;", "\"");
    result = replace(result, "&apos;", "'");
    
    return result;
}

std::string TXUtils::generateUUID() {
    // 简单的UUID v4生成（伪随机）
    static std::random_device rd;
    static std::mt19937 gen(rd());
    static std::uniform_int_distribution<> dis(0, 15);
    static std::uniform_int_distribution<> dis2(8, 11);
    
    std::ostringstream oss;
    oss << std::hex;
    
    for (int i = 0; i < 8; i++) {
        oss << dis(gen);
    }
    oss << "-";
    for (int i = 0; i < 4; i++) {
        oss << dis(gen);
    }
    oss << "-4";
    for (int i = 0; i < 3; i++) {
        oss << dis(gen);
    }
    oss << "-" << dis2(gen);
    for (int i = 0; i < 3; i++) {
        oss << dis(gen);
    }
    oss << "-";
    for (int i = 0; i < 12; i++) {
        oss << dis(gen);
    }
    
    return oss.str();
}

size_t TXUtils::getFileSize(const std::string& filename) {
    std::ifstream file(filename, std::ios::binary | std::ios::ate);
    if (!file.is_open()) {
        return 0;
    }
    
    return static_cast<size_t>(file.tellg());
}

bool TXUtils::fileExists(const std::string& filename) {
    std::ifstream file(filename);
    return file.good();
}

std::string TXUtils::readTextFile(const std::string& filename) {
    std::ifstream file(filename);
    if (!file.is_open()) {
        return "";
    }
    
    std::ostringstream buffer;
    buffer << file.rdbuf();
    return buffer.str();
}

bool TXUtils::writeTextFile(const std::string& filename, const std::string& content) {
    std::ofstream file(filename);
    if (!file.is_open()) {
        return false;
    }
    
    file << content;
    return file.good();
}

std::string TXUtils::getFileExtension(const std::string& filename) {
    size_t dot_pos = filename.find_last_of('.');
    if (dot_pos == std::string::npos || dot_pos == filename.length() - 1) {
        return "";
    }
    return filename.substr(dot_pos + 1);
}

std::string TXUtils::getBaseName(const std::string& filename) {
    size_t dot_pos = filename.find_last_of('.');
    if (dot_pos == std::string::npos) {
        return filename;
    }
    return filename.substr(0, dot_pos);
}

std::string TXUtils::formatBytes(size_t bytes) {
    const char* units[] = {"B", "KB", "MB", "GB", "TB"};
    int unit = 0;
    double size = static_cast<double>(bytes);
    
    while (size >= 1024.0 && unit < 4) {
        size /= 1024.0;
        unit++;
    }
    
    std::ostringstream oss;
    oss << std::fixed << std::setprecision(2) << size << " " << units[unit];
    return oss.str();
}

} // namespace TinaXlsx 