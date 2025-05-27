#include "TinaXlsx/TXTypes.hpp"
#include <algorithm>
#include <cctype>
#include <vector>

namespace TinaXlsx {

// ==================== column_t 静态方法实现 ====================

column_t::index_t column_t::column_index_from_string(const std::string& column_string) {
    if (column_string.empty()) {
        return INVALID_COLUMN;
    }
    
    index_t result = 0;
    for (char c : column_string) {
        if (c < 'A' || c > 'Z') {
            return INVALID_COLUMN;
        }
        result = result * 26 + (c - 'A' + 1);
    }
    
    return result;
}

std::string column_t::column_string_from_index(index_t column_index) {
    if (column_index == 0 || column_index > MAX_COLUMNS) {
        return "";
    }
    
    std::string result;
    index_t index = column_index;
    
    while (index > 0) {
        index--; // 转换为 0-based
        result = char('A' + (index % 26)) + result;
        index /= 26;
    }
    
    return result;
}

// ==================== 工具函数实现 ====================

bool is_valid_sheet_name(const std::string& name) {
    // 检查长度
    if (name.empty() || name.length() > MAX_SHEET_NAME) {
        return false;
    }
    
    // 检查是否包含非法字符
    const std::string illegal_chars = "[]*/\\?:";
    for (char c : illegal_chars) {
        if (name.find(c) != std::string::npos) {
            return false;
        }
    }
    
    // 检查是否以单引号开头或结尾
    if (name.front() == '\'' || name.back() == '\'') {
        return false;
    }
    
    // 检查是否是保留名称
    std::string lower_name = name;
    std::transform(lower_name.begin(), lower_name.end(), lower_name.begin(), 
                   [](unsigned char c) { return std::tolower(c); });
    
    const std::vector<std::string> reserved_names = {
        "history"
    };
    
    for (const auto& reserved : reserved_names) {
        if (lower_name == reserved) {
            return false;
        }
    }
    
    return true;
}

std::string get_file_extension(const std::string& filename) {
    auto pos = filename.find_last_of('.');
    if (pos == std::string::npos) {
        return "";
    }
    return filename.substr(pos);
}

bool is_excel_file(const std::string& filename) {
    std::string ext = get_file_extension(filename);
    std::transform(ext.begin(), ext.end(), ext.begin(), 
                   [](unsigned char c) { return std::tolower(c); });
    
    return ext == XLSX_EXTENSION || ext == XLS_EXTENSION;
}

} // namespace TinaXlsx 
