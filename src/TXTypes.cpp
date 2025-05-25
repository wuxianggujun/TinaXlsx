#include "TinaXlsx/TXTypes.hpp"
#include <algorithm>
#include <regex>

namespace TinaXlsx {
namespace TXTypes {

// TXTypes 现在主要包含常量和基础验证函数
// 坐标转换功能已移至 TXCoordinate 类中
// 颜色操作功能已移至 TXColor 类中

// 这个文件主要用于存放一些静态初始化或复杂的工具函数
// 目前所有的 TXTypes 方法都是内联的，所以这个文件可以保持简洁

bool isValidSheetName(const std::string& name) {
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
    std::transform(lower_name.begin(), lower_name.end(), lower_name.begin(), ::tolower);
    
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

std::string getFileExtension(const std::string& filename) {
    auto pos = filename.find_last_of('.');
    if (pos == std::string::npos) {
        return "";
    }
    return filename.substr(pos);
}

bool isExcelFile(const std::string& filename) {
    std::string ext = getFileExtension(filename);
    std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);
    
    return ext == XLSX_EXTENSION || ext == XLS_EXTENSION;
}

} // namespace TXTypes
} // namespace TinaXlsx 