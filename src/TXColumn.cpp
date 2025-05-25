#include "TinaXlsx/TXColumn.hpp"
#include <algorithm>
#include <cctype>

namespace TinaXlsx {

TXTypes::ColIndex TXColumn::colNameToIndex(const std::string& name) {
    if (name.empty()) {
        return TXTypes::INVALID_COL;
    }
    
    std::string upperName = name;
    std::transform(upperName.begin(), upperName.end(), upperName.begin(), ::toupper);
    
    // 检查是否只包含字母
    for (char c : upperName) {
        if (c < 'A' || c > 'Z') {
            return TXTypes::INVALID_COL;
        }
    }
    
    TXTypes::ColIndex result = 0;
    for (char c : upperName) {
        result = result * 26 + (c - 'A' + 1);
    }
    
    // 检查是否超出Excel限制
    if (result > TXTypes::MAX_COLS) {
        return TXTypes::INVALID_COL;
    }
    
    return result;
}

std::string TXColumn::colIndexToName(TXTypes::ColIndex index) {
    if (!TXTypes::isValidCol(index)) {
        return "";
    }
    
    std::string result;
    TXTypes::ColIndex temp = index;
    
    while (temp > 0) {
        temp--; // 调整为0-based
        char letter = static_cast<char>('A' + (temp % 26));
        result = letter + result;
        temp /= 26;
    }
    
    return result;
}

} // namespace TinaXlsx 