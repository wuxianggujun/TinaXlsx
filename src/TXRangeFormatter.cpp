#include "TinaXlsx/TXRangeFormatter.hpp"
#include <sstream>
#include <cctype>

namespace TinaXlsx {

std::string TXRangeFormatter::formatCategoryRange(const TXRange& range, const std::string& sheetName) {
    auto startCoord = range.getStart();
    auto endCoord = range.getEnd();
    
    // 类别轴使用第一列，跳过标题行
    return buildRangeString(sheetName,
                          startCoord.getCol(), row_t(startCoord.getRow().index() + 1),
                          startCoord.getCol(), endCoord.getRow());
}

std::string TXRangeFormatter::formatValueRange(const TXRange& range, const std::string& sheetName) {
    auto startCoord = range.getStart();
    auto endCoord = range.getEnd();
    
    // 数值使用第二列，跳过标题行
    column_t valueColumn(startCoord.getCol().index() + 1);
    return buildRangeString(sheetName,
                          valueColumn, row_t(startCoord.getRow().index() + 1),
                          valueColumn, endCoord.getRow());
}

std::string TXRangeFormatter::formatScatterXRange(const TXRange& range, const std::string& sheetName) {
    auto startCoord = range.getStart();
    auto endCoord = range.getEnd();
    
    // 散点图X值使用第一列，跳过标题行
    return buildRangeString(sheetName,
                          startCoord.getCol(), row_t(startCoord.getRow().index() + 1),
                          startCoord.getCol(), endCoord.getRow());
}

std::string TXRangeFormatter::formatScatterYRange(const TXRange& range, const std::string& sheetName) {
    auto startCoord = range.getStart();
    auto endCoord = range.getEnd();
    
    // 散点图Y值使用第二列，跳过标题行
    column_t valueColumn(startCoord.getCol().index() + 1);
    return buildRangeString(sheetName,
                          valueColumn, row_t(startCoord.getRow().index() + 1),
                          valueColumn, endCoord.getRow());
}

std::string TXRangeFormatter::formatPieLabelRange(const TXRange& range, const std::string& sheetName) {
    // 饼图标签范围与类别轴相同
    return formatCategoryRange(range, sheetName);
}

std::string TXRangeFormatter::formatPieValueRange(const TXRange& range, const std::string& sheetName) {
    // 饼图数值范围与普通数值范围相同
    return formatValueRange(range, sheetName);
}

std::string TXRangeFormatter::buildRangeString(const std::string& sheetName,
                                             const column_t& startCol, const row_t& startRow,
                                             const column_t& endCol, const row_t& endRow) {
    std::ostringstream oss;
    
    // 转义工作表名称
    std::string escapedSheetName = escapeSheetName(sheetName);
    
    // 构建范围字符串：SheetName!$A$1:$B$10
    oss << escapedSheetName << "!$"
        << startCol.column_string() << "$" << startRow.index()
        << ":$"
        << endCol.column_string() << "$" << endRow.index();
    
    return oss.str();
}

std::string TXRangeFormatter::escapeSheetName(const std::string& sheetName) {
    // 检查工作表名称是否需要转义
    bool needsEscape = false;

    // 如果包含空格、特殊字符、非ASCII字符或以数字开头，需要用单引号包围
    for (unsigned char c : sheetName) {
        // 检查特殊字符
        if (c == ' ' || c == '\'' || c == '!' || c == '#' || c == '$' ||
            c == '%' || c == '&' || c == '(' || c == ')' || c == '*' ||
            c == '+' || c == ',' || c == '-' || c == '.' || c == '/' ||
            c == ':' || c == ';' || c == '<' || c == '=' || c == '>' ||
            c == '?' || c == '@' || c == '[' || c == '\\' || c == ']' ||
            c == '^' || c == '`' || c == '{' || c == '|' || c == '}' || c == '~') {
            needsEscape = true;
            break;
        }

        // 检查非ASCII字符（如中文）
        if (c > 127) {
            needsEscape = true;
            break;
        }
    }

    // 检查是否以数字开头（只检查ASCII数字）
    if (!sheetName.empty() && sheetName[0] >= '0' && sheetName[0] <= '9') {
        needsEscape = true;
    }

    if (needsEscape) {
        std::string escaped = "'";
        for (char c : sheetName) {
            if (c == '\'') {
                escaped += "''"; // 单引号需要双写
            } else {
                escaped += c;
            }
        }
        escaped += "'";
        return escaped;
    }

    return sheetName;
}

} // namespace TinaXlsx
