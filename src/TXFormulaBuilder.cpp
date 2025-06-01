#include "TinaXlsx/TXFormulaBuilder.hpp"
#include <sstream>

namespace TinaXlsx {

// ==================== 统计函数 ====================

std::string TXFormulaBuilder::sum(const TXRange& range) {
    return "=SUM(" + formatRange(range) + ")";
}

std::string TXFormulaBuilder::sum(const std::string& rangeAddress) {
    return "=SUM(" + rangeAddress + ")";
}

std::string TXFormulaBuilder::average(const TXRange& range) {
    return "=AVERAGE(" + formatRange(range) + ")";
}

std::string TXFormulaBuilder::count(const TXRange& range) {
    return "=COUNT(" + formatRange(range) + ")";
}

std::string TXFormulaBuilder::countA(const TXRange& range) {
    return "=COUNTA(" + formatRange(range) + ")";
}

std::string TXFormulaBuilder::max(const TXRange& range) {
    return "=MAX(" + formatRange(range) + ")";
}

std::string TXFormulaBuilder::min(const TXRange& range) {
    return "=MIN(" + formatRange(range) + ")";
}

std::string TXFormulaBuilder::stdev(const TXRange& range) {
    return "=STDEV(" + formatRange(range) + ")";
}

std::string TXFormulaBuilder::var(const TXRange& range) {
    return "=VAR(" + formatRange(range) + ")";
}

// ==================== 条件函数 ====================

std::string TXFormulaBuilder::sumIf(const TXRange& range, const std::string& criteria, const TXRange& sumRange) {
    std::string formula = "=SUMIF(" + formatRange(range) + "," + escapeText(criteria);
    if (sumRange.isValid() && sumRange != TXRange()) {
        formula += "," + formatRange(sumRange);
    }
    formula += ")";
    return formula;
}

std::string TXFormulaBuilder::countIf(const TXRange& range, const std::string& criteria) {
    return "=COUNTIF(" + formatRange(range) + "," + escapeText(criteria) + ")";
}

std::string TXFormulaBuilder::averageIf(const TXRange& range, const std::string& criteria, const TXRange& averageRange) {
    std::string formula = "=AVERAGEIF(" + formatRange(range) + "," + escapeText(criteria);
    if (averageRange.isValid() && averageRange != TXRange()) {
        formula += "," + formatRange(averageRange);
    }
    formula += ")";
    return formula;
}

std::string TXFormulaBuilder::ifFormula(const std::string& condition, const std::string& valueIfTrue, const std::string& valueIfFalse) {
    return "=IF(" + condition + "," + escapeText(valueIfTrue) + "," + escapeText(valueIfFalse) + ")";
}

// ==================== 查找函数 ====================

std::string TXFormulaBuilder::vlookup(const std::string& lookupValue, const TXRange& tableArray, int colIndexNum, bool rangeLookup) {
    return "=VLOOKUP(" + escapeText(lookupValue) + "," + formatRange(tableArray) + "," + 
           std::to_string(colIndexNum) + "," + (rangeLookup ? "TRUE" : "FALSE") + ")";
}

std::string TXFormulaBuilder::hlookup(const std::string& lookupValue, const TXRange& tableArray, int rowIndexNum, bool rangeLookup) {
    return "=HLOOKUP(" + escapeText(lookupValue) + "," + formatRange(tableArray) + "," + 
           std::to_string(rowIndexNum) + "," + (rangeLookup ? "TRUE" : "FALSE") + ")";
}

std::string TXFormulaBuilder::index(const TXRange& array, int rowNum, int colNum) {
    std::string formula = "=INDEX(" + formatRange(array) + "," + std::to_string(rowNum);
    if (colNum > 0) {
        formula += "," + std::to_string(colNum);
    }
    formula += ")";
    return formula;
}

std::string TXFormulaBuilder::match(const std::string& lookupValue, const TXRange& lookupArray, int matchType) {
    return "=MATCH(" + escapeText(lookupValue) + "," + formatRange(lookupArray) + "," + std::to_string(matchType) + ")";
}

// ==================== 文本函数 ====================

std::string TXFormulaBuilder::concatenate(const std::vector<std::string>& values) {
    std::string formula = "=CONCATENATE(";
    for (size_t i = 0; i < values.size(); ++i) {
        if (i > 0) formula += ",";
        formula += escapeText(values[i]);
    }
    formula += ")";
    return formula;
}

std::string TXFormulaBuilder::left(const std::string& text, int numChars) {
    return "=LEFT(" + escapeText(text) + "," + std::to_string(numChars) + ")";
}

std::string TXFormulaBuilder::right(const std::string& text, int numChars) {
    return "=RIGHT(" + escapeText(text) + "," + std::to_string(numChars) + ")";
}

std::string TXFormulaBuilder::mid(const std::string& text, int startNum, int numChars) {
    return "=MID(" + escapeText(text) + "," + std::to_string(startNum) + "," + std::to_string(numChars) + ")";
}

std::string TXFormulaBuilder::len(const std::string& text) {
    return "=LEN(" + escapeText(text) + ")";
}

std::string TXFormulaBuilder::upper(const std::string& text) {
    return "=UPPER(" + escapeText(text) + ")";
}

std::string TXFormulaBuilder::lower(const std::string& text) {
    return "=LOWER(" + escapeText(text) + ")";
}

// ==================== 日期时间函数 ====================

std::string TXFormulaBuilder::today() {
    return "=TODAY()";
}

std::string TXFormulaBuilder::now() {
    return "=NOW()";
}

std::string TXFormulaBuilder::date(int year, int month, int day) {
    return "=DATE(" + std::to_string(year) + "," + std::to_string(month) + "," + std::to_string(day) + ")";
}

std::string TXFormulaBuilder::year(const std::string& dateValue) {
    return "=YEAR(" + escapeText(dateValue) + ")";
}

std::string TXFormulaBuilder::month(const std::string& dateValue) {
    return "=MONTH(" + escapeText(dateValue) + ")";
}

std::string TXFormulaBuilder::day(const std::string& dateValue) {
    return "=DAY(" + escapeText(dateValue) + ")";
}

// ==================== 数学函数 ====================

std::string TXFormulaBuilder::round(const std::string& number, int numDigits) {
    return "=ROUND(" + number + "," + std::to_string(numDigits) + ")";
}

std::string TXFormulaBuilder::abs(const std::string& number) {
    return "=ABS(" + number + ")";
}

std::string TXFormulaBuilder::power(const std::string& number, const std::string& power) {
    return "=POWER(" + number + "," + power + ")";
}

std::string TXFormulaBuilder::sqrt(const std::string& number) {
    return "=SQRT(" + number + ")";
}

// ==================== 私有辅助方法 ====================

std::string TXFormulaBuilder::formatRange(const TXRange& range) {
    return range.toAddress();
}

std::string TXFormulaBuilder::escapeText(const std::string& text) {
    // 如果文本包含引号，需要转义
    if (text.find('"') != std::string::npos) {
        std::string escaped = text;
        size_t pos = 0;
        while ((pos = escaped.find('"', pos)) != std::string::npos) {
            escaped.replace(pos, 1, "\"\"");
            pos += 2;
        }
        return "\"" + escaped + "\"";
    }
    
    // 如果文本看起来像单元格引用或数字，不加引号
    if (text.empty() || 
        (text[0] >= 'A' && text[0] <= 'Z') || 
        (text[0] >= '0' && text[0] <= '9') ||
        text[0] == '$' || text[0] == '=' || text[0] == '+' || text[0] == '-') {
        return text;
    }
    
    // 其他情况加引号
    return "\"" + text + "\"";
}

} // namespace TinaXlsx
