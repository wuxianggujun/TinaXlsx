#include "TinaXlsx/TXFormula.hpp"
#include "TinaXlsx/TXSheet.hpp"

#include "TinaXlsx/TXCoordinate.hpp"
#include <algorithm>
#include <cmath>
#include <chrono>
#include <sstream>
#include <regex>
#include <cctype>

namespace TinaXlsx {

// ==================== CellReference 实现 ====================

std::string TXFormula::CellReference::toString() const {
    std::string result;
    
    if (!sheetName.empty()) {
        result += sheetName + "!";
    }
    
    if (absoluteCol) result += "$";
    
    // 转换列号为字母
    column_t tempCol = col;
    std::string colStr;
    while (tempCol.index() > 0) {
        column_t::index_t val = tempCol.index() - 1;  // 转换为0-based
        colStr = char('A' + (val % 26)) + colStr;
        tempCol = column_t(val / 26);
    }
    result += colStr;
    
    if (absoluteRow) result += "$";
    result += std::to_string(row.index());
    
    return result;
}

TXFormula::CellReference TXFormula::CellReference::fromString(const std::string& ref) {
    CellReference result;
    
    std::string workingRef = ref;
    
    // 解析工作表名称
    size_t sheetSep = workingRef.find('!');
    if (sheetSep != std::string::npos) {
        result.sheetName = workingRef.substr(0, sheetSep);
        workingRef = workingRef.substr(sheetSep + 1);
    }
    
    // 解析绝对/相对引用
    size_t pos = 0;
    if (pos < workingRef.length() && workingRef[pos] == '$') {
        result.absoluteCol = true;
        pos++;
    }
    
    // 解析列
    std::string colStr;
    while (pos < workingRef.length() && std::isalpha(workingRef[pos])) {
        colStr += std::toupper(workingRef[pos]);
        pos++;
    }
    
    // 转换列字母为数字
    column_t::index_t colNum = 0;
    for (char c : colStr) {
        colNum = colNum * 26 + (c - 'A' + 1);
    }
    result.col = column_t(colNum);
    
    // 检查行的绝对引用
    if (pos < workingRef.length() && workingRef[pos] == '$') {
        result.absoluteRow = true;
        pos++;
    }
    
    // 解析行
    std::string rowStr = workingRef.substr(pos);
    if (!rowStr.empty()) {
        result.row = row_t(std::stoi(rowStr));
    }
    
    return result;
}

bool TXFormula::CellReference::isValid() const {
    return row.index() > 0 && col.index() > 0;
}

// ==================== RangeReference 实现 ====================

std::string TXFormula::RangeReference::toString() const {
    return start.toString() + ":" + end.toString();
}

TXFormula::RangeReference TXFormula::RangeReference::fromString(const std::string& range) {
    size_t colonPos = range.find(':');
    if (colonPos == std::string::npos) {
        // 单个单元格
        CellReference ref = CellReference::fromString(range);
        return RangeReference(ref, ref);
    }
    
    CellReference startRef = CellReference::fromString(range.substr(0, colonPos));
    CellReference endRef = CellReference::fromString(range.substr(colonPos + 1));
    
    return RangeReference(startRef, endRef);
}

bool TXFormula::RangeReference::contains(const CellReference& cell) const {
    return cell.row.index() >= start.row.index() && cell.row.index() <= end.row.index() &&
           cell.col.index() >= start.col.index() && cell.col.index() <= end.col.index();
}

std::vector<TXFormula::CellReference> TXFormula::RangeReference::getAllCells() const {
    std::vector<CellReference> cells;
    
    for (row_t r(start.row.index()); r.index() <= end.row.index(); ++r) {
        for (column_t c(start.col.index()); c.index() <= end.col.index(); ++c) {
            cells.emplace_back(r, c);
        }
    }
    
    return cells;
}

bool TXFormula::RangeReference::isValid() const {
    return start.isValid() && end.isValid() && 
           start.row.index() <= end.row.index() && start.col.index() <= end.col.index();
}

// ==================== TXFormula 实现 ====================

TXFormula::TXFormula() : lastError_(FormulaError::None) {
    registerBuiltinFunctions();
}

TXFormula::TXFormula(const std::string& formula) 
    : formulaString_(formula), lastError_(FormulaError::None) {
    registerBuiltinFunctions();
    updateDependencies();
}

TXFormula::~TXFormula() = default;

TXFormula::TXFormula(const TXFormula& other)
    : formulaString_(other.formulaString_)
    , lastError_(other.lastError_)
    , dependencies_(other.dependencies_)
    , customFunctions_(other.customFunctions_) {
}

TXFormula& TXFormula::operator=(const TXFormula& other) {
    if (this != &other) {
        formulaString_ = other.formulaString_;
        lastError_ = other.lastError_;
        dependencies_ = other.dependencies_;
        customFunctions_ = other.customFunctions_;
    }
    return *this;
}

TXFormula::TXFormula(TXFormula&& other) noexcept
    : formulaString_(std::move(other.formulaString_))
    , lastError_(other.lastError_)
    , dependencies_(std::move(other.dependencies_))
    , customFunctions_(std::move(other.customFunctions_)) {
}

TXFormula& TXFormula::operator=(TXFormula&& other) noexcept {
    if (this != &other) {
        formulaString_ = std::move(other.formulaString_);
        lastError_ = other.lastError_;
        dependencies_ = std::move(other.dependencies_);
        customFunctions_ = std::move(other.customFunctions_);
    }
    return *this;
}

bool TXFormula::parseFormula(const std::string& formula) {
    formulaString_ = formula;
    lastError_ = FormulaError::None;
    updateDependencies();
    return true; // 简化实现，总是返回true
}

TXFormula::FormulaValue TXFormula::evaluate(const TXSheet* sheet, row_t currentRow, column_t currentCol) {
    if (!sheet) {
        lastError_ = FormulaError::Reference;
        return std::monostate{};
    }
    
    try {
        return evaluateExpression(formulaString_, sheet, currentRow, currentCol);
    } catch (...) {
        lastError_ = FormulaError::Syntax;
        return std::monostate{};
    }
}

const std::string& TXFormula::getFormulaString() const {
    return formulaString_;
}

void TXFormula::setFormulaString(const std::string& formula) {
    formulaString_ = formula;
    lastError_ = FormulaError::None;
    updateDependencies();
}

TXFormula::FormulaError TXFormula::getLastError() const {
    return lastError_;
}

std::string TXFormula::getErrorDescription() const {
    switch (lastError_) {
        case FormulaError::None: return "No error";
        case FormulaError::Syntax: return "Syntax error";
        case FormulaError::Reference: return "Reference error";
        case FormulaError::Name: return "Name error";
        case FormulaError::Value: return "Value error";
        case FormulaError::Division: return "Division by zero";
        case FormulaError::Circular: return "Circular reference";
        default: return "Unknown error";
    }
}

std::vector<TXFormula::CellReference> TXFormula::getDependencies() const {
    return dependencies_;
}

bool TXFormula::isValidFormula(const std::string& formula) {
    // 简化实现，检查基本格式
    if (formula.empty()) return false;
    
    // 检查是否包含函数调用或引用
    std::regex functionRegex(R"([A-Z]+\s*\()");
    std::regex cellRefRegex(R"([A-Z]+[0-9]+)");
    
    return std::regex_search(formula, functionRegex) || 
           std::regex_search(formula, cellRefRegex) ||
           formula.find_first_of("+-*/()") != std::string::npos;
}

void TXFormula::registerFunction(const std::string& name, const FormulaFunction& func) {
    customFunctions_[name] = func;
}

void TXFormula::clearCustomFunctions() {
    customFunctions_.clear();
}

// ==================== 内置函数实现 ====================

TXFormula::FormulaValue TXFormula::sumFunction(const std::vector<FormulaValue>& args) {
    double sum = 0.0;
    for (const auto& arg : args) {
        sum += valueToNumber(arg);
    }
    return sum;
}

TXFormula::FormulaValue TXFormula::averageFunction(const std::vector<FormulaValue>& args) {
    if (args.empty()) return 0.0;
    
    double sum = 0.0;
    for (const auto& arg : args) {
        sum += valueToNumber(arg);
    }
    return sum / static_cast<double>(args.size());
}

TXFormula::FormulaValue TXFormula::countFunction(const std::vector<FormulaValue>& args) {
    double count = 0.0;
    for (const auto& arg : args) {
        if (!std::holds_alternative<std::monostate>(arg)) {
            count += 1.0;
        }
    }
    return count;
}

TXFormula::FormulaValue TXFormula::maxFunction(const std::vector<FormulaValue>& args) {
    if (args.empty()) return 0.0;
    
    double maxVal = valueToNumber(args[0]);
    for (size_t i = 1; i < args.size(); ++i) {
        maxVal = std::max(maxVal, valueToNumber(args[i]));
    }
    return maxVal;
}

TXFormula::FormulaValue TXFormula::minFunction(const std::vector<FormulaValue>& args) {
    if (args.empty()) return 0.0;
    
    double minVal = valueToNumber(args[0]);
    for (size_t i = 1; i < args.size(); ++i) {
        minVal = std::min(minVal, valueToNumber(args[i]));
    }
    return minVal;
}

TXFormula::FormulaValue TXFormula::ifFunction(const std::vector<FormulaValue>& args) {
    if (args.size() < 2) return std::monostate{};
    
    bool condition = valueToBool(args[0]);
    if (condition) {
        return args[1];
    } else if (args.size() >= 3) {
        return args[2];
    }
    return std::monostate{};
}

TXFormula::FormulaValue TXFormula::concatenateFunction(const std::vector<FormulaValue>& args) {
    std::string result;
    for (const auto& arg : args) {
        result += valueToString(arg);
    }
    return result;
}

TXFormula::FormulaValue TXFormula::lenFunction(const std::vector<FormulaValue>& args) {
    if (args.empty()) return 0.0;
    
    std::string str = valueToString(args[0]);
    return static_cast<double>(str.length());
}

TXFormula::FormulaValue TXFormula::roundFunction(const std::vector<FormulaValue>& args) {
    if (args.empty()) return 0.0;
    
    double value = valueToNumber(args[0]);
    int digits = args.size() > 1 ? static_cast<int>(valueToNumber(args[1])) : 0;
    
    double factor = std::pow(10.0, digits);
    return std::round(value * factor) / factor;
}

TXFormula::FormulaValue TXFormula::nowFunction(const std::vector<FormulaValue>& args) {
    auto now = std::chrono::system_clock::now();
    auto time_t = std::chrono::system_clock::to_time_t(now);
    return static_cast<double>(time_t);
}

TXFormula::FormulaValue TXFormula::todayFunction(const std::vector<FormulaValue>& args) {
    auto now = std::chrono::system_clock::now();
    auto time_t = std::chrono::system_clock::to_time_t(now);
    // 只返回日期部分，去除时间
    time_t = (time_t / 86400) * 86400; // 86400 seconds in a day
    return static_cast<double>(time_t);
}

// ==================== 工具函数实现 ====================

double TXFormula::valueToNumber(const FormulaValue& value) {
    if (std::holds_alternative<double>(value)) {
        return std::get<double>(value);
    } else if (std::holds_alternative<int64_t>(value)) {
        return static_cast<double>(std::get<int64_t>(value));
    } else if (std::holds_alternative<bool>(value)) {
        return std::get<bool>(value) ? 1.0 : 0.0;
    } else if (std::holds_alternative<std::string>(value)) {
        try {
            return std::stod(std::get<std::string>(value));
        } catch (...) {
            return 0.0;
        }
    }
    return 0.0;
}

std::string TXFormula::valueToString(const FormulaValue& value) {
    if (std::holds_alternative<std::string>(value)) {
        return std::get<std::string>(value);
    } else if (std::holds_alternative<double>(value)) {
        return std::to_string(std::get<double>(value));
    } else if (std::holds_alternative<int64_t>(value)) {
        return std::to_string(std::get<int64_t>(value));
    } else if (std::holds_alternative<bool>(value)) {
        return std::get<bool>(value) ? "TRUE" : "FALSE";
    }
    return "";
}

bool TXFormula::valueToBool(const FormulaValue& value) {
    if (std::holds_alternative<bool>(value)) {
        return std::get<bool>(value);
    } else if (std::holds_alternative<double>(value)) {
        return std::get<double>(value) != 0.0;
    } else if (std::holds_alternative<int64_t>(value)) {
        return std::get<int64_t>(value) != 0;
    } else if (std::holds_alternative<std::string>(value)) {
        const auto& str = std::get<std::string>(value);
        std::string lower_str = str;
        std::transform(lower_str.begin(), lower_str.end(), lower_str.begin(), ::tolower);
        return lower_str == "true" || lower_str == "1" || lower_str == "yes";
    }
    return false;
}

TXFormula::FormulaValue TXFormula::valueFromString(const std::string& str) {
    if (str.empty()) return std::monostate{};
    
    // 尝试解析为数字
    try {
        if (str.find('.') != std::string::npos) {
            return std::stod(str);
        } else {
            return std::stoll(str);
        }
    } catch (...) {
        // 如果不是数字，检查是否为布尔值
        std::string lower_str = str;
        std::transform(lower_str.begin(), lower_str.end(), lower_str.begin(), ::tolower);
        if (lower_str == "true" || lower_str == "false") {
            return lower_str == "true";
        }
        
        // 否则作为字符串返回
        return str;
    }
}

bool TXFormula::valuesEqual(const FormulaValue& a, const FormulaValue& b) {
    return a == b;
}

// ==================== 私有辅助方法 ====================

void TXFormula::registerBuiltinFunctions() {
    customFunctions_["SUM"] = sumFunction;
    customFunctions_["AVERAGE"] = averageFunction;
    customFunctions_["COUNT"] = countFunction;
    customFunctions_["MAX"] = maxFunction;
    customFunctions_["MIN"] = minFunction;
    customFunctions_["IF"] = ifFunction;
    customFunctions_["CONCATENATE"] = concatenateFunction;
    customFunctions_["LEN"] = lenFunction;
    customFunctions_["ROUND"] = roundFunction;
    customFunctions_["NOW"] = nowFunction;
    customFunctions_["TODAY"] = todayFunction;
}

TXFormula::FormulaValue TXFormula::evaluateExpression(const std::string& expr, const TXSheet* sheet, row_t currentRow, column_t currentCol) {
    std::string workingExpr = expr;
    
    // 移除前导的 '=' 符号
    if (!workingExpr.empty() && workingExpr[0] == '=') {
        workingExpr = workingExpr.substr(1);
    }
    
    // 替换单元格引用
    workingExpr = replaceReferences(workingExpr, sheet, currentRow, currentCol);
    
    // 检查函数调用
    size_t parenPos = workingExpr.find('(');
    if (parenPos != std::string::npos) {
        std::string funcName = workingExpr.substr(0, parenPos);
        std::string argsStr = workingExpr.substr(parenPos + 1);
        
        // 移除最后的 ')'
        if (!argsStr.empty() && argsStr.back() == ')') {
            argsStr.pop_back();
        }
        
        auto args = parseArguments(argsStr, sheet, currentRow, currentCol);
        return callFunction(funcName, args);
    }
    
    // 否则作为简单表达式计算
    return evaluateSimpleExpression(workingExpr);
}

std::vector<TXFormula::FormulaValue> TXFormula::parseArguments(const std::string& args, const TXSheet* sheet, row_t currentRow, column_t currentCol) {
    std::vector<FormulaValue> result;
    
    if (args.empty()) return result;
    
    std::stringstream ss(args);
    std::string item;
    
    while (std::getline(ss, item, ',')) {
        // 去除前后空格
        item.erase(0, item.find_first_not_of(" \t"));
        item.erase(item.find_last_not_of(" \t") + 1);
        
        if (!item.empty()) {
            auto value = evaluateExpression(item, sheet, currentRow, currentCol);
            result.push_back(value);
        }
    }
    
    return result;
}

TXFormula::FormulaValue TXFormula::callFunction(const std::string& funcName, const std::vector<FormulaValue>& args) {
    auto it = customFunctions_.find(funcName);
    if (it != customFunctions_.end()) {
        return it->second(args);
    }
    
    lastError_ = FormulaError::Name;
    return std::monostate{};
}

void TXFormula::updateDependencies() {
    dependencies_.clear();
    
    // 使用正则表达式查找单元格引用
    std::regex cellRefRegex(R"([A-Z]+[0-9]+)");
    std::sregex_iterator iter(formulaString_.begin(), formulaString_.end(), cellRefRegex);
    std::sregex_iterator end;
    
    for (; iter != end; ++iter) {
        std::string ref = iter->str();
        CellReference cellRef = CellReference::fromString(ref);
        if (cellRef.isValid()) {
            dependencies_.push_back(cellRef);
        }
    }
}

std::string TXFormula::replaceReferences(const std::string& expression, const TXSheet* sheet, row_t currentRow, column_t currentCol) {
    std::string result = expression;
    
    // 查找并替换单元格引用
    std::regex cellRefRegex(R"([A-Z]+[0-9]+)");
    std::sregex_iterator iter(result.begin(), result.end(), cellRefRegex);
    std::sregex_iterator end;
    
    std::vector<std::pair<std::string, std::string>> replacements;
    
    for (; iter != end; ++iter) {
        std::string ref = iter->str();
        CellReference cellRef = CellReference::fromString(ref);
        
        if (cellRef.isValid()) {
            TXCoordinate coord(cellRef.row, cellRef.col);
            auto cellValue = sheet->getCellValue(coord);
            
            std::string value = "0";
            if (std::holds_alternative<double>(cellValue)) {
                value = std::to_string(std::get<double>(cellValue));
            } else if (std::holds_alternative<int64_t>(cellValue)) {
                value = std::to_string(std::get<int64_t>(cellValue));
            } else if (std::holds_alternative<std::string>(cellValue)) {
                value = "\"" + std::get<std::string>(cellValue) + "\"";
            } else if (std::holds_alternative<bool>(cellValue)) {
                value = std::get<bool>(cellValue) ? "1" : "0";
            }
            
            replacements.push_back({ref, value});
        }
    }
    
    // 执行替换
    for (const auto& replacement : replacements) {
        size_t pos = 0;
        while ((pos = result.find(replacement.first, pos)) != std::string::npos) {
            result.replace(pos, replacement.first.length(), replacement.second);
            pos += replacement.second.length();
        }
    }
    
    return result;
}

TXFormula::FormulaValue TXFormula::evaluateSimpleExpression(const std::string& expr) {
    std::string cleanExpr = expr;
    
    // 移除空格
    cleanExpr.erase(std::remove(cleanExpr.begin(), cleanExpr.end(), ' '), cleanExpr.end());
    
    // 检查是否是字符串（被引号包围）
    if (cleanExpr.length() >= 2 && cleanExpr[0] == '"' && cleanExpr.back() == '"') {
        return cleanExpr.substr(1, cleanExpr.length() - 2);
    }
    
    try {
        // 简单的算术运算
        if (cleanExpr.find('+') != std::string::npos) {
            size_t pos = cleanExpr.find_last_of('+');
            double left = std::stod(cleanExpr.substr(0, pos));
            double right = std::stod(cleanExpr.substr(pos + 1));
            return left + right;
        } else if (cleanExpr.find('-') != std::string::npos && cleanExpr[0] != '-') {
            size_t pos = cleanExpr.find_last_of('-');
            double left = std::stod(cleanExpr.substr(0, pos));
            double right = std::stod(cleanExpr.substr(pos + 1));
            return left - right;
        } else if (cleanExpr.find('*') != std::string::npos) {
            size_t pos = cleanExpr.find_last_of('*');
            double left = std::stod(cleanExpr.substr(0, pos));
            double right = std::stod(cleanExpr.substr(pos + 1));
            return left * right;
        } else if (cleanExpr.find('/') != std::string::npos) {
            size_t pos = cleanExpr.find_last_of('/');
            double left = std::stod(cleanExpr.substr(0, pos));
            double right = std::stod(cleanExpr.substr(pos + 1));
            if (right != 0.0) {
                return left / right;
            } else {
                lastError_ = FormulaError::Division;
                return std::monostate{};
            }
        }
        
        // 尝试解析为数字
        return valueFromString(cleanExpr);
        
    } catch (...) {
        lastError_ = FormulaError::Syntax;
        return std::monostate{};
    }
}

} // namespace TinaXlsx
