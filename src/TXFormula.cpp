#include "TinaXlsx/TXFormula.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCoordinate.hpp"
#include <sstream>
#include <algorithm>
#include <cctype>
#include <cmath>
#include <chrono>
#include <regex>
#include <stack>
#include <iostream>

namespace TinaXlsx {

// ==================== CellReference实现 ====================

std::string TXFormula::CellReference::toString() const {
    std::string result;
    
    if (!sheetName.empty()) {
        result += sheetName + "!";
    }
    
    if (absoluteCol) result += "$";
    
    // 转换列号为字母
    TXTypes::ColIndex tempCol = col;
    std::string colStr;
    while (tempCol > 0) {
        tempCol--;
        colStr = char('A' + (tempCol % 26)) + colStr;
        tempCol /= 26;
    }
    result += colStr;
    
    if (absoluteRow) result += "$";
    result += std::to_string(row);
    
    return result;
}

TXFormula::CellReference TXFormula::CellReference::fromString(const std::string& ref) {
    CellReference result;
    
    std::string str = ref;
    
    // 处理工作表名称
    size_t exclamationPos = str.find('!');
    if (exclamationPos != std::string::npos) {
        result.sheetName = str.substr(0, exclamationPos);
        str = str.substr(exclamationPos + 1);
    }
    
    // 解析单元格引用
    std::regex cellRegex(R"((\$?)([A-Z]+)(\$?)([0-9]+))");
    std::smatch match;
    
    if (std::regex_match(str, match, cellRegex)) {
        result.absoluteCol = !match[1].str().empty();
        
        // 转换字母为列号
        std::string colStr = match[2].str();
        TXTypes::ColIndex colNum = 0;
        for (char c : colStr) {
            colNum = colNum * 26 + (c - 'A' + 1);
        }
        result.col = colNum;
        
        result.absoluteRow = !match[3].str().empty();
        result.row = std::stoul(match[4].str());
    }
    
    return result;
}

// ==================== RangeReference实现 ====================

std::string TXFormula::RangeReference::toString() const {
    return start.toString() + ":" + end.toString();
}

TXFormula::RangeReference TXFormula::RangeReference::fromString(const std::string& range) {
    RangeReference result;
    
    size_t colonPos = range.find(':');
    if (colonPos != std::string::npos) {
        result.start = CellReference::fromString(range.substr(0, colonPos));
        result.end = CellReference::fromString(range.substr(colonPos + 1));
    }
    
    return result;
}

bool TXFormula::RangeReference::contains(const CellReference& cell) const {
    return cell.row >= start.row && cell.row <= end.row &&
           cell.col >= start.col && cell.col <= end.col;
}

std::vector<TXFormula::CellReference> TXFormula::RangeReference::getAllCells() const {
    std::vector<CellReference> result;
    result.reserve((end.row - start.row + 1) * (end.col - start.col + 1));
    
    for (TXTypes::RowIndex r = start.row; r <= end.row; ++r) {
        for (TXTypes::ColIndex c = start.col; c <= end.col; ++c) {
            result.emplace_back(r, c);
        }
    }
    
    return result;
}

// ==================== TXFormula::Impl实现 ====================

class TXFormula::Impl {
public:
    std::string formulaString_;
    FormulaError lastError_ = FormulaError::None;
    std::vector<CellReference> dependencies_;
    std::unordered_map<std::string, FormulaFunction> customFunctions_;
    
    // 预编译的表达式树（简化版）
    struct Token {
        enum Type { Number, String, Function, Operator, CellRef, RangeRef, LeftParen, RightParen };
        Type type;
        std::string value;
        std::vector<Token> args;
        
        Token(Type t, const std::string& v) : type(t), value(v) {}
    };
    
    std::vector<Token> tokens_;
    
    Impl() {
        registerBuiltinFunctions();
    }
    
    void registerBuiltinFunctions() {
        customFunctions_["SUM"] = TXFormula::sumFunction;
        customFunctions_["AVERAGE"] = TXFormula::averageFunction;
        customFunctions_["COUNT"] = TXFormula::countFunction;
        customFunctions_["MAX"] = TXFormula::maxFunction;
        customFunctions_["MIN"] = TXFormula::minFunction;
        customFunctions_["IF"] = TXFormula::ifFunction;
        customFunctions_["CONCATENATE"] = TXFormula::concatenateFunction;
        customFunctions_["LEN"] = TXFormula::lenFunction;
        customFunctions_["ROUND"] = TXFormula::roundFunction;
        customFunctions_["NOW"] = TXFormula::nowFunction;
        customFunctions_["TODAY"] = TXFormula::todayFunction;
    }
    
    bool parseFormula(const std::string& formula) {
        formulaString_ = formula;
        lastError_ = FormulaError::None;
        dependencies_.clear();
        tokens_.clear();
        
        try {
            tokenize(formula);
            return true;
        } catch (...) {
            lastError_ = FormulaError::Syntax;
            return false;
        }
    }
    
    void tokenize(const std::string& formula) {
        size_t i = 0;
        while (i < formula.length()) {
            char c = formula[i];
            
            if (std::isspace(c)) {
                i++;
                continue;
            }
            
            if (std::isdigit(c) || c == '.') {
                // 数字
                size_t start = i;
                while (i < formula.length() && (std::isdigit(formula[i]) || formula[i] == '.')) {
                    i++;
                }
                tokens_.emplace_back(Token::Number, formula.substr(start, i - start));
            } else if (c == '"') {
                // 字符串
                i++; // 跳过开始引号
                size_t start = i;
                while (i < formula.length() && formula[i] != '"') {
                    i++;
                }
                tokens_.emplace_back(Token::String, formula.substr(start, i - start));
                if (i < formula.length()) i++; // 跳过结束引号
            } else if (std::isalpha(c)) {
                // 函数名或单元格引用或范围引用
                size_t start = i;
                while (i < formula.length() && (std::isalnum(formula[i]) || formula[i] == '_' || formula[i] == ':')) {
                    i++;
                }
                std::string name = formula.substr(start, i - start);
                
                // 检查是否为函数调用
                if (i < formula.length() && formula[i] == '(') {
                    tokens_.emplace_back(Token::Function, name);
                } else {
                    // 检查是否为范围引用（包含冒号）
                    if (name.find(':') != std::string::npos) {
                        tokens_.emplace_back(Token::String, name);  // 范围作为字符串处理
                    } else if (isCellReference(name)) {
                        tokens_.emplace_back(Token::CellRef, name);
                        dependencies_.push_back(CellReference::fromString(name));
                    } else {
                        tokens_.emplace_back(Token::String, name);
                    }
                }
            } else if (c == '(' || c == ')' || c == ',' || c == '+' || c == '-' || c == '*' || c == '/') {
                if (c == '(') tokens_.emplace_back(Token::LeftParen, std::string(1, c));
                else if (c == ')') tokens_.emplace_back(Token::RightParen, std::string(1, c));
                else tokens_.emplace_back(Token::Operator, std::string(1, c));
                i++;
            } else {
                i++;
            }
        }
    }
    
    bool isCellReference(const std::string& str) {
        std::regex cellPattern(R"([A-Z]+[0-9]+)");
        return std::regex_match(str, cellPattern);
    }
    
    FormulaValue evaluate(const TXSheet* sheet, TXTypes::RowIndex currentRow, TXTypes::ColIndex currentCol) {
        if (tokens_.empty()) {
            lastError_ = FormulaError::Syntax;
            return std::monostate{};
        }
        
        try {
            return evaluateTokens(tokens_, sheet, currentRow, currentCol);
        } catch (...) {
            lastError_ = FormulaError::Value;
            return std::monostate{};
        }
    }
    
    FormulaValue evaluateTokens(const std::vector<Token>& tokens, const TXSheet* sheet, 
                               TXTypes::RowIndex currentRow, TXTypes::ColIndex currentCol) {
        (void)currentRow;  // 标记参数为有意未使用
        (void)currentCol;  // 标记参数为有意未使用
        
        if (tokens.empty()) return std::monostate{};
        
        // 检查是否为SUM函数调用
        if (tokens.size() >= 4 && tokens[0].type == Token::Function && tokens[0].value == "SUM") {
            // 简化处理：SUM(A1:A3) 格式
            if (tokens.size() >= 4 && tokens[2].type == Token::String) {
                std::string rangeStr = tokens[2].value;
                
                // 解析范围 A1:A3
                size_t colonPos = rangeStr.find(':');
                if (colonPos != std::string::npos) {
                    std::string startCell = rangeStr.substr(0, colonPos);
                    std::string endCell = rangeStr.substr(colonPos + 1);
                    
                    auto startRef = CellReference::fromString(startCell);
                    auto endRef = CellReference::fromString(endCell);
                    
                    double sum = 0.0;
                    if (sheet) {
                        // 确保正确的范围遍历
                        TXTypes::RowIndex startRow = std::min(startRef.row, endRef.row);
                        TXTypes::RowIndex endRow = std::max(startRef.row, endRef.row);
                        TXTypes::ColIndex startCol = std::min(startRef.col, endRef.col);
                        TXTypes::ColIndex endCol = std::max(startRef.col, endRef.col);
                        
                        for (TXTypes::RowIndex row = startRow; row <= endRow; ++row) {
                            for (TXTypes::ColIndex col = startCol; col <= endCol; ++col) {
                                auto cellValue = sheet->getCellValue(row, col);
                                double numValue = valueToNumber(convertCellValue(cellValue));
                                sum += numValue;
                            }
                        }
                    }
                    return sum;
                }
            }
        }
        
        // 原有的简化版表达式求值
        std::stack<FormulaValue> values;
        std::stack<std::string> operators;
        
        for (const auto& token : tokens) {
            switch (token.type) {
                case Token::Number:
                    values.push(std::stod(token.value));
                    break;
                    
                case Token::String:
                    values.push(token.value);
                    break;
                    
                case Token::CellRef: {
                    auto cellRef = CellReference::fromString(token.value);
                    if (sheet) {
                        auto cellValue = sheet->getCellValue(cellRef.row, cellRef.col);
                        values.push(convertCellValue(cellValue));
                    } else {
                        values.push(0.0);
                    }
                    break;
                }
                
                case Token::Function: {
                    auto it = customFunctions_.find(token.value);
                    if (it != customFunctions_.end()) {
                        // 简化版：假设函数参数已经计算好
                        std::vector<FormulaValue> args;
                        values.push(it->second(args));
                    }
                    break;
                }
                
                case Token::Operator:
                    if (values.size() >= 2) {
                        auto right = values.top(); values.pop();
                        auto left = values.top(); values.pop();
                        values.push(applyOperator(token.value, left, right));
                    }
                    break;
                    
                default:
                    break;
            }
        }
        
        return values.empty() ? std::monostate{} : values.top();
    }
    
    FormulaValue convertCellValue(const std::variant<std::monostate, std::string, double, int64_t, bool>& value) {
        if (std::holds_alternative<std::string>(value)) {
            return std::get<std::string>(value);
        } else if (std::holds_alternative<double>(value)) {
            return std::get<double>(value);
        } else if (std::holds_alternative<int64_t>(value)) {
            return std::get<int64_t>(value);
        } else if (std::holds_alternative<bool>(value)) {
            return std::get<bool>(value);
        }
        return std::monostate{};
    }
    
    FormulaValue applyOperator(const std::string& op, const FormulaValue& left, const FormulaValue& right) {
        double leftNum = TXFormula::valueToNumber(left);
        double rightNum = TXFormula::valueToNumber(right);
        
        if (op == "+") return leftNum + rightNum;
        if (op == "-") return leftNum - rightNum;
        if (op == "*") return leftNum * rightNum;
        if (op == "/") {
            if (rightNum == 0.0) {
                lastError_ = FormulaError::Division;
                return std::monostate{};
            }
            return leftNum / rightNum;
        }
        
        return std::monostate{};
    }
};

// ==================== TXFormula公共接口实现 ====================

TXFormula::TXFormula() : pImpl(std::make_unique<Impl>()) {}

TXFormula::~TXFormula() = default;

TXFormula::TXFormula(TXFormula&& other) noexcept : pImpl(std::move(other.pImpl)) {}

TXFormula& TXFormula::operator=(TXFormula&& other) noexcept {
    if (this != &other) {
        pImpl = std::move(other.pImpl);
    }
    return *this;
}

bool TXFormula::parseFormula(const std::string& formula) {
    return pImpl->parseFormula(formula);
}

TXFormula::FormulaValue TXFormula::evaluate(const TXSheet* sheet, TXTypes::RowIndex currentRow, TXTypes::ColIndex currentCol) {
    return pImpl->evaluate(sheet, currentRow, currentCol);
}

const std::string& TXFormula::getFormulaString() const {
    return pImpl->formulaString_;
}

TXFormula::FormulaError TXFormula::getLastError() const {
    return pImpl->lastError_;
}

std::string TXFormula::getErrorDescription() const {
    switch (pImpl->lastError_) {
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
    return pImpl->dependencies_;
}

bool TXFormula::isValidFormula(const std::string& formula) {
    TXFormula temp;
    return temp.parseFormula(formula);
}

void TXFormula::registerFunction(const std::string& name, const FormulaFunction& func) {
    pImpl->customFunctions_[name] = func;
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
    return sum / args.size();
}

TXFormula::FormulaValue TXFormula::countFunction(const std::vector<FormulaValue>& args) {
    int64_t count = 0;
    for (const auto& arg : args) {
        if (!std::holds_alternative<std::monostate>(arg)) {
            count++;
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
    } else if (args.size() > 2) {
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
    if (args.empty()) return 0;
    
    std::string str = valueToString(args[0]);
    return static_cast<int64_t>(str.length());
}

TXFormula::FormulaValue TXFormula::roundFunction(const std::vector<FormulaValue>& args) {
    if (args.empty()) return 0.0;
    
    double value = valueToNumber(args[0]);
    int64_t digits = args.size() > 1 ? static_cast<int64_t>(valueToNumber(args[1])) : 0;
    
    double factor = std::pow(10.0, digits);
    return std::round(value * factor) / factor;
}

TXFormula::FormulaValue TXFormula::nowFunction(const std::vector<FormulaValue>& args) {
    (void)args;  // 标记参数为有意未使用
    auto now = std::chrono::system_clock::now();
    auto time_t = std::chrono::system_clock::to_time_t(now);
    return static_cast<double>(time_t) / 86400.0 + 25569.0; // Excel日期格式
}

TXFormula::FormulaValue TXFormula::todayFunction(const std::vector<FormulaValue>& args) {
    (void)args;  // 标记参数为有意未使用
    auto now = std::chrono::system_clock::now();
    auto time_t = std::chrono::system_clock::to_time_t(now);
    return static_cast<double>(time_t / 86400) + 25569.0; // Excel日期格式，仅日期部分
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
        const std::string& str = std::get<std::string>(value);
        return !str.empty() && str != "0" && str != "FALSE" && str != "false";
    }
    return false;
}

} // namespace TinaXlsx 