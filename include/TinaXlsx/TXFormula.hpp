#pragma once

#include "TXTypes.hpp"
#include <string>
#include <vector>
#include <unordered_map>
#include <memory>
#include <functional>
#include <variant>

namespace TinaXlsx {

// Forward declarations
class TXSheet;

/**
 * @brief Excel公式处理类
 * 
 * 提供公式解析、计算和生成功能，支持常用的Excel函数
 */
class TXFormula {
public:
    /**
     * @brief 公式值类型
     */
    using FormulaValue = cell_value_t;
    
    /**
     * @brief 公式函数类型
     */
    using FormulaFunction = std::function<FormulaValue(const std::vector<FormulaValue>&)>;
    
    /**
     * @brief 公式错误类型
     */
    enum class FormulaError {
        None,           ///< 无错误
        Syntax,         ///< 语法错误
        Reference,      ///< 引用错误
        Name,           ///< 名称错误
        Value,          ///< 值错误
        Division,       ///< 除零错误
        Circular        ///< 循环引用
    };

    /**
     * @brief 单元格引用类型
     */
    struct CellReference {
        row_t row;
        column_t col;
        bool absoluteRow;
        bool absoluteCol;
        std::string sheetName;       ///< 工作表名称（跨表引用）
        
        CellReference() : row(1), col(1), absoluteRow(false), absoluteCol(false) {}
        CellReference(row_t r, column_t c) : row(r), col(c), absoluteRow(false), absoluteCol(false) {}
        
        std::string toString() const;
        static CellReference fromString(const std::string& ref);
        bool isValid() const;
    };
    
    /**
     * @brief 范围引用类型
     */
    struct RangeReference {
        CellReference start;
        CellReference end;
        
        RangeReference() = default;
        RangeReference(const CellReference& s, const CellReference& e) : start(s), end(e) {}
        
        std::string toString() const;
        static RangeReference fromString(const std::string& range);
        bool contains(const CellReference& cell) const;
        std::vector<CellReference> getAllCells() const;
        bool isValid() const;
    };

public:
    TXFormula();
    explicit TXFormula(const std::string& formula);
    ~TXFormula();
    
    // 支持拷贝构造和赋值
    TXFormula(const TXFormula& other);
    TXFormula& operator=(const TXFormula& other);
    
    // 支持移动构造和赋值
    TXFormula(TXFormula&& other) noexcept;
    TXFormula& operator=(TXFormula&& other) noexcept;

    /**
     * @brief 解析公式
     * @param formula 公式字符串（不含等号）
     * @return 成功返回true，失败返回false
     */
    bool parseFormula(const std::string& formula);

    /**
     * @brief 计算公式结果
     * @param sheet 当前工作表
     * @param currentRow 当前单元格行号
     * @param currentCol 当前单元格列号
     * @return 计算结果
     */
    FormulaValue evaluate(const TXSheet* sheet, row_t currentRow, column_t currentCol);

    /**
     * @brief 获取公式字符串
     * @return 公式字符串
     */
    const std::string& getFormulaString() const;

    /**
     * @brief 设置公式字符串
     * @param formula 公式字符串
     */
    void setFormulaString(const std::string& formula);

    /**
     * @brief 获取最后的错误信息
     * @return 错误类型
     */
    FormulaError getLastError() const;

    /**
     * @brief 获取错误描述
     * @return 错误描述字符串
     */
    std::string getErrorDescription() const;

    /**
     * @brief 获取公式依赖的单元格
     * @return 依赖的单元格引用列表
     */
    std::vector<CellReference> getDependencies() const;

    /**
     * @brief 检查是否为有效的公式
     * @param formula 公式字符串
     * @return 有效返回true，否则返回false
     */
    static bool isValidFormula(const std::string& formula);

    /**
     * @brief 注册自定义函数
     * @param name 函数名称
     * @param func 函数实现
     */
    void registerFunction(const std::string& name, const FormulaFunction& func);

    /**
     * @brief 清除自定义函数
     */
    void clearCustomFunctions();

    // ==================== 内置函数 ====================

    /**
     * @brief SUM函数 - 求和
     */
    static FormulaValue sumFunction(const std::vector<FormulaValue>& args);

    /**
     * @brief AVERAGE函数 - 平均值
     */
    static FormulaValue averageFunction(const std::vector<FormulaValue>& args);

    /**
     * @brief COUNT函数 - 计数
     */
    static FormulaValue countFunction(const std::vector<FormulaValue>& args);

    /**
     * @brief MAX函数 - 最大值
     */
    static FormulaValue maxFunction(const std::vector<FormulaValue>& args);

    /**
     * @brief MIN函数 - 最小值
     */
    static FormulaValue minFunction(const std::vector<FormulaValue>& args);

    /**
     * @brief IF函数 - 条件判断
     */
    static FormulaValue ifFunction(const std::vector<FormulaValue>& args);

    /**
     * @brief CONCATENATE函数 - 字符串连接
     */
    static FormulaValue concatenateFunction(const std::vector<FormulaValue>& args);

    /**
     * @brief LEN函数 - 字符串长度
     */
    static FormulaValue lenFunction(const std::vector<FormulaValue>& args);

    /**
     * @brief ROUND函数 - 四舍五入
     */
    static FormulaValue roundFunction(const std::vector<FormulaValue>& args);

    /**
     * @brief NOW函数 - 当前时间
     */
    static FormulaValue nowFunction(const std::vector<FormulaValue>& args);

    /**
     * @brief TODAY函数 - 今天日期
     */
    static FormulaValue todayFunction(const std::vector<FormulaValue>& args);

    // ==================== 工具函数 ====================

    /**
     * @brief 将FormulaValue转换为double
     */
    static double valueToNumber(const FormulaValue& value);

    /**
     * @brief 将FormulaValue转换为string
     */
    static std::string valueToString(const FormulaValue& value);

    /**
     * @brief 将FormulaValue转换为bool
     */
    static bool valueToBool(const FormulaValue& value);

    /**
     * @brief 从字符串创建FormulaValue
     */
    static FormulaValue valueFromString(const std::string& str);

    /**
     * @brief 检查两个值是否相等
     */
    static bool valuesEqual(const FormulaValue& a, const FormulaValue& b);

private:
    std::string formulaString_;
    FormulaError lastError_;
    std::vector<CellReference> dependencies_;
    std::unordered_map<std::string, FormulaFunction> customFunctions_;
    
    // Helper methods
    void registerBuiltinFunctions();
    FormulaValue evaluateExpression(const std::string& expr, const TXSheet* sheet, row_t currentRow, column_t currentCol);
    std::vector<FormulaValue> parseArguments(const std::string& args, const TXSheet* sheet, row_t currentRow, column_t currentCol);
    FormulaValue callFunction(const std::string& funcName, const std::vector<FormulaValue>& args);
    void updateDependencies();
    std::string replaceReferences(const std::string& expression, const TXSheet* sheet, row_t currentRow, column_t currentCol);
    FormulaValue evaluateSimpleExpression(const std::string& expr);
};

} // namespace TinaXlsx 
