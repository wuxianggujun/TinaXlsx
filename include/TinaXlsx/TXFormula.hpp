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
    using FormulaValue = std::variant<std::monostate, std::string, double, int64_t, bool>;
    
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
        TXTypes::RowIndex row;
        TXTypes::ColIndex col;
        bool absoluteRow = false;    ///< 绝对行引用（$A1）
        bool absoluteCol = false;    ///< 绝对列引用（A$1）
        std::string sheetName;       ///< 工作表名称（跨表引用）
        
        CellReference() : row(0), col(0) {}
        CellReference(TXTypes::RowIndex r, TXTypes::ColIndex c) : row(r), col(c) {}
        
        std::string toString() const;
        static CellReference fromString(const std::string& ref);
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
    };

public:
    TXFormula();
    explicit TXFormula(const std::string& formula);
    ~TXFormula();
    
    // 禁用拷贝，支持移动
    TXFormula(const TXFormula&) = delete;
    TXFormula& operator=(const TXFormula&) = delete;
    TXFormula(TXFormula&&) noexcept;
    TXFormula& operator=(TXFormula&&) noexcept;

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
    FormulaValue evaluate(const TXSheet* sheet, TXTypes::RowIndex currentRow, TXTypes::ColIndex currentCol);

    /**
     * @brief 获取公式字符串
     * @return 公式字符串
     */
    const std::string& getFormulaString() const;

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

private:
    class Impl;
    std::unique_ptr<Impl> pImpl;
};

} // namespace TinaXlsx 
