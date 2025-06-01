#pragma once

#include "TXTypes.hpp"
#include "TXRange.hpp"
#include <string>
#include <vector>

namespace TinaXlsx {

/**
 * @brief 公式构建器
 *
 * 提供便捷的静态方法来构建Excel公式字符串
 * 作为TXFormula和TXFormulaManager的补充工具
 * 专注于公式字符串的生成，不涉及计算逻辑
 */
class TXFormulaBuilder {
public:
    // ==================== 统计函数 ====================

    /**
     * @brief 构建SUM公式
     * @param range 求和范围
     * @return 公式字符串
     */
    static std::string sum(const TXRange& range);

    /**
     * @brief 构建SUM公式
     * @param rangeAddress 范围地址字符串
     * @return 公式字符串
     */
    static std::string sum(const std::string& rangeAddress);

    /**
     * @brief 构建AVERAGE公式
     * @param range 平均值范围
     * @return 公式字符串
     */
    static std::string average(const TXRange& range);

    /**
     * @brief 构建COUNT公式
     * @param range 计数范围
     * @return 公式字符串
     */
    static std::string count(const TXRange& range);

    /**
     * @brief 构建COUNTA公式（计算非空单元格）
     * @param range 计数范围
     * @return 公式字符串
     */
    static std::string countA(const TXRange& range);

    /**
     * @brief 构建MAX公式
     * @param range 最大值范围
     * @return 公式字符串
     */
    static std::string max(const TXRange& range);

    /**
     * @brief 构建MIN公式
     * @param range 最小值范围
     * @return 公式字符串
     */
    static std::string min(const TXRange& range);

    /**
     * @brief 构建STDEV公式（标准差）
     * @param range 标准差范围
     * @return 公式字符串
     */
    static std::string stdev(const TXRange& range);

    /**
     * @brief 构建VAR公式（方差）
     * @param range 方差范围
     * @return 公式字符串
     */
    static std::string var(const TXRange& range);

    // ==================== 条件函数 ====================

    /**
     * @brief 构建SUMIF公式
     * @param range 条件范围
     * @param criteria 条件
     * @param sumRange 求和范围（可选）
     * @return 公式字符串
     */
    static std::string sumIf(const TXRange& range, const std::string& criteria, const TXRange& sumRange = TXRange());

    /**
     * @brief 构建COUNTIF公式
     * @param range 条件范围
     * @param criteria 条件
     * @return 公式字符串
     */
    static std::string countIf(const TXRange& range, const std::string& criteria);

    /**
     * @brief 构建AVERAGEIF公式
     * @param range 条件范围
     * @param criteria 条件
     * @param averageRange 平均值范围（可选）
     * @return 公式字符串
     */
    static std::string averageIf(const TXRange& range, const std::string& criteria, const TXRange& averageRange = TXRange());

    /**
     * @brief 构建IF公式
     * @param condition 条件
     * @param valueIfTrue 条件为真时的值
     * @param valueIfFalse 条件为假时的值
     * @return 公式字符串
     */
    static std::string ifFormula(const std::string& condition, const std::string& valueIfTrue, const std::string& valueIfFalse);

    // ==================== 查找函数 ====================

    /**
     * @brief 构建VLOOKUP公式
     * @param lookupValue 查找值
     * @param tableArray 查找表
     * @param colIndexNum 列索引
     * @param rangeLookup 是否近似匹配
     * @return 公式字符串
     */
    static std::string vlookup(const std::string& lookupValue, const TXRange& tableArray, int colIndexNum, bool rangeLookup = false);

    /**
     * @brief 构建HLOOKUP公式
     * @param lookupValue 查找值
     * @param tableArray 查找表
     * @param rowIndexNum 行索引
     * @param rangeLookup 是否近似匹配
     * @return 公式字符串
     */
    static std::string hlookup(const std::string& lookupValue, const TXRange& tableArray, int rowIndexNum, bool rangeLookup = false);

    /**
     * @brief 构建INDEX公式
     * @param array 数组
     * @param rowNum 行号
     * @param colNum 列号（可选）
     * @return 公式字符串
     */
    static std::string index(const TXRange& array, int rowNum, int colNum = 0);

    /**
     * @brief 构建MATCH公式
     * @param lookupValue 查找值
     * @param lookupArray 查找数组
     * @param matchType 匹配类型
     * @return 公式字符串
     */
    static std::string match(const std::string& lookupValue, const TXRange& lookupArray, int matchType = 0);

    // ==================== 文本函数 ====================

    /**
     * @brief 构建CONCATENATE公式
     * @param values 要连接的值列表
     * @return 公式字符串
     */
    static std::string concatenate(const std::vector<std::string>& values);

    /**
     * @brief 构建LEFT公式
     * @param text 文本
     * @param numChars 字符数
     * @return 公式字符串
     */
    static std::string left(const std::string& text, int numChars);

    /**
     * @brief 构建RIGHT公式
     * @param text 文本
     * @param numChars 字符数
     * @return 公式字符串
     */
    static std::string right(const std::string& text, int numChars);

    /**
     * @brief 构建MID公式
     * @param text 文本
     * @param startNum 起始位置
     * @param numChars 字符数
     * @return 公式字符串
     */
    static std::string mid(const std::string& text, int startNum, int numChars);

    /**
     * @brief 构建LEN公式
     * @param text 文本
     * @return 公式字符串
     */
    static std::string len(const std::string& text);

    /**
     * @brief 构建UPPER公式
     * @param text 文本
     * @return 公式字符串
     */
    static std::string upper(const std::string& text);

    /**
     * @brief 构建LOWER公式
     * @param text 文本
     * @return 公式字符串
     */
    static std::string lower(const std::string& text);

    // ==================== 日期时间函数 ====================

    /**
     * @brief 构建TODAY公式
     * @return 公式字符串
     */
    static std::string today();

    /**
     * @brief 构建NOW公式
     * @return 公式字符串
     */
    static std::string now();

    /**
     * @brief 构建DATE公式
     * @param year 年
     * @param month 月
     * @param day 日
     * @return 公式字符串
     */
    static std::string date(int year, int month, int day);

    /**
     * @brief 构建YEAR公式
     * @param dateValue 日期值
     * @return 公式字符串
     */
    static std::string year(const std::string& dateValue);

    /**
     * @brief 构建MONTH公式
     * @param dateValue 日期值
     * @return 公式字符串
     */
    static std::string month(const std::string& dateValue);

    /**
     * @brief 构建DAY公式
     * @param dateValue 日期值
     * @return 公式字符串
     */
    static std::string day(const std::string& dateValue);

    // ==================== 数学函数 ====================

    /**
     * @brief 构建ROUND公式
     * @param number 数值
     * @param numDigits 小数位数
     * @return 公式字符串
     */
    static std::string round(const std::string& number, int numDigits);

    /**
     * @brief 构建ABS公式
     * @param number 数值
     * @return 公式字符串
     */
    static std::string abs(const std::string& number);

    /**
     * @brief 构建POWER公式
     * @param number 底数
     * @param power 指数
     * @return 公式字符串
     */
    static std::string power(const std::string& number, const std::string& power);

    /**
     * @brief 构建SQRT公式
     * @param number 数值
     * @return 公式字符串
     */
    static std::string sqrt(const std::string& number);

private:
    /**
     * @brief 格式化范围为公式字符串
     * @param range 范围
     * @return 范围字符串
     */
    static std::string formatRange(const TXRange& range);

    /**
     * @brief 转义文本值
     * @param text 文本
     * @return 转义后的文本
     */
    static std::string escapeText(const std::string& text);
};

} // namespace TinaXlsx
