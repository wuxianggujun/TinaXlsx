#pragma once

#include "TXTypes.hpp"
#include "TXRange.hpp"
#include <string>
#include <vector>

namespace TinaXlsx {

/**
 * @brief 数据验证类型
 */
enum class DataValidationType {
    None,           ///< 无验证
    Whole,          ///< 整数
    Decimal,        ///< 小数
    List,           ///< 列表
    Date,           ///< 日期
    Time,           ///< 时间
    TextLength,     ///< 文本长度
    Custom          ///< 自定义公式
};

/**
 * @brief 数据验证操作符
 */
enum class DataValidationOperator {
    Between,        ///< 介于
    NotBetween,     ///< 不介于
    Equal,          ///< 等于
    NotEqual,       ///< 不等于
    GreaterThan,    ///< 大于
    LessThan,       ///< 小于
    GreaterThanOrEqual, ///< 大于等于
    LessThanOrEqual     ///< 小于等于
};

/**
 * @brief 错误样式
 */
enum class DataValidationErrorStyle {
    Stop,           ///< 停止
    Warning,        ///< 警告
    Information     ///< 信息
};

/**
 * @brief 数据验证规则
 */
class TXDataValidation {
public:
    /**
     * @brief 默认构造函数
     */
    TXDataValidation();

    // ==================== 基本属性 ====================

    /**
     * @brief 设置验证类型
     * @param type 验证类型
     */
    void setType(DataValidationType type) { type_ = type; }

    /**
     * @brief 获取验证类型
     */
    DataValidationType getType() const { return type_; }

    /**
     * @brief 设置操作符
     * @param op 操作符
     */
    void setOperator(DataValidationOperator op) { operator_ = op; }

    /**
     * @brief 获取操作符
     */
    DataValidationOperator getOperator() const { return operator_; }

    // ==================== 验证值 ====================

    /**
     * @brief 设置第一个验证值
     * @param formula 公式或值
     */
    void setFormula1(const std::string& formula) { formula1_ = formula; }

    /**
     * @brief 获取第一个验证值
     */
    const std::string& getFormula1() const { return formula1_; }

    /**
     * @brief 设置第二个验证值（用于Between等操作）
     * @param formula 公式或值
     */
    void setFormula2(const std::string& formula) { formula2_ = formula; }

    /**
     * @brief 获取第二个验证值
     */
    const std::string& getFormula2() const { return formula2_; }

    // ==================== 下拉列表 ====================

    /**
     * @brief 设置是否显示下拉箭头
     * @param show 是否显示
     */
    void setShowDropDown(bool show) { showDropDown_ = show; }

    /**
     * @brief 获取是否显示下拉箭头
     */
    bool getShowDropDown() const { return showDropDown_; }

    /**
     * @brief 设置列表项（用于List类型）
     * @param items 列表项
     */
    void setListItems(const std::vector<std::string>& items);

    /**
     * @brief 获取列表项
     */
    const std::vector<std::string>& getListItems() const { return listItems_; }

    // ==================== 错误处理 ====================

    /**
     * @brief 设置错误样式
     * @param style 错误样式
     */
    void setErrorStyle(DataValidationErrorStyle style) { errorStyle_ = style; }

    /**
     * @brief 获取错误样式
     */
    DataValidationErrorStyle getErrorStyle() const { return errorStyle_; }

    /**
     * @brief 设置错误标题
     * @param title 错误标题
     */
    void setErrorTitle(const std::string& title) { errorTitle_ = title; }

    /**
     * @brief 获取错误标题
     */
    const std::string& getErrorTitle() const { return errorTitle_; }

    /**
     * @brief 设置错误消息
     * @param message 错误消息
     */
    void setErrorMessage(const std::string& message) { errorMessage_ = message; }

    /**
     * @brief 获取错误消息
     */
    const std::string& getErrorMessage() const { return errorMessage_; }

    /**
     * @brief 设置是否显示错误警告
     * @param show 是否显示
     */
    void setShowErrorMessage(bool show) { showErrorMessage_ = show; }

    /**
     * @brief 获取是否显示错误警告
     */
    bool getShowErrorMessage() const { return showErrorMessage_; }

    // ==================== 输入提示 ====================

    /**
     * @brief 设置输入提示标题
     * @param title 提示标题
     */
    void setPromptTitle(const std::string& title) { promptTitle_ = title; }

    /**
     * @brief 获取输入提示标题
     */
    const std::string& getPromptTitle() const { return promptTitle_; }

    /**
     * @brief 设置输入提示消息
     * @param message 提示消息
     */
    void setPromptMessage(const std::string& message) { promptMessage_ = message; }

    /**
     * @brief 获取输入提示消息
     */
    const std::string& getPromptMessage() const { return promptMessage_; }

    /**
     * @brief 设置是否显示输入提示
     * @param show 是否显示
     */
    void setShowInputMessage(bool show) { showInputMessage_ = show; }

    /**
     * @brief 获取是否显示输入提示
     */
    bool getShowInputMessage() const { return showInputMessage_; }

    // ==================== 便捷方法 ====================

    /**
     * @brief 创建整数验证
     * @param minValue 最小值
     * @param maxValue 最大值
     * @return 数据验证对象
     */
    static TXDataValidation createIntegerValidation(int minValue, int maxValue);

    /**
     * @brief 创建小数验证
     * @param minValue 最小值
     * @param maxValue 最大值
     * @return 数据验证对象
     */
    static TXDataValidation createDecimalValidation(double minValue, double maxValue);

    /**
     * @brief 创建列表验证（直接列表）
     * @param items 列表项
     * @param showDropDown 是否显示下拉箭头
     * @return 数据验证对象
     */
    static TXDataValidation createListValidation(const std::vector<std::string>& items, bool showDropDown = true);

    /**
     * @brief 创建列表验证（范围引用）
     * @param sourceRange 包含选项的单元格范围
     * @param showDropDown 是否显示下拉箭头
     * @return 数据验证对象
     */
    static TXDataValidation createListValidationFromRange(const TXRange& sourceRange, bool showDropDown = true);

    /**
     * @brief 创建文本长度验证
     * @param minLength 最小长度
     * @param maxLength 最大长度
     * @return 数据验证对象
     */
    static TXDataValidation createTextLengthValidation(int minLength, int maxLength);

    /**
     * @brief 创建自定义公式验证
     * @param formula 验证公式
     * @return 数据验证对象
     */
    static TXDataValidation createCustomValidation(const std::string& formula);

private:
    DataValidationType type_;               ///< 验证类型
    DataValidationOperator operator_;       ///< 操作符
    
    std::string formula1_;                  ///< 第一个验证值
    std::string formula2_;                  ///< 第二个验证值
    
    bool showDropDown_;                     ///< 是否显示下拉箭头
    std::vector<std::string> listItems_;   ///< 列表项
    
    DataValidationErrorStyle errorStyle_;  ///< 错误样式
    std::string errorTitle_;                ///< 错误标题
    std::string errorMessage_;              ///< 错误消息
    bool showErrorMessage_;                 ///< 是否显示错误警告
    
    std::string promptTitle_;               ///< 输入提示标题
    std::string promptMessage_;             ///< 输入提示消息
    bool showInputMessage_;                 ///< 是否显示输入提示
};

} // namespace TinaXlsx
