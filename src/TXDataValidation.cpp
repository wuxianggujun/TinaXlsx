#include "TinaXlsx/TXDataValidation.hpp"
#include <sstream>

namespace TinaXlsx {

TXDataValidation::TXDataValidation()
    : type_(DataValidationType::None)
    , operator_(DataValidationOperator::Between)
    , showDropDown_(true)
    , errorStyle_(DataValidationErrorStyle::Stop)
    , errorTitle_("输入错误")
    , errorMessage_("输入的值不符合验证规则")
    , showErrorMessage_(true)
    , promptTitle_("输入提示")
    , promptMessage_("请输入有效的值")
    , showInputMessage_(false)
{
}

void TXDataValidation::setListItems(const std::vector<std::string>& items) {
    listItems_ = items;

    // 自动构建列表公式 - Excel需要用引号包围每个选项
    if (!items.empty()) {
        std::ostringstream oss;
        for (size_t i = 0; i < items.size(); ++i) {
            if (i > 0) oss << ",";
            // 为每个选项添加引号，并转义内部的引号
            oss << "\"";
            std::string item = items[i];
            // 转义引号：将 " 替换为 ""
            size_t pos = 0;
            while ((pos = item.find('"', pos)) != std::string::npos) {
                item.replace(pos, 1, "\"\"");
                pos += 2;
            }
            oss << item << "\"";
        }
        formula1_ = oss.str();
    }
}

// ==================== 便捷方法 ====================

TXDataValidation TXDataValidation::createIntegerValidation(int minValue, int maxValue) {
    TXDataValidation validation;
    validation.setType(DataValidationType::Whole);
    validation.setOperator(DataValidationOperator::Between);
    validation.setFormula1(std::to_string(minValue));
    validation.setFormula2(std::to_string(maxValue));
    validation.setErrorTitle("整数验证错误");
    validation.setErrorMessage("请输入 " + std::to_string(minValue) + " 到 " + std::to_string(maxValue) + " 之间的整数");
    validation.setPromptTitle("整数输入");
    validation.setPromptMessage("请输入 " + std::to_string(minValue) + " 到 " + std::to_string(maxValue) + " 之间的整数");
    validation.setShowInputMessage(true);
    return validation;
}

TXDataValidation TXDataValidation::createDecimalValidation(double minValue, double maxValue) {
    TXDataValidation validation;
    validation.setType(DataValidationType::Decimal);
    validation.setOperator(DataValidationOperator::Between);
    validation.setFormula1(std::to_string(minValue));
    validation.setFormula2(std::to_string(maxValue));
    validation.setErrorTitle("小数验证错误");
    validation.setErrorMessage("请输入 " + std::to_string(minValue) + " 到 " + std::to_string(maxValue) + " 之间的数值");
    validation.setPromptTitle("数值输入");
    validation.setPromptMessage("请输入 " + std::to_string(minValue) + " 到 " + std::to_string(maxValue) + " 之间的数值");
    validation.setShowInputMessage(true);
    return validation;
}

TXDataValidation TXDataValidation::createListValidation(const std::vector<std::string>& items, bool showDropDown) {
    TXDataValidation validation;
    validation.setType(DataValidationType::List);
    // 列表验证不需要设置操作符
    validation.setListItems(items);
    validation.setShowDropDown(showDropDown);
    validation.setErrorTitle("列表验证错误");
    validation.setErrorMessage("请从下拉列表中选择一个有效的选项");
    validation.setPromptTitle("选择选项");
    validation.setPromptMessage("请从下拉列表中选择一个选项");
    validation.setShowInputMessage(true);
    return validation;
}

TXDataValidation TXDataValidation::createListValidationFromRange(const TXRange& sourceRange, bool showDropDown) {
    TXDataValidation validation;
    validation.setType(DataValidationType::List);
    // 列表验证不需要设置操作符
    validation.setFormula1(sourceRange.toAbsoluteAddress());  // 使用绝对引用地址（推荐）
    validation.setShowDropDown(showDropDown);
    validation.setErrorTitle("列表验证错误");
    validation.setErrorMessage("请从下拉列表中选择一个有效的选项");
    validation.setPromptTitle("选择选项");
    validation.setPromptMessage("请从下拉列表中选择一个选项");
    validation.setShowInputMessage(true);
    return validation;
}

TXDataValidation TXDataValidation::createTextLengthValidation(int minLength, int maxLength) {
    TXDataValidation validation;
    validation.setType(DataValidationType::TextLength);
    validation.setOperator(DataValidationOperator::Between);
    validation.setFormula1(std::to_string(minLength));
    validation.setFormula2(std::to_string(maxLength));
    validation.setErrorTitle("文本长度验证错误");
    validation.setErrorMessage("文本长度必须在 " + std::to_string(minLength) + " 到 " + std::to_string(maxLength) + " 个字符之间");
    validation.setPromptTitle("文本长度");
    validation.setPromptMessage("请输入 " + std::to_string(minLength) + " 到 " + std::to_string(maxLength) + " 个字符的文本");
    validation.setShowInputMessage(true);
    return validation;
}

TXDataValidation TXDataValidation::createCustomValidation(const std::string& formula) {
    TXDataValidation validation;
    validation.setType(DataValidationType::Custom);
    validation.setFormula1(formula);
    validation.setErrorTitle("自定义验证错误");
    validation.setErrorMessage("输入的值不符合自定义验证规则");
    validation.setPromptTitle("自定义验证");
    validation.setPromptMessage("请输入符合验证规则的值");
    validation.setShowInputMessage(true);
    return validation;
}

} // namespace TinaXlsx
