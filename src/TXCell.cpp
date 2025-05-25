#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXFormula.hpp"
#include "TinaXlsx/TXNumberFormat.hpp"
#include <sstream>
#include <algorithm>
#include <cctype>
#include <iostream>

namespace TinaXlsx {

class TXCell::Impl {
public:
    Impl() : value_(std::monostate{}), type_(TXCell::CellType::Empty), 
             number_format_(TXCell::NumberFormat::General), 
             custom_format_(""), formula_(""),
             is_merged_(false), is_master_cell_(false),
             master_row_(0), master_col_(0) {}
    
    explicit Impl(const CellValue& value) : value_(value), 
             number_format_(TXCell::NumberFormat::General), 
             custom_format_(""), formula_(""),
             is_merged_(false), is_master_cell_(false),
             master_row_(0), master_col_(0) {
        updateType();
    }

    Impl(const Impl& other) : value_(other.value_), type_(other.type_),
             number_format_(other.number_format_), 
             custom_format_(other.custom_format_), formula_(other.formula_),
             is_merged_(other.is_merged_), is_master_cell_(other.is_master_cell_),
             master_row_(other.master_row_), master_col_(other.master_col_) {
        // 注意：TXFormula禁用了拷贝构造，所以暂时不拷贝公式对象
        // 只拷贝公式字符串，需要时重新解析
        if (other.formula_object_) {
            // TODO: 实现公式对象的深拷贝或序列化/反序列化
            // formula_object_ = std::make_unique<TXFormula>(*other.formula_object_);
        }
        
        // 深拷贝数字格式对象
        if (other.number_format_object_) {
            number_format_object_ = std::make_unique<TXNumberFormat>(*other.number_format_object_);
        }
    }

    Impl& operator=(const Impl& other) {
        if (this != &other) {
            value_ = other.value_;
            type_ = other.type_;
            number_format_ = other.number_format_;
            custom_format_ = other.custom_format_;
            formula_ = other.formula_;
            is_merged_ = other.is_merged_;
            is_master_cell_ = other.is_master_cell_;
            master_row_ = other.master_row_;
            master_col_ = other.master_col_;
            
            // 注意：TXFormula禁用了拷贝构造，所以暂时不拷贝公式对象
            // 只拷贝公式字符串，需要时重新解析
            if (other.formula_object_) {
                // TODO: 实现公式对象的深拷贝或序列化/反序列化
                // formula_object_ = std::make_unique<TXFormula>(*other.formula_object_);
                formula_object_.reset();
            } else {
                formula_object_.reset();
            }
            
            // 深拷贝数字格式对象
            if (other.number_format_object_) {
                number_format_object_ = std::make_unique<TXNumberFormat>(*other.number_format_object_);
            } else {
                number_format_object_.reset();
            }
        }
        return *this;
    }

    void setValue(const CellValue& value) {
        value_ = value;
        formula_.clear();
        formula_object_.reset();
        updateType();
    }

    const CellValue& getValue() const {
        return value_;
    }

    TXCell::CellType getType() const {
        return type_;
    }

    bool isEmpty() const {
        return std::holds_alternative<std::monostate>(value_);
    }

    void updateType() {
        if (std::holds_alternative<std::monostate>(value_)) {
            type_ = TXCell::CellType::Empty;
        } else if (std::holds_alternative<std::string>(value_)) {
            if (!formula_.empty() || formula_object_) {
                type_ = TXCell::CellType::Formula;
            } else {
                type_ = TXCell::CellType::String;
            }
        } else if (std::holds_alternative<double>(value_)) {
            type_ = TXCell::CellType::Number;
        } else if (std::holds_alternative<int64_t>(value_)) {
            type_ = TXCell::CellType::Integer;
        } else if (std::holds_alternative<bool>(value_)) {
            type_ = TXCell::CellType::Boolean;
        } else {
            type_ = TXCell::CellType::Empty;
        }
    }

    std::string getStringValue() const {
        if (std::holds_alternative<std::string>(value_)) {
            return std::get<std::string>(value_);
        } else if (std::holds_alternative<double>(value_)) {
            return std::to_string(std::get<double>(value_));
        } else if (std::holds_alternative<int64_t>(value_)) {
            return std::to_string(std::get<int64_t>(value_));
        } else if (std::holds_alternative<bool>(value_)) {
            return std::get<bool>(value_) ? "TRUE" : "FALSE";
        }
        return "";
    }

    double getNumberValue() const {
        if (std::holds_alternative<double>(value_)) {
            return std::get<double>(value_);
        } else if (std::holds_alternative<int64_t>(value_)) {
            return static_cast<double>(std::get<int64_t>(value_));
        } else if (std::holds_alternative<bool>(value_)) {
            return std::get<bool>(value_) ? 1.0 : 0.0;
        } else if (std::holds_alternative<std::string>(value_)) {
            try {
                return std::stod(std::get<std::string>(value_));
            } catch (...) {
                return 0.0;
            }
        }
        return 0.0;
    }

    int64_t getIntegerValue() const {
        if (std::holds_alternative<int64_t>(value_)) {
            return std::get<int64_t>(value_);
        } else if (std::holds_alternative<double>(value_)) {
            return static_cast<int64_t>(std::get<double>(value_));
        } else if (std::holds_alternative<bool>(value_)) {
            return std::get<bool>(value_) ? 1 : 0;
        } else if (std::holds_alternative<std::string>(value_)) {
            try {
                return std::stoll(std::get<std::string>(value_));
            } catch (...) {
                return 0;
            }
        }
        return 0;
    }

    bool getBooleanValue() const {
        if (std::holds_alternative<bool>(value_)) {
            return std::get<bool>(value_);
        } else if (std::holds_alternative<double>(value_)) {
            return std::get<double>(value_) != 0.0;
        } else if (std::holds_alternative<int64_t>(value_)) {
            return std::get<int64_t>(value_) != 0;
        } else if (std::holds_alternative<std::string>(value_)) {
            const auto& str = std::get<std::string>(value_);
            std::string lower_str = str;
            std::transform(lower_str.begin(), lower_str.end(), lower_str.begin(), ::tolower);
            return lower_str == "true" || lower_str == "1" || lower_str == "yes";
        }
        return false;
    }

    void setFormula(const std::string& formula) {
        formula_ = formula;
        if (!formula.empty()) {
            type_ = TXCell::CellType::Formula;
        }
    }

    const std::string& getFormula() const {
        return formula_;
    }

    bool isFormula() const {
        return !formula_.empty();
    }

    void clear() {
        value_ = std::monostate{};
        type_ = TXCell::CellType::Empty;
        formula_.clear();
        formula_object_.reset();
        custom_format_.clear();
        number_format_ = TXCell::NumberFormat::General;
        is_merged_ = false;
        is_master_cell_ = false;
        master_row_ = 0;
        master_col_ = 0;
    }

    bool fromString(const std::string& str, bool auto_detect_type) {
        if (!auto_detect_type) {
            setValue(str);
            return true;
        }

        // 自动检测类型
        if (str.empty()) {
            clear();
            return true;
        }

        // 检查是否为布尔值
        std::string lower_str = str;
        std::transform(lower_str.begin(), lower_str.end(), lower_str.begin(), ::tolower);
        if (lower_str == "true" || lower_str == "false") {
            setValue(lower_str == "true");
            return true;
        }

        // 检查是否为数字
        std::istringstream iss(str);
        double d;
        if (iss >> d && iss.eof()) {
            // 检查是否为整数
            if (str.find('.') == std::string::npos && 
                str.find('e') == std::string::npos && 
                str.find('E') == std::string::npos) {
                try {
                    int64_t i = std::stoll(str);
                    setValue(i);
                    return true;
                } catch (...) {
                    // 如果转换失败，作为double处理
                }
            }
            setValue(d);
            return true;
        }

        // 默认作为字符串处理
        setValue(str);
        return true;
    }

    // Getter和Setter方法
    TXCell::NumberFormat getNumberFormat() const {
        return number_format_;
    }

    void setNumberFormat(TXCell::NumberFormat format) {
        number_format_ = format;
    }

    std::string getCustomFormat() const {
        return custom_format_;
    }

    void setCustomFormat(const std::string& format) {
        custom_format_ = format;
    }

    // ==================== 新功能的访问方法 ====================
    
    const std::unique_ptr<TXFormula>& getFormulaObject() const {
        return formula_object_;
    }
    
    void setFormulaObject(std::unique_ptr<TXFormula> formula) {
        formula_object_ = std::move(formula);
        if (formula_object_) {
            formula_ = formula_object_->getFormulaString();
            type_ = TXCell::CellType::Formula;
        } else {
            formula_.clear();
            updateType();
        }
    }
    
    const std::unique_ptr<TXNumberFormat>& getNumberFormatObject() const {
        return number_format_object_;
    }
    
    void setNumberFormatObject(std::unique_ptr<TXNumberFormat> numberFormat) {
        number_format_object_ = std::move(numberFormat);
    }
    
    std::string getFormattedValue() const {
        if (number_format_object_) {
            // 使用新的TXNumberFormat对象
            if (std::holds_alternative<double>(value_)) {
                return number_format_object_->formatNumber(std::get<double>(value_));
            } else if (std::holds_alternative<int64_t>(value_)) {
                return number_format_object_->formatNumber(static_cast<double>(std::get<int64_t>(value_)));
            }
        }
        
        // 回退到基本格式化
        return getStringValue();
    }
    
    bool isMerged() const {
        return is_merged_;
    }
    
    void setMerged(bool merged) {
        is_merged_ = merged;
    }
    
    bool isMasterCell() const {
        return is_master_cell_;
    }
    
    void setMasterCell(bool master) {
        is_master_cell_ = master;
    }
    
    std::pair<uint32_t, uint32_t> getMasterCellPosition() const {
        return {master_row_, master_col_};
    }
    
    void setMasterCellPosition(uint32_t row, uint32_t col) {
        master_row_ = row;
        master_col_ = col;
    }
    
    void copyFormatTo(Impl& target) const {
        target.number_format_ = number_format_;
        target.custom_format_ = custom_format_;
        
        // 深拷贝格式对象
        if (number_format_object_) {
            target.number_format_object_ = std::make_unique<TXNumberFormat>(*number_format_object_);
        } else {
            target.number_format_object_.reset();
        }
    }

private:
    CellValue value_;
    TXCell::CellType type_;
    TXCell::NumberFormat number_format_;
    std::string custom_format_;
    std::string formula_;
    bool is_merged_;
    bool is_master_cell_;
    int master_row_;
    int master_col_;
    std::unique_ptr<TXFormula> formula_object_;
    std::unique_ptr<TXNumberFormat> number_format_object_;
};

// TXCell 实现
TXCell::TXCell() : pImpl(std::make_unique<Impl>()) {}

TXCell::TXCell(const CellValue& value) : pImpl(std::make_unique<Impl>(value)) {}

TXCell::~TXCell() = default;

TXCell::TXCell(const TXCell& other) : pImpl(std::make_unique<Impl>(*other.pImpl)) {}

TXCell& TXCell::operator=(const TXCell& other) {
    if (this != &other) {
        *pImpl = *other.pImpl;
    }
    return *this;
}

TXCell::TXCell(TXCell&& other) noexcept : pImpl(std::move(other.pImpl)) {}

TXCell& TXCell::operator=(TXCell&& other) noexcept {
    if (this != &other) {
        pImpl = std::move(other.pImpl);
    }
    return *this;
}

const TXCell::CellValue& TXCell::getValue() const {
    return pImpl->getValue();
}

void TXCell::setValue(const CellValue& value) {
    pImpl->setValue(value);
}

TXCell::CellType TXCell::getType() const {
    return pImpl->getType();
}

bool TXCell::isEmpty() const {
    return pImpl->isEmpty();
}

std::string TXCell::getStringValue() const {
    return pImpl->getStringValue();
}

double TXCell::getNumberValue() const {
    return pImpl->getNumberValue();
}

int64_t TXCell::getIntegerValue() const {
    return pImpl->getIntegerValue();
}

bool TXCell::getBooleanValue() const {
    return pImpl->getBooleanValue();
}

void TXCell::setStringValue(const std::string& value) {
    pImpl->setValue(value);
}

void TXCell::setNumberValue(double value) {
    pImpl->setValue(value);
}

void TXCell::setIntegerValue(int64_t value) {
    pImpl->setValue(value);
}

void TXCell::setBooleanValue(bool value) {
    pImpl->setValue(value);
}

std::string TXCell::getFormula() const {
    return pImpl->getFormula();
}

void TXCell::setFormula(const std::string& formula) {
    pImpl->setFormula(formula);
}

bool TXCell::isFormula() const {
    return pImpl->isFormula();
}

TXCell::NumberFormat TXCell::getNumberFormat() const {
    return pImpl->getNumberFormat();
}

void TXCell::setNumberFormat(NumberFormat format) {
    pImpl->setNumberFormat(format);
}

std::string TXCell::getCustomFormat() const {
    return pImpl->getCustomFormat();
}

void TXCell::setCustomFormat(const std::string& format_string) {
    pImpl->setCustomFormat(format_string);
}

void TXCell::clear() {
    pImpl->clear();
}

std::string TXCell::toString() const {
    return pImpl->getStringValue();
}

bool TXCell::fromString(const std::string& str, bool auto_detect_type) {
    return pImpl->fromString(str, auto_detect_type);
}

TXCell& TXCell::operator=(const std::string& value) {
    setStringValue(value);
    return *this;
}

TXCell& TXCell::operator=(const char* value) {
    setStringValue(std::string(value));
    return *this;
}

TXCell& TXCell::operator=(double value) {
    setNumberValue(value);
    return *this;
}

TXCell& TXCell::operator=(int64_t value) {
    setIntegerValue(value);
    return *this;
}

TXCell& TXCell::operator=(int value) {
    setIntegerValue(static_cast<int64_t>(value));
    return *this;
}

TXCell& TXCell::operator=(bool value) {
    setBooleanValue(value);
    return *this;
}

bool TXCell::operator==(const TXCell& other) const {
    return pImpl->getValue() == other.pImpl->getValue();
}

bool TXCell::operator!=(const TXCell& other) const {
    return !(*this == other);
}

// ==================== 公式功能实现 ====================

const TXFormula* TXCell::getFormulaObject() const {
    return pImpl->getFormulaObject().get();
}

void TXCell::setFormulaObject(std::unique_ptr<TXFormula> formula) {
    pImpl->setFormulaObject(std::move(formula));
}

TXCell::CellValue TXCell::evaluateFormula(const TXSheet* sheet, uint32_t currentRow, uint32_t currentCol) {
    const auto& formulaObj = pImpl->getFormulaObject();
    if (!formulaObj) {
        return std::monostate{};
    }
    
    try {
        auto result = formulaObj->evaluate(sheet, currentRow, currentCol);
        
        // 将FormulaValue转换为CellValue
        if (std::holds_alternative<double>(result)) {
            return std::get<double>(result);
        } else if (std::holds_alternative<std::string>(result)) {
            return std::get<std::string>(result);
        } else if (std::holds_alternative<int64_t>(result)) {
            return std::get<int64_t>(result);
        } else if (std::holds_alternative<bool>(result)) {
            return std::get<bool>(result);
        } else if (std::holds_alternative<std::monostate>(result)) {
            // 错误或空值情况
            return std::monostate{};
        }
    } catch (const std::exception& e) {
        return std::string("#ERROR: ") + e.what();
    }
    
    return std::monostate{};
}

// ==================== 数字格式化功能实现 ====================

const TXNumberFormat* TXCell::getNumberFormatObject() const {
    return pImpl->getNumberFormatObject().get();
}

void TXCell::setNumberFormatObject(std::unique_ptr<TXNumberFormat> numberFormat) {
    pImpl->setNumberFormatObject(std::move(numberFormat));
}

std::string TXCell::getFormattedValue() const {
    return pImpl->getFormattedValue();
}

void TXCell::setPredefinedFormat(NumberFormat type, int decimalPlaces, bool useThousandSeparator) {
    // 创建对应的TXNumberFormat对象
    auto formatObject = std::make_unique<TXNumberFormat>();
    
    TXNumberFormat::FormatType formatType;
    switch (type) {
        case NumberFormat::Number:
            formatType = TXNumberFormat::FormatType::Number;
            break;
        case NumberFormat::Currency:
            formatType = TXNumberFormat::FormatType::Currency;
            break;
        case NumberFormat::Percentage:
            formatType = TXNumberFormat::FormatType::Percentage;
            break;
        case NumberFormat::Date:
            formatType = TXNumberFormat::FormatType::Date;
            break;
        case NumberFormat::Time:
            formatType = TXNumberFormat::FormatType::Time;
            break;
        case NumberFormat::Scientific:
            formatType = TXNumberFormat::FormatType::Scientific;
            break;
        default:
            formatType = TXNumberFormat::FormatType::General;
            break;
    }
    
    TXNumberFormat::FormatOptions options;
    options.decimalPlaces = decimalPlaces;
    options.useThousandSeparator = useThousandSeparator;
    
    formatObject->setFormat(formatType, options);
    setNumberFormatObject(std::move(formatObject));
    
    // 同时设置兼容性格式
    pImpl->setNumberFormat(type);
}

// ==================== 合并单元格功能实现 ====================

bool TXCell::isMerged() const {
    return pImpl->isMerged();
}

void TXCell::setMerged(bool merged) {
    pImpl->setMerged(merged);
}

bool TXCell::isMasterCell() const {
    return pImpl->isMasterCell();
}

void TXCell::setMasterCell(bool master) {
    pImpl->setMasterCell(master);
}

std::pair<uint32_t, uint32_t> TXCell::getMasterCellPosition() const {
    return pImpl->getMasterCellPosition();
}

void TXCell::setMasterCellPosition(uint32_t row, uint32_t col) {
    pImpl->setMasterCellPosition(row, col);
}

// ==================== 其他功能实现 ====================

std::unique_ptr<TXCell> TXCell::clone() const {
    return std::make_unique<TXCell>(*this);
}

void TXCell::copyFormatTo(TXCell& target) const {
    pImpl->copyFormatTo(*target.pImpl);
}

bool TXCell::isValueEqual(const TXCell& other) const {
    return pImpl->getValue() == other.pImpl->getValue();
}

} // namespace TinaXlsx 