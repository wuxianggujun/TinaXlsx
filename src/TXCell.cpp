#include "TinaXlsx/TXCell.hpp"
#include <sstream>
#include <algorithm>
#include <cctype>

namespace TinaXlsx {

class TXCell::Impl {
public:
    Impl() : value_(""), type_(TXCell::CellType::Empty), 
             number_format_(TXCell::NumberFormat::General), 
             custom_format_(""), formula_("") {}
    
    explicit Impl(const CellValue& value) : value_(value), 
             number_format_(TXCell::NumberFormat::General), 
             custom_format_(""), formula_("") {
        updateType();
    }

    void setValue(const CellValue& value) {
        value_ = value;
        formula_.clear();
        updateType();
    }

    const CellValue& getValue() const {
        return value_;
    }

    TXCell::CellType getType() const {
        return type_;
    }

    bool isEmpty() const {
        return type_ == TXCell::CellType::Empty;
    }

    void updateType() {
        if (std::holds_alternative<std::string>(value_)) {
            const auto& str = std::get<std::string>(value_);
            if (str.empty()) {
                type_ = TXCell::CellType::Empty;
            } else if (!formula_.empty()) {
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
        value_ = std::string("");
        type_ = TXCell::CellType::Empty;
        formula_.clear();
        custom_format_.clear();
        number_format_ = TXCell::NumberFormat::General;
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

private:
    CellValue value_;
    TXCell::CellType type_;
    TXCell::NumberFormat number_format_;
    std::string custom_format_;
    std::string formula_;
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

} // namespace TinaXlsx 