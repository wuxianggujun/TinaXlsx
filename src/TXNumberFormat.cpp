#include "TinaXlsx/TXNumberFormat.hpp"
#include <sstream>
#include <iomanip>
#include <cmath>
#include <chrono>
#include <regex>
#include <locale>
#include <ctime>

namespace TinaXlsx {

// ==================== TXNumberFormat::Impl实现 ====================

class TXNumberFormat::Impl {
public:
    FormatType formatType_ = FormatType::General;
    FormatOptions options_;
    std::string customFormatString_;
    
    // 预编译的格式模式
    std::regex numberPattern_;
    std::regex datePattern_;
    std::regex timePattern_;
    
    Impl() {
        updatePatterns();
    }
    
    explicit Impl(FormatType type, const FormatOptions& options) 
        : formatType_(type), options_(options) {
        updatePatterns();
    }
    
    explicit Impl(const std::string& customFormat) 
        : formatType_(FormatType::Custom), customFormatString_(customFormat) {
        updatePatterns();
    }
    
    void updatePatterns() {
        // 根据格式类型更新正则表达式模式
        switch (formatType_) {
            case FormatType::Number:
            case FormatType::Decimal:
                numberPattern_ = std::regex(R"(^-?\d{1,3}(?:,\d{3})*(?:\.\d+)?$)");
                break;
            case FormatType::Currency:
                numberPattern_ = std::regex(R"(^[^\d]*-?\d{1,3}(?:,\d{3})*(?:\.\d+)?[^\d]*$)");
                break;
            case FormatType::Percentage:
                numberPattern_ = std::regex(R"(^-?\d+(?:\.\d+)?%$)");
                break;
            case FormatType::Date:
                datePattern_ = std::regex(R"(^\d{4}-\d{2}-\d{2}$)");
                break;
            case FormatType::Time:
                timePattern_ = std::regex(R"(^\d{2}:\d{2}:\d{2}$)");
                break;
            default:
                break;
        }
    }
    
    std::string formatValue(const Value& value) const {
        if (std::holds_alternative<std::monostate>(value)) {
            return options_.showZero ? "0" : "";
        }
        
        switch (formatType_) {
            case FormatType::General:
                return formatGeneral(value);
            case FormatType::Number:
            case FormatType::Decimal:
                return formatNumber(valueToNumber(value));
            case FormatType::Currency:
                return formatCurrency(valueToNumber(value));
            case FormatType::Percentage:
                return formatPercentage(valueToNumber(value));
            case FormatType::Date:
                return formatDate(valueToNumber(value));
            case FormatType::Time:
                return formatTime(valueToNumber(value));
            case FormatType::DateTime:
                return formatDateTime(valueToNumber(value));
            case FormatType::Scientific:
                return formatScientific(valueToNumber(value));
            case FormatType::Text:
                return formatText(value);
            case FormatType::Custom:
                return formatCustom(value);
            default:
                return formatGeneral(value);
        }
    }
    
    std::string formatGeneral(const Value& value) const {
        if (std::holds_alternative<std::string>(value)) {
            return std::get<std::string>(value);
        } else if (std::holds_alternative<double>(value)) {
            double d = std::get<double>(value);
            // 智能数字格式化
            if (std::abs(d) >= 1e6 || (std::abs(d) < 1e-3 && d != 0.0)) {
                return formatScientific(d);
            } else {
                return formatNumber(d);
            }
        } else if (std::holds_alternative<int64_t>(value)) {
            return std::to_string(std::get<int64_t>(value));
        } else if (std::holds_alternative<bool>(value)) {
            return std::get<bool>(value) ? "TRUE" : "FALSE";
        }
        return "";
    }
    
    std::string formatNumber(double number) const {
        std::ostringstream oss;
        
        // 处理负数
        bool isNegative = number < 0;
        if (isNegative) {
            number = -number;
        }
        
        // 设置小数位数
        oss << std::fixed << std::setprecision(options_.decimalPlaces);
        
        if (options_.useThousandSeparator) {
            // 先格式化数字
            oss << number;
            std::string numStr = oss.str();
            
            // 找到小数点位置
            size_t dotPos = numStr.find('.');
            std::string intPart = numStr.substr(0, dotPos);
            std::string fracPart = (dotPos != std::string::npos) ? numStr.substr(dotPos) : "";
            
            // 添加千位分隔符到整数部分
            std::string result;
            int count = 0;
            for (int i = static_cast<int>(intPart.length()) - 1; i >= 0; --i) {
                if (count > 0 && count % 3 == 0) {
                    result = "," + result;
                }
                result = intPart[i] + result;
                count++;
            }
            
            // 组合结果
            if (isNegative) result = "-" + result;
            if (!fracPart.empty()) result += fracPart;
            
            return result;
        } else {
            if (isNegative) oss.str("-");
            oss << number;
            return oss.str();
        }
    }
    
    std::string formatCurrency(double amount) const {
        std::string numStr = formatNumber(std::abs(amount));
        
        if (amount < 0) {
            // 移除负号，因为formatNumber已经处理了
            if (numStr[0] == '-') {
                numStr = numStr.substr(1);
            }
            return "(" + options_.currencySymbol + numStr + ")";
        } else {
            return options_.currencySymbol + numStr;
        }
    }
    
    std::string formatPercentage(double value) const {
        double percentage = value * 100.0;
        std::ostringstream oss;
        oss << std::fixed << std::setprecision(options_.decimalPlaces) << percentage << "%";
        return oss.str();
    }
    
    std::string formatDate(double excelDate) const {
        time_t timeT = excelDateToSystemTime(excelDate);
        
        // 使用安全的localtime函数
        struct tm tmBuffer = {};
        struct tm* tmPtr = nullptr;
#ifdef _WIN32
        if (localtime_s(&tmBuffer, &timeT) == 0) {
            tmPtr = &tmBuffer;
        }
#else
        tmPtr = localtime_r(&timeT, &tmBuffer);
#endif
        if (!tmPtr) {
            return ""; // 错误处理
        }
        
        std::string format = options_.dateFormat;
        // 简化的日期格式转换
        std::regex yyyy(R"(yyyy)");
        std::regex mm(R"(mm)");
        std::regex dd(R"(dd)");
        
        format = std::regex_replace(format, yyyy, "%Y");
        format = std::regex_replace(format, mm, "%m");
        format = std::regex_replace(format, dd, "%d");
        
        char buffer[32];
        strftime(buffer, sizeof(buffer), format.c_str(), tmPtr);
        return std::string(buffer);
    }
    
    std::string formatTime(double excelTime) const {
        // Excel时间是一天的小数部分
        double dayFraction = excelTime - std::floor(excelTime);
        int totalSeconds = static_cast<int>(dayFraction * 86400.0);
        
        int hours = totalSeconds / 3600;
        int minutes = (totalSeconds % 3600) / 60;
        int seconds = totalSeconds % 60;
        
        std::ostringstream oss;
        oss << std::setfill('0') << std::setw(2) << hours << ":"
            << std::setw(2) << minutes << ":"
            << std::setw(2) << seconds;
        
        return oss.str();
    }
    
    std::string formatDateTime(double excelDateTime) const {
        return formatDate(excelDateTime) + " " + formatTime(excelDateTime);
    }
    
    std::string formatScientific(double number) const {
        std::ostringstream oss;
        oss << std::scientific << std::setprecision(options_.decimalPlaces) << number;
        return oss.str();
    }
    
    std::string formatText(const Value& value) const {
        if (std::holds_alternative<std::string>(value)) {
            return std::get<std::string>(value);
        } else {
            return formatGeneral(value);
        }
    }
    
    std::string formatCustom(const Value& value) const {
        // 简化的自定义格式处理
        if (customFormatString_.empty()) {
            return formatGeneral(value);
        }
        
        // 处理常见的自定义格式
        std::string result = customFormatString_;
        double numValue = valueToNumber(value);
        
        // 替换占位符
        std::regex zeroPattern(R"(0+)");
        std::regex hashPattern(R"(#+)");
        
        if (std::regex_search(result, zeroPattern) || std::regex_search(result, hashPattern)) {
            std::ostringstream oss;
            oss << std::fixed << std::setprecision(2) << numValue;
            result = std::regex_replace(result, std::regex(R"([0#]+)"), oss.str());
        }
        
        return result;
    }
    
    Value parseValue(const std::string& formattedStr) const {
        if (formattedStr.empty()) {
            return std::monostate{};
        }
        
        switch (formatType_) {
            case FormatType::Number:
            case FormatType::Decimal:
                return parseNumber(formattedStr);
            case FormatType::Currency:
                return parseCurrency(formattedStr);
            case FormatType::Percentage:
                return parsePercentage(formattedStr);
            case FormatType::Date:
                return parseDate(formattedStr);
            case FormatType::Time:
                return parseTime(formattedStr);
            default:
                return parseGeneral(formattedStr);
        }
    }
    
    Value parseGeneral(const std::string& str) const {
        // 尝试解析为不同类型
        try {
            if (str == "TRUE" || str == "true") return true;
            if (str == "FALSE" || str == "false") return false;
            
            // 尝试解析为整数
            if (str.find('.') == std::string::npos) {
                return std::stoll(str);
            }
            
            // 尝试解析为浮点数
            return std::stod(str);
        } catch (...) {
            return str;
        }
    }
    
    Value parseNumber(const std::string& str) const {
        try {
            std::string cleaned = str;
            // 移除千位分隔符
            cleaned.erase(std::remove(cleaned.begin(), cleaned.end(), ','), cleaned.end());
            return std::stod(cleaned);
        } catch (...) {
            return str;
        }
    }
    
    Value parseCurrency(const std::string& str) const {
        try {
            std::string cleaned = str;
            // 移除货币符号
            size_t symbolPos = cleaned.find(options_.currencySymbol);
            if (symbolPos != std::string::npos) {
                cleaned.erase(symbolPos, options_.currencySymbol.length());
            }
            
            // 处理括号表示的负数
            bool isNegative = false;
            if (cleaned.front() == '(' && cleaned.back() == ')') {
                isNegative = true;
                cleaned = cleaned.substr(1, cleaned.length() - 2);
            }
            
            // 移除千位分隔符
            cleaned.erase(std::remove(cleaned.begin(), cleaned.end(), ','), cleaned.end());
            
            double value = std::stod(cleaned);
            return isNegative ? -value : value;
        } catch (...) {
            return str;
        }
    }
    
    Value parsePercentage(const std::string& str) const {
        try {
            if (str.back() == '%') {
                std::string numStr = str.substr(0, str.length() - 1);
                double percentage = std::stod(numStr);
                return percentage / 100.0;
            }
        } catch (...) {
        }
        return str;
    }
    
    Value parseDate(const std::string& str) const {
        // 简化的日期解析
        try {
            return parseDateString(str, options_.dateFormat);
        } catch (...) {
            return str;
        }
    }
    
    Value parseTime(const std::string& str) const {
        try {
            std::regex timeRegex(R"((\d{1,2}):(\d{2}):(\d{2}))");
            std::smatch match;
            if (std::regex_match(str, match, timeRegex)) {
                int hours = std::stoi(match[1].str());
                int minutes = std::stoi(match[2].str());
                int seconds = std::stoi(match[3].str());
                
                double timeValue = (hours * 3600.0 + minutes * 60.0 + seconds) / 86400.0;
                return timeValue;
            }
        } catch (...) {
        }
        return str;
    }
    
    static double valueToNumber(const Value& value) {
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
    
    static time_t excelDateToSystemTime(double excelDate) {
        // Excel日期系统：1900年1月1日为1（实际上Excel错误地认为1900年是闰年）
        const int secondsPerDay = 86400;
        
        // Excel日期25569对应Unix纪元（1970年1月1日）
        double daysSinceUnixEpoch = excelDate - 25569.0;
        return static_cast<time_t>(daysSinceUnixEpoch * secondsPerDay);
    }
    
    static double systemTimeToExcelDate(time_t timeT) {
        const int secondsPerDay = 86400;
        double daysSinceUnixEpoch = static_cast<double>(timeT) / secondsPerDay;
        return daysSinceUnixEpoch + 25569.0;
    }
    
    static double parseDateString(const std::string& dateStr, const std::string& format) {
        (void)format;  // 标记参数为有意未使用
        // 简化的日期解析实现
        std::regex dateRegex(R"((\d{4})-(\d{2})-(\d{2}))");
        std::smatch match;
        
        if (std::regex_match(dateStr, match, dateRegex)) {
            int year = std::stoi(match[1].str());
            int month = std::stoi(match[2].str());
            int day = std::stoi(match[3].str());
            
            struct tm tm = {};
            tm.tm_year = year - 1900;
            tm.tm_mon = month - 1;
            tm.tm_mday = day;
            
            time_t timeT = mktime(&tm);
            return systemTimeToExcelDate(timeT);
        }
        
        return 0.0;
    }
};

// ==================== TXNumberFormat公共接口实现 ====================

TXNumberFormat::TXNumberFormat() : pImpl(std::make_unique<Impl>()) {}

TXNumberFormat::TXNumberFormat(FormatType type, const FormatOptions& options) 
    : pImpl(std::make_unique<Impl>(type, options)) {}

TXNumberFormat::TXNumberFormat(const std::string& customFormat) 
    : pImpl(std::make_unique<Impl>(customFormat)) {}

TXNumberFormat::~TXNumberFormat() = default;

TXNumberFormat::TXNumberFormat(const TXNumberFormat& other) 
    : pImpl(std::make_unique<Impl>(*other.pImpl)) {}

TXNumberFormat& TXNumberFormat::operator=(const TXNumberFormat& other) {
    if (this != &other) {
        *pImpl = *other.pImpl;
    }
    return *this;
}

TXNumberFormat::TXNumberFormat(TXNumberFormat&& other) noexcept 
    : pImpl(std::move(other.pImpl)) {}

TXNumberFormat& TXNumberFormat::operator=(TXNumberFormat&& other) noexcept {
    if (this != &other) {
        pImpl = std::move(other.pImpl);
    }
    return *this;
}

// ==================== 格式设置 ====================

void TXNumberFormat::setFormat(FormatType type, const FormatOptions& options) {
    pImpl->formatType_ = type;
    pImpl->options_ = options;
    pImpl->updatePatterns();
}

void TXNumberFormat::setCustomFormat(const std::string& formatString) {
    pImpl->formatType_ = FormatType::Custom;
    pImpl->customFormatString_ = formatString;
    pImpl->updatePatterns();
}

TXNumberFormat::FormatType TXNumberFormat::getFormatType() const {
    return pImpl->formatType_;
}

std::string TXNumberFormat::getFormatString() const {
    if (pImpl->formatType_ == FormatType::Custom) {
        return pImpl->customFormatString_;
    }
    return generateExcelFormatString();
}

const TXNumberFormat::FormatOptions& TXNumberFormat::getFormatOptions() const {
    return pImpl->options_;
}

// ==================== 格式化方法 ====================

std::string TXNumberFormat::format(const Value& value) const {
    return pImpl->formatValue(value);
}

std::string TXNumberFormat::formatNumber(double number) const {
    return pImpl->formatNumber(number);
}

std::string TXNumberFormat::formatInteger(int64_t integer) const {
    return pImpl->formatNumber(static_cast<double>(integer));
}

std::string TXNumberFormat::formatPercentage(double value) const {
    return pImpl->formatPercentage(value);
}

std::string TXNumberFormat::formatCurrency(double amount) const {
    return pImpl->formatCurrency(amount);
}

std::string TXNumberFormat::formatDate(double excelDate) const {
    return pImpl->formatDate(excelDate);
}

std::string TXNumberFormat::formatTime(double excelTime) const {
    return pImpl->formatTime(excelTime);
}

std::string TXNumberFormat::formatScientific(double number) const {
    return pImpl->formatScientific(number);
}

// ==================== 解析方法 ====================

TXNumberFormat::Value TXNumberFormat::parse(const std::string& formattedStr) const {
    return pImpl->parseValue(formattedStr);
}

bool TXNumberFormat::matches(const std::string& str) const {
    switch (pImpl->formatType_) {
        case FormatType::Number:
        case FormatType::Decimal:
            return std::regex_match(str, pImpl->numberPattern_);
        case FormatType::Date:
            return std::regex_match(str, pImpl->datePattern_);
        case FormatType::Time:
            return std::regex_match(str, pImpl->timePattern_);
        default:
            return true; // 其他格式总是匹配
    }
}

// ==================== 预定义格式 ====================

TXNumberFormat TXNumberFormat::createNumberFormat(int decimalPlaces, bool useThousandSeparator) {
    FormatOptions options;
    options.decimalPlaces = decimalPlaces;
    options.useThousandSeparator = useThousandSeparator;
    return TXNumberFormat(FormatType::Number, options);
}

TXNumberFormat TXNumberFormat::createCurrencyFormat(const std::string& currencySymbol, int decimalPlaces) {
    FormatOptions options;
    options.currencySymbol = currencySymbol;
    options.decimalPlaces = decimalPlaces;
    return TXNumberFormat(FormatType::Currency, options);
}

TXNumberFormat TXNumberFormat::createPercentageFormat(int decimalPlaces) {
    FormatOptions options;
    options.decimalPlaces = decimalPlaces;
    return TXNumberFormat(FormatType::Percentage, options);
}

TXNumberFormat TXNumberFormat::createDateFormat(const std::string& dateFormat) {
    FormatOptions options;
    options.dateFormat = dateFormat;
    return TXNumberFormat(FormatType::Date, options);
}

TXNumberFormat TXNumberFormat::createTimeFormat(const std::string& timeFormat) {
    FormatOptions options;
    options.timeFormat = timeFormat;
    return TXNumberFormat(FormatType::Time, options);
}

TXNumberFormat TXNumberFormat::createScientificFormat(int decimalPlaces) {
    FormatOptions options;
    options.decimalPlaces = decimalPlaces;
    return TXNumberFormat(FormatType::Scientific, options);
}

// ==================== 日期时间工具 ====================

double TXNumberFormat::systemTimeToExcelDate(time_t timeT) {
    return Impl::systemTimeToExcelDate(timeT);
}

time_t TXNumberFormat::excelDateToSystemTime(double excelDate) {
    return Impl::excelDateToSystemTime(excelDate);
}

double TXNumberFormat::getCurrentExcelDate() {
    auto now = std::chrono::system_clock::now();
    time_t timeT = std::chrono::system_clock::to_time_t(now);
    return systemTimeToExcelDate(timeT);
}

double TXNumberFormat::parseDateString(const std::string& dateStr, const std::string& format) {
    return Impl::parseDateString(dateStr, format);
}

// ==================== 格式代码生成 ====================

int TXNumberFormat::getExcelFormatCode() const {
    return static_cast<int>(pImpl->formatType_);
}

std::string TXNumberFormat::generateExcelFormatString() const {
    switch (pImpl->formatType_) {
        case FormatType::General:
            return "General";
        case FormatType::Number:
            return pImpl->options_.useThousandSeparator ? "#,##0.00" : "0.00";
        case FormatType::Currency:
            return pImpl->options_.currencySymbol + "#,##0.00";
        case FormatType::Percentage:
            return "0.00%";
        case FormatType::Date:
            return "yyyy-mm-dd";
        case FormatType::Time:
            return "hh:mm:ss";
        case FormatType::Scientific:
            return "0.00E+00";
        case FormatType::Custom:
            return pImpl->customFormatString_;
        default:
            return "General";
    }
}

// ==================== 验证和工具 ====================

bool TXNumberFormat::isValidFormatString(const std::string& formatString) {
    // 简化的格式字符串验证
    return !formatString.empty();
}

std::unordered_map<TXNumberFormat::FormatType, std::string> TXNumberFormat::getPredefinedFormats() {
    return {
        {FormatType::General, "常规"},
        {FormatType::Number, "数字"},
        {FormatType::Currency, "货币"},
        {FormatType::Percentage, "百分比"},
        {FormatType::Date, "日期"},
        {FormatType::Time, "时间"},
        {FormatType::Scientific, "科学计数法"},
        {FormatType::Text, "文本"}
    };
}

std::string TXNumberFormat::getFormatDescription(FormatType type) {
    auto formats = getPredefinedFormats();
    auto it = formats.find(type);
    return (it != formats.end()) ? it->second : "未知格式";
}

bool TXNumberFormat::isNumericValue(const Value& value) {
    return std::holds_alternative<double>(value) || 
           std::holds_alternative<int64_t>(value) ||
           std::holds_alternative<bool>(value);
}

double TXNumberFormat::valueToNumber(const Value& value) {
    return Impl::valueToNumber(value);
}

} // namespace TinaXlsx 