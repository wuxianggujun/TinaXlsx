#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXSheet.hpp" // For evaluateFormula context
// TXFormula.hpp and TXNumberFormat.hpp are included via TXCell.hpp
#include <sstream>
#include <algorithm> // For std::transform
#include <cctype>    // For ::tolower
#include <iomanip>   // For std::fixed, std::setprecision in fallback formatting

namespace TinaXlsx
{
    TXCell::TXCell()
        : value_(std::monostate{}), // 默认空值
          type_(CellType::Empty),
          formula_object_(nullptr),
          number_format_object_(nullptr), // 初始无特定格式对象
          is_merged_(false),
          is_master_cell_(false),
          master_row_idx_(0),
          master_col_idx_(0),
          has_style_(false),
          style_index_(0),
          is_locked_(true) // 默认锁定
    {
        // 默认单元格使用“常规”格式，可以通过创建一个默认的TXNumberFormat对象来实现
        // 或者让getFormattedValue在number_format_object_为nullptr时有特定行为
        // 为保持一致性，可以创建一个默认的通用格式对象
        number_format_object_ = std::make_unique<TXNumberFormat>(TXNumberFormat::FormatType::General);
    }

    TXCell::TXCell(const CellValue& value)
        : value_(value),
          formula_object_(nullptr),
          number_format_object_(std::make_unique<TXNumberFormat>(TXNumberFormat::FormatType::General)), // 默认常规格式
          is_merged_(false),
          is_master_cell_(false),
          master_row_idx_(0),
          master_col_idx_(0),
          has_style_(false),
          style_index_(0),
          is_locked_(true) // 默认锁定
    {
        updateType(); // 类型会根据值自动更新
    }

    TXCell::~TXCell() = default;

    // 拷贝构造函数
    TXCell::TXCell(const TXCell& other)
        : value_(other.value_),
          type_(other.type_),
          is_merged_(other.is_merged_),
          is_master_cell_(other.is_master_cell_),
          master_row_idx_(other.master_row_idx_),
          master_col_idx_(other.master_col_idx_),
          has_style_(other.has_style_),
          style_index_(other.style_index_),
          is_locked_(other.is_locked_)
    {
        if (other.formula_object_) {
            formula_object_ = std::make_unique<TXFormula>(*other.formula_object_);
        }
        if (other.number_format_object_) {
            number_format_object_ = std::make_unique<TXNumberFormat>(*other.number_format_object_);
        } else {
            // 确保即使源的格式对象为空，此副本也有一个默认的
            number_format_object_ = std::make_unique<TXNumberFormat>(TXNumberFormat::FormatType::General);
        }
    }

    // 拷贝赋值操作符
    TXCell& TXCell::operator=(const TXCell& other) {
        if (this != &other) {
            value_ = other.value_;
            type_ = other.type_;

            if (other.formula_object_) {
                formula_object_ = std::make_unique<TXFormula>(*other.formula_object_);
            } else {
                formula_object_.reset();
            }

            if (other.number_format_object_) {
                number_format_object_ = std::make_unique<TXNumberFormat>(*other.number_format_object_);
            } else {
                number_format_object_ = std::make_unique<TXNumberFormat>(TXNumberFormat::FormatType::General);
            }

            is_merged_ = other.is_merged_;
            is_master_cell_ = other.is_master_cell_;
            master_row_idx_ = other.master_row_idx_;
            master_col_idx_ = other.master_col_idx_;
            has_style_ = other.has_style_;
            style_index_ = other.style_index_;
            is_locked_ = other.is_locked_;
        }
        return *this;
    }

    // 移动构造函数
    TXCell::TXCell(TXCell&& other) noexcept
        : value_(std::move(other.value_)),
          type_(other.type_),
          formula_object_(std::move(other.formula_object_)),
          number_format_object_(std::move(other.number_format_object_)),
          is_merged_(other.is_merged_),
          is_master_cell_(other.is_master_cell_),
          master_row_idx_(other.master_row_idx_),
          master_col_idx_(other.master_col_idx_),
          has_style_(other.has_style_),
          style_index_(other.style_index_),
          is_locked_(other.is_locked_)
    {
        // 将源对象置于有效的空状态
        other.type_ = CellType::Empty;
        other.value_ = std::monostate{};
        other.is_merged_ = false;
        other.is_master_cell_ = false;
        other.has_style_ = false;
        other.style_index_ = 0;
        // 确保源对象的 number_format_object_ 也被置为有效状态（如果它被移走后为 null）
        if (!other.number_format_object_) { // 如果被移走
             other.number_format_object_ = std::make_unique<TXNumberFormat>(TXNumberFormat::FormatType::General);
        }
    }

    // 移动赋值操作符
    TXCell& TXCell::operator=(TXCell&& other) noexcept {
        if (this != &other) {
            value_ = std::move(other.value_);
            type_ = other.type_;
            formula_object_ = std::move(other.formula_object_);
            number_format_object_ = std::move(other.number_format_object_);
            is_merged_ = other.is_merged_;
            is_master_cell_ = other.is_master_cell_;
            master_row_idx_ = other.master_row_idx_;
            master_col_idx_ = other.master_col_idx_;
            has_style_ = other.has_style_;
            style_index_ = other.style_index_;
            is_locked_ = other.is_locked_;

            other.type_ = CellType::Empty;
            other.value_ = std::monostate{};
            other.is_merged_ = false;
            other.is_master_cell_ = false;
            other.has_style_ = false;
            other.style_index_ = 0;
            if (!other.number_format_object_) {
                other.number_format_object_ = std::make_unique<TXNumberFormat>(TXNumberFormat::FormatType::General);
            }
        }
        return *this;
    }

    const TXCell::CellValue& TXCell::getValue() const {
        return value_;
    }

    void TXCell::setValue(const CellValue& value) {
        value_ = value;
        // 如果设置了值，并且当前是公式类型但值不是由公式计算的（例如用户直接覆盖），
        // 则应该清除公式。但通常 setValue 用于设置非公式值。
        // 如果确实要清除公式，则：
        // if (type_ == CellType::Formula) {
        //     formula_object_.reset();
        // }
        // 这里假设 setValue 是用于数据单元格，或者公式单元格的缓存结果。
        // updateType 会在没有 formula_object_ 时根据 value_ 调整类型。
        updateType();
    }

    TXCell::CellType TXCell::getType() const {
        return type_;
    }

    bool TXCell::isEmpty() const {
        // 一个单元格被认为是空的，如果：
        // 1. 值为monostate，或
        // 2. 值为空字符串
        // 并且没有公式
        if (formula_object_) {
            return false; // 有公式的单元格不为空
        }

        if (std::holds_alternative<std::monostate>(value_)) {
            return true;
        }

        if (std::holds_alternative<std::string>(value_)) {
            return std::get<std::string>(value_).empty();
        }

        return false; // 其他类型的值不为空
    }

    std::string TXCell::getStringValue() const {
        if (std::holds_alternative<std::string>(value_)) {
            return std::get<std::string>(value_);
        }
        if (std::holds_alternative<double>(value_)) {
            std::ostringstream oss;
            oss << std::get<double>(value_);
            return oss.str();
        }
        if (std::holds_alternative<int64_t>(value_)) {
            return std::to_string(std::get<int64_t>(value_));
        }
        if (std::holds_alternative<bool>(value_)) {
            return std::get<bool>(value_) ? "TRUE" : "FALSE";
        }
        // 对于 std::monostate (Empty) 或其他未处理类型
        return "";
    }

    double TXCell::getNumberValue() const {
        if (std::holds_alternative<double>(value_)) {
            return std::get<double>(value_);
        }
        if (std::holds_alternative<int64_t>(value_)) {
            return static_cast<double>(std::get<int64_t>(value_));
        }
        if (std::holds_alternative<bool>(value_)) {
            return std::get<bool>(value_) ? 1.0 : 0.0;
        }
        if (std::holds_alternative<std::string>(value_)) {
            try {
                return std::stod(std::get<std::string>(value_));
            } catch (...) { /* 转换失败 */ }
        }
        return 0.0;
    }

    int64_t TXCell::getIntegerValue() const {
        if (std::holds_alternative<int64_t>(value_)) {
            return std::get<int64_t>(value_);
        }
        if (std::holds_alternative<double>(value_)) {
            return static_cast<int64_t>(std::get<double>(value_)); // 注意截断
        }
        if (std::holds_alternative<bool>(value_)) {
            return std::get<bool>(value_) ? 1 : 0;
        }
        if (std::holds_alternative<std::string>(value_)) {
            try {
                return std::stoll(std::get<std::string>(value_));
            } catch (...) { /* 转换失败 */ }
        }
        return 0;
    }

    bool TXCell::getBooleanValue() const {
        if (std::holds_alternative<bool>(value_)) {
            return std::get<bool>(value_);
        }
        if (std::holds_alternative<double>(value_)) {
            return std::get<double>(value_) != 0.0;
        }
        if (std::holds_alternative<int64_t>(value_)) {
            return std::get<int64_t>(value_) != 0;
        }
        if (std::holds_alternative<std::string>(value_)) {
            std::string s = std::get<std::string>(value_);
            std::transform(s.begin(), s.end(), s.begin(), ::tolower);
            return s == "true";
        }
        return false;
    }

    void TXCell::setStringValue(const std::string& value) {
        setValue(value);
        if (isFormula()) formula_object_.reset(); // 如果设置值，清除公式
        updateType();
    }

    void TXCell::setNumberValue(double value) {
        setValue(value);
        if (isFormula()) formula_object_.reset();
        updateType();
    }

    void TXCell::setIntegerValue(int64_t value) {
        setValue(value);
        if (isFormula()) formula_object_.reset();
        updateType();
    }

    void TXCell::setBooleanValue(bool value) {
        setValue(value);
        if (isFormula()) formula_object_.reset();
        updateType();
    }

    std::string TXCell::getFormula() const {
        if (formula_object_) {
            return formula_object_->getFormulaString();
        }
        return "";
    }

    void TXCell::setFormula(const std::string& formula_str) {
        if (formula_str.empty()) {
            formula_object_.reset();
            // value_ 保持不变 (可能是之前公式的缓存结果或用户设置的值)
            updateType(); // updateType 会根据 value_ (因为 formula_object_ 为空) 设置类型
        } else {
            if (formula_object_) {
                formula_object_->setFormulaString(formula_str); // 复用对象
            } else {
                formula_object_ = std::make_unique<TXFormula>(formula_str);
            }
            type_ = CellType::Formula; // 显式设置类型为公式
            // value_ 此时可能过时，evaluateFormula 被调用时会更新它
        }
    }

    bool TXCell::isFormula() const {
        // 单元格是公式的唯一判断标准是有无 formula_object_
        // type_ 应当与此状态保持一致
        return formula_object_ != nullptr;
    }

    const TXFormula* TXCell::getFormulaObject() const {
        return formula_object_.get();
    }

    void TXCell::setFormulaObject(std::unique_ptr<TXFormula> formula_ptr) {
        formula_object_ = std::move(formula_ptr);
        if (formula_object_) {
            type_ = CellType::Formula;
        } else {
            updateType(); // 如果公式对象被移除，则根据当前值更新类型
        }
    }

    TXCell::CellValue TXCell::evaluateFormula(const TXSheet* sheet, row_t currentRow, column_t currentCol) {
        if (!formula_object_) {
            // 不是公式单元格，或者公式已被清除，返回当前值
            return value_;
        }

        // 假设 TXFormula::evaluate 返回 TXFormula::FormulaValue (也是 cell_value_t)
        CellValue result = formula_object_->evaluate(sheet, currentRow, currentCol);
        
        // 可选：将计算结果更新到单元格的 value_ 中作为缓存
        // value_ = result;
        // updateType(); // 如果 value_ 更新了，类型也可能需要更新（尽管它仍是Formula类型）
        
        return result;
    }


    void TXCell::setCustomFormat(const std::string& format_string) {
        if (!number_format_object_ ||
            number_format_object_->getFormatType() != TXNumberFormat::FormatType::Custom ||
            number_format_object_->getFormatString() != format_string) {
            number_format_object_ = std::make_unique<TXNumberFormat>(format_string);
        }
    }

    const TXNumberFormat* TXCell::getNumberFormatObject() const {
        return number_format_object_.get();
    }

    void TXCell::setNumberFormatObject(std::unique_ptr<TXNumberFormat> number_format_ptr) {
        number_format_object_ = std::move(number_format_ptr);
        if (!number_format_object_) { // 确保总有一个有效的格式对象
            number_format_object_ = std::make_unique<TXNumberFormat>(TXNumberFormat::FormatType::General);
        }
    }

    std::string TXCell::getFormattedValue() const {
        if (number_format_object_) {
            return number_format_object_->format(value_);
        }
        // 如果没有格式对象（理论上构造函数会创建一个默认的），则回退到简单字符串转换
        return getStringValue();
    }

    void TXCell::setPredefinedFormat(TXNumberFormat::FormatType type, int decimalPlaces, bool useThousandSeparator) {
        TXNumberFormat::FormatOptions options;
        options.decimalPlaces = decimalPlaces;
        options.useThousandSeparator = useThousandSeparator;
        // 注意: 对于货币、日期等，可能需要从 TXWorkbook 或其他地方获取默认符号/格式字符串
        // 例如 options.currencySymbol = workbook->getLocaleCurrencySymbol();

        number_format_object_ = std::make_unique<TXNumberFormat>(type, options);
    }

    bool TXCell::isMerged() const { return is_merged_; }
    void TXCell::setMerged(bool merged) { is_merged_ = merged; }
    bool TXCell::isMasterCell() const { return is_master_cell_; }
    void TXCell::setMasterCell(bool master) { is_master_cell_ = master; }

    std::pair<row_t::index_t, column_t::index_t> TXCell::getMasterCellPosition() const {
        return {master_row_idx_, master_col_idx_};
    }

    void TXCell::setMasterCellPosition(row_t::index_t row_idx, column_t::index_t col_idx) {
        master_row_idx_ = row_idx;
        master_col_idx_ = col_idx;
    }

    bool TXCell::hasStyle() const { return has_style_; }
    u32 TXCell::getStyleIndex() const { return style_index_; }
    void TXCell::setStyleIndex(u32 index) {
        style_index_ = index;
        has_style_ = (index != 0); // 假设索引0表示无特定样式或默认样式
    }

    void TXCell::clear() {
        value_ = std::monostate{};
        type_ = CellType::Empty;
        formula_object_.reset(); // 清除公式
        // 保留数字格式对象，但可以重置为 General，或者根据需求清除
        if (number_format_object_) {
            number_format_object_->setFormat(TXNumberFormat::FormatType::General);
        } else {
            number_format_object_ = std::make_unique<TXNumberFormat>(TXNumberFormat::FormatType::General);
        }
        is_merged_ = false;
        is_master_cell_ = false;
        master_row_idx_ = 0;
        master_col_idx_ = 0;
        has_style_ = false;
        style_index_ = 0;
    }

    std::string TXCell::toString() const {
        // 通常 toString() 返回用户最期望看到的文本形式
        return getFormattedValue();
    }

    bool TXCell::fromString(const std::string& str, bool auto_detect_type) {
        formula_object_.reset(); // 从字符串设置值时，清除任何现有公式

        if (!auto_detect_type) {
            setValue(str);
            return true;
        }

        if (str.empty()) {
            setValue(std::monostate{});
            updateType();
            return true;
        }

        std::string lower_str = str;
        std::transform(lower_str.begin(), lower_str.end(), lower_str.begin(),
                       [](unsigned char c){ return static_cast<char>(::tolower(c)); });

        if (lower_str == "true") {
            setValue(true);
            return true;
        }
        if (lower_str == "false") {
            setValue(false);
            return true;
        }

        // 尝试解析为整数
        try {
            size_t parsed_chars = 0;
            int64_t i_val = std::stoll(str, &parsed_chars);
            if (parsed_chars == str.length()) { // 确保整个字符串都被解析为整数
                setValue(i_val);
                return true;
            }
        } catch (const std::invalid_argument&) {
        } catch (const std::out_of_range&) {}

        // 尝试解析为浮点数
        try {
            size_t parsed_chars = 0;
            double d_val = std::stod(str, &parsed_chars);
            if (parsed_chars == str.length()) { // 确保整个字符串都被解析为浮点数
                setValue(d_val);
                return true;
            }
        } catch (const std::invalid_argument&) {
        } catch (const std::out_of_range&) {}

        // 默认设为字符串
        setValue(str);
        return true;
    }


    std::unique_ptr<TXCell> TXCell::clone() const {
        return std::make_unique<TXCell>(*this); // 使用拷贝构造函数
    }

    void TXCell::copyFormatTo(TXCell& target) const {
        if (number_format_object_) {
            target.setNumberFormatObject(std::make_unique<TXNumberFormat>(*number_format_object_));
        } else {
            target.setNumberFormatObject(nullptr); // 或一个默认的通用格式对象
        }
        target.setStyleIndex(style_index_);
        target.has_style_ = has_style_;
        // 注意：合并状态通常不通过此方法复制，因为它是区域属性
    }

    bool TXCell::isValueEqual(const TXCell& other) const {
        return value_ == other.value_;
    }

    // 赋值操作符实现
    TXCell& TXCell::operator=(const std::string& value_str) { setStringValue(value_str); return *this; }
    TXCell& TXCell::operator=(const char* value_cstr) { setStringValue(value_cstr ? std::string(value_cstr) : std::string()); return *this; }
    TXCell& TXCell::operator=(double value_dbl) { setNumberValue(value_dbl); return *this; }
    TXCell& TXCell::operator=(int64_t value_i64) { setIntegerValue(value_i64); return *this; }
    TXCell& TXCell::operator=(int value_int) { setIntegerValue(static_cast<int64_t>(value_int)); return *this; }
    TXCell& TXCell::operator=(bool value_bool) { setBooleanValue(value_bool); return *this; }

    // 比较操作符实现 (基于 std::variant 的默认比较)
    bool TXCell::operator==(const TXCell& other) const { return value_ == other.value_ && isFormula() == other.isFormula() && getFormula() == other.getFormula(); }
    bool TXCell::operator!=(const TXCell& other) const { return !(*this == other); }
    bool TXCell::operator<(const TXCell& other) const {
        if (isFormula() != other.isFormula()) return isFormula() < other.isFormula();
        if (isFormula()) { // Both are formulas
            int comp = getFormula().compare(other.getFormula());
            if (comp != 0) return comp < 0;
        }
        return value_ < other.value_; // Fallback to value comparison
    }
    bool TXCell::operator<=(const TXCell& other) const { return (*this < other) || (*this == other); }
    bool TXCell::operator>(const TXCell& other) const { return !(*this <= other); }
    bool TXCell::operator>=(const TXCell& other) const { return !(*this < other); }

    void TXCell::updateType() {
        if (formula_object_) { // 如果存在公式对象，则单元格类型定义为公式
            type_ = CellType::Formula;
            return; // 公式类型优先
        }

        // 如果没有公式，则根据值的类型来确定单元格类型
        std::visit([this](auto&& arg) {
            using T = std::decay_t<decltype(arg)>;
            if constexpr (std::is_same_v<T, std::monostate>) {
                type_ = CellType::Empty;
            } else if constexpr (std::is_same_v<T, std::string>) {
                type_ = CellType::String;
            } else if constexpr (std::is_same_v<T, double>) {
                type_ = CellType::Number;
            } else if constexpr (std::is_same_v<T, int64_t>) {
                type_ = CellType::Integer;
            } else if constexpr (std::is_same_v<T, bool>) {
                type_ = CellType::Boolean;
            }
            // CellType::Error 需要显式设置，不会通过 updateType 自动检测
        }, value_);
    }

    // ==================== 保护功能实现 ====================

    void TXCell::setLocked(bool locked) {
        is_locked_ = locked;

        // 注意：单元格锁定状态的实际应用需要通过样式系统来实现
        // 这里只是设置内部状态，实际的样式更新需要在TXSheet层面处理
        // 因为TXCell本身不直接管理样式对象，而是通过样式索引来引用
    }

    bool TXCell::isLocked() const {
        return is_locked_;
    }

    bool TXCell::hasFormula() const {
        return formula_object_ != nullptr;
    }

} // namespace TinaXlsx
