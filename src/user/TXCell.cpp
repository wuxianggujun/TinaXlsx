//
// @file TXCell.cpp
// @brief 🚀 用户层单元格类实现
//

#include "TinaXlsx/user/TXCell.hpp"
#include "TinaXlsx/TXInMemorySheet.hpp"
#include "TinaXlsx/TXHighPerformanceLogger.hpp"
#include <sstream>

namespace TinaXlsx {

// ==================== 构造和析构 ====================

TXCell::TXCell(TXInMemorySheet& sheet, const TXCoordinate& coord)
    : sheet_(sheet), coord_(coord) {
    // 轻量级构造，无需额外操作
}

TXCell::TXCell(TXInMemorySheet& sheet, std::string_view excel_coord)
    : sheet_(sheet) {
    // 🚀 使用统一的坐标转换工具
    auto result = TXCoordUtils::parseCoord(excel_coord);
    if (result.isOk()) {
        coord_ = result.value();
    } else {
        // 解析失败，使用无效坐标
        coord_ = TXCoordinate(row_t(0), column_t(static_cast<uint32_t>(0)));
        handleError("构造TXCell", result.error());
    }
}

// ==================== 值操作 ====================

TXCell& TXCell::setValue(double value) {
    auto result = sheet_.setNumber(coord_, value);
    if (result.isError()) {
        handleError("设置数值", result.error());
    }
    return *this;
}

TXCell& TXCell::setValue(const std::string& value) {
    auto result = sheet_.setString(coord_, value);
    if (result.isError()) {
        handleError("设置字符串", result.error());
    }
    return *this;
}

TXCell& TXCell::setValue(std::string_view value) {
    return setValue(std::string(value));
}

TXCell& TXCell::setValue(const char* value) {
    return setValue(std::string(value));
}

TXCell& TXCell::setValue(bool value) {
    // 布尔值作为数值存储
    return setValue(value ? 1.0 : 0.0);
}

TXCell& TXCell::setValue(const TXVariant& value) {
    // 根据TXVariant的类型调用相应的方法
    TXResult<void> result;
    switch (value.getType()) {
        case TXVariant::Type::Number:
            result = sheet_.setNumber(coord_, value.getNumber());
            break;
        case TXVariant::Type::String:
            result = sheet_.setString(coord_, value.getString());
            break;
        case TXVariant::Type::Boolean:
            // 布尔值作为数值存储
            result = sheet_.setNumber(coord_, value.getBoolean() ? 1.0 : 0.0);
            break;
        default:
            // 空值或其他类型，暂时不处理
            result = TXResult<void>(TXError(TXErrorCode::InvalidArgument, "不支持的值类型"));
            break;
    }

    if (result.isError()) {
        handleError("设置TXVariant值", result.error());
    }
    return *this;
}

TXCell& TXCell::setFormula(const std::string& formula) {
    // TODO: TXInMemorySheet暂时不支持公式，先作为字符串存储
    TX_LOG_DEBUG("设置公式: {} (暂时作为字符串存储)", formula);
    auto result = sheet_.setString(coord_, formula);
    if (result.isError()) {
        handleError("设置公式", result.error());
    }
    return *this;
}

TXVariant TXCell::getValue() const {
    auto result = sheet_.getValue(coord_);
    if (result.isOk()) {
        return result.value();
    } else {
        handleError("获取值", result.error());
        return TXVariant(); // 返回空值
    }
}

std::string TXCell::getFormula() const {
    // TODO: TXInMemorySheet暂时不支持公式，返回空字符串
    TX_LOG_DEBUG("获取公式 (暂未实现)");
    return "";
}

TXVariant::Type TXCell::getType() const {
    auto value = getValue();
    return value.getType();
}

bool TXCell::isEmpty() const {
    auto value = getValue();
    return value.getType() == TXVariant::Type::Empty ||
           (value.getType() == TXVariant::Type::String && value.getString().empty());
}

TXCell& TXCell::clear() {
    // TODO: TXInMemorySheet暂时不支持单个单元格清除，先设置为空字符串
    TX_LOG_DEBUG("清除单元格 (暂时设置为空字符串)");
    auto result = sheet_.setString(coord_, "");
    if (result.isError()) {
        handleError("清除单元格", result.error());
    }
    return *this;
}

// ==================== 坐标信息 ====================

std::string TXCell::getAddress() const {
    return TXCoordUtils::coordToExcel(coord_);
}

// ==================== 便捷操作符 ====================

bool TXCell::operator==(const TXCell& other) const {
    // 比较坐标和工作表引用
    return coord_ == other.coord_ && &sheet_ == &other.sheet_;
}

// ==================== 数学操作 ====================

TXCell& TXCell::add(double value) {
    auto current = getValue();
    if (current.getType() == TXVariant::Type::Number) {
        return setValue(current.getNumber() + value);
    } else {
        handleError("数学运算", TXError(TXErrorCode::InvalidOperation, "单元格不包含数值"));
        return *this;
    }
}

TXCell& TXCell::subtract(double value) {
    auto current = getValue();
    if (current.getType() == TXVariant::Type::Number) {
        return setValue(current.getNumber() - value);
    } else {
        handleError("数学运算", TXError(TXErrorCode::InvalidOperation, "单元格不包含数值"));
        return *this;
    }
}

TXCell& TXCell::multiply(double value) {
    auto current = getValue();
    if (current.getType() == TXVariant::Type::Number) {
        return setValue(current.getNumber() * value);
    } else {
        handleError("数学运算", TXError(TXErrorCode::InvalidOperation, "单元格不包含数值"));
        return *this;
    }
}

TXCell& TXCell::divide(double value) {
    if (value == 0.0) {
        handleError("数学运算", TXError(TXErrorCode::InvalidArgument, "除数不能为零"));
        return *this;
    }
    
    auto current = getValue();
    if (current.getType() == TXVariant::Type::Number) {
        return setValue(current.getNumber() / value);
    } else {
        handleError("数学运算", TXError(TXErrorCode::InvalidOperation, "单元格不包含数值"));
        return *this;
    }
}

// ==================== 格式化 (预留接口) ====================

TXCell& TXCell::setNumberFormat(const std::string& format) {
    // TODO: 实现数字格式设置
    TX_LOG_DEBUG("设置数字格式: {} (暂未实现)", format);
    return *this;
}

TXCell& TXCell::setFontColor(uint32_t color) {
    // TODO: 实现字体颜色设置
    TX_LOG_DEBUG("设置字体颜色: 0x{:X} (暂未实现)", color);
    return *this;
}

TXCell& TXCell::setBackgroundColor(uint32_t color) {
    // TODO: 实现背景颜色设置
    TX_LOG_DEBUG("设置背景颜色: 0x{:X} (暂未实现)", color);
    return *this;
}

// ==================== 调试和诊断 ====================

std::string TXCell::toString() const {
    std::ostringstream oss;
    oss << "TXCell{";
    oss << "地址=" << getAddress();
    oss << ", 坐标=(" << getRow() << "," << getColumn() << ")";
    
    auto value = getValue();
    oss << ", 类型=";
    switch (value.getType()) {
        case TXVariant::Type::Number:
            oss << "数值, 值=" << value.getNumber();
            break;
        case TXVariant::Type::String:
            oss << "字符串, 值=\"" << value.getString() << "\"";
            break;
        case TXVariant::Type::Boolean:
            oss << "布尔值, 值=" << (value.getBoolean() ? "true" : "false");
            break;
        case TXVariant::Type::Empty:
            oss << "空值";
            break;
        default:
            oss << "未知";
            break;
    }
    
    oss << "}";
    return oss.str();
}

bool TXCell::isValid() const {
    return coord_.isValid();
}

// ==================== 内部辅助方法 ====================

void TXCell::handleError(const std::string& operation, const TXError& error) const {
    TX_LOG_WARN("TXCell操作失败: {} - 地址={}, 错误={}", 
                operation, getAddress(), error.getMessage());
}

} // namespace TinaXlsx
