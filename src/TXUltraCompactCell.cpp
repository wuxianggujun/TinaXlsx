//
// @file TXUltraCompactCell.cpp
// @brief 超紧凑单元格实现
//

#include "TinaXlsx/TXUltraCompactCell.hpp"
#include <cstring>
#include <algorithm>

namespace TinaXlsx {

// ==================== 构造函数实现 ====================

UltraCompactCell::UltraCompactCell() {
    clear();
}

UltraCompactCell::UltraCompactCell(const std::string& value, uint32_t string_offset) {
    clear();
    secondary.cell_type = static_cast<uint8_t>(CellType::String);
    primary.string.offset = string_offset;
    primary.string.length = static_cast<uint16_t>(std::min(value.length(), size_t(65535)));
}

UltraCompactCell::UltraCompactCell(double value) {
    clear();
    secondary.cell_type = static_cast<uint8_t>(CellType::Number);
    primary.number_value = value;
}

UltraCompactCell::UltraCompactCell(int64_t value) {
    clear();
    secondary.cell_type = static_cast<uint8_t>(CellType::Integer);
    primary.integer_value = value;
}

UltraCompactCell::UltraCompactCell(bool value) {
    clear();
    secondary.cell_type = static_cast<uint8_t>(CellType::Boolean);
    primary.boolean.value = value ? 1 : 0;
}

// ==================== 数据访问方法实现 ====================

UltraCompactCell::CellType UltraCompactCell::getType() const {
    return static_cast<CellType>(secondary.cell_type);
}

void UltraCompactCell::setType(CellType type) {
    secondary.cell_type = static_cast<uint8_t>(type);
}

double UltraCompactCell::getNumberValue() const {
    if (getType() == CellType::Number) {
        return primary.number_value;
    }
    return 0.0;
}

int64_t UltraCompactCell::getIntegerValue() const {
    if (getType() == CellType::Integer) {
        return primary.integer_value;
    }
    return 0;
}

bool UltraCompactCell::getBooleanValue() const {
    if (getType() == CellType::Boolean) {
        return primary.boolean.value == 1;
    }
    return false;
}

uint32_t UltraCompactCell::getStringOffset() const {
    if (getType() == CellType::String || getType() == CellType::Formula) {
        return primary.string.offset;
    }
    return 0;
}

uint16_t UltraCompactCell::getStringLength() const {
    if (getType() == CellType::String || getType() == CellType::Formula) {
        return primary.string.length;
    }
    return 0;
}

// ==================== 样式和属性实现 ====================

bool UltraCompactCell::hasStyle() const {
    return (secondary.flags & 0x01) != 0;
}

void UltraCompactCell::setHasStyle(bool has_style) {
    if (has_style) {
        secondary.flags |= 0x01;
    } else {
        secondary.flags &= ~0x01;
    }
}

uint8_t UltraCompactCell::getStyleIndex() const {
    return secondary.style_index;
}

void UltraCompactCell::setStyleIndex(uint8_t index) {
    secondary.style_index = index;
    setHasStyle(index != 0);
}

bool UltraCompactCell::isFormula() const {
    return (secondary.flags & 0x02) != 0;
}

void UltraCompactCell::setIsFormula(bool is_formula) {
    if (is_formula) {
        secondary.flags |= 0x02;
        if (getType() == CellType::String) {
            setType(CellType::Formula);
        }
    } else {
        secondary.flags &= ~0x02;
    }
}

uint32_t UltraCompactCell::getFormulaOffset() const {
    // 检查类型或标志位
    if (getType() == CellType::Formula || isFormula()) {
        // 组合低8位和高24位
        uint32_t low = secondary.formula_offset_low;
        uint32_t high = 0;

        // 如果是公式类型，从primary的padding中获取高位
        if (getType() == CellType::Formula) {
            // 使用string结构的padding字段存储高位
            high = static_cast<uint32_t>(primary.string.padding) << 8;
        }

        return high | low;
    }
    return 0;
}

void UltraCompactCell::setFormulaOffset(uint32_t offset) {
    // 存储低8位
    secondary.formula_offset_low = static_cast<uint8_t>(offset & 0xFF);

    // 如果是公式类型，存储高24位到primary的padding中
    if (getType() == CellType::Formula || isFormula()) {
        primary.string.padding = static_cast<uint16_t>((offset >> 8) & 0xFFFF);
    }
}

bool UltraCompactCell::isMerged() const {
    return (secondary.flags & 0x04) != 0;
}

void UltraCompactCell::setIsMerged(bool is_merged) {
    if (is_merged) {
        secondary.flags |= 0x04;
    } else {
        secondary.flags &= ~0x04;
    }
}

// ==================== 批处理优化方法实现 ====================

void UltraCompactCell::encodeBatch(const std::vector<cell_value_t>& values,
                                  const std::vector<TXCoordinate>& coords,
                                  const char* string_buffer,
                                  UltraCompactCell* output,
                                  size_t count) {
    // 基础实现，后续可以用SIMD优化
    for (size_t i = 0; i < count && i < values.size() && i < coords.size(); ++i) {
        UltraCompactCell& cell = output[i];
        cell.clear();
        cell.setCoordinate(coords[i]);
        
        // 根据值类型设置单元格
        std::visit([&cell, string_buffer](const auto& value) {
            using T = std::decay_t<decltype(value)>;
            if constexpr (std::is_same_v<T, std::monostate>) {
                cell.setType(CellType::Empty);
            } else if constexpr (std::is_same_v<T, std::string>) {
                // 这里需要计算字符串在缓冲区中的偏移量
                // 简化实现：假设字符串已经在缓冲区中
                cell = UltraCompactCell(value, 0); // 偏移量需要外部计算
            } else if constexpr (std::is_same_v<T, double>) {
                cell = UltraCompactCell(value);
            } else if constexpr (std::is_same_v<T, int64_t>) {
                cell = UltraCompactCell(value);
            } else if constexpr (std::is_same_v<T, bool>) {
                cell = UltraCompactCell(value);
            }
        }, values[i]);
    }
}

void UltraCompactCell::decodeBatch(const UltraCompactCell* input,
                                  const char* string_buffer,
                                  std::vector<cell_value_t>& values,
                                  std::vector<TXCoordinate>& coords,
                                  size_t count) {
    values.resize(count);
    coords.resize(count);
    
    for (size_t i = 0; i < count; ++i) {
        const UltraCompactCell& cell = input[i];
        coords[i] = cell.getCoordinate();
        
        switch (cell.getType()) {
            case CellType::Empty:
                values[i] = std::monostate{};
                break;
            case CellType::String:
            case CellType::Formula: {
                uint32_t offset = cell.getStringOffset();
                uint16_t length = cell.getStringLength();
                if (string_buffer && offset + length <= UINT32_MAX) {
                    values[i] = std::string(string_buffer + offset, length);
                } else {
                    values[i] = std::string{};
                }
                break;
            }
            case CellType::Number:
                values[i] = cell.getNumberValue();
                break;
            case CellType::Integer:
                values[i] = cell.getIntegerValue();
                break;
            case CellType::Boolean:
                values[i] = cell.getBooleanValue();
                break;
            default:
                values[i] = std::monostate{};
                break;
        }
    }
}

// ==================== 内部辅助方法实现 ====================

void UltraCompactCell::clear() {
    primary.raw_primary = 0;
    secondary.cell_type = 0;
    secondary.style_index = 0;
    secondary.flags = 0;
    secondary.formula_offset_low = 0;
    secondary.row = 0;
    secondary.col = 0;
}

// ==================== 比较操作符实现 ====================

bool UltraCompactCell::operator==(const UltraCompactCell& other) const {
    // 首先比较类型
    if (getType() != other.getType()) {
        return false;
    }

    // 比较坐标
    if (getRow() != other.getRow() || getCol() != other.getCol()) {
        return false;
    }

    // 比较样式
    if (getStyleIndex() != other.getStyleIndex()) {
        return false;
    }

    // 比较标志位
    if (hasStyle() != other.hasStyle() ||
        isFormula() != other.isFormula() ||
        isMerged() != other.isMerged()) {
        return false;
    }

    // 根据类型比较值
    switch (getType()) {
        case CellType::Empty:
            return true; // 空单元格都相等

        case CellType::String:
            return getStringOffset() == other.getStringOffset() &&
                   getStringLength() == other.getStringLength();

        case CellType::Number:
            return getNumberValue() == other.getNumberValue();

        case CellType::Integer:
            return getIntegerValue() == other.getIntegerValue();

        case CellType::Boolean:
            return getBooleanValue() == other.getBooleanValue();

        case CellType::Formula:
            return getFormulaOffset() == other.getFormulaOffset();

        default:
            return false;
    }
}

bool UltraCompactCell::operator!=(const UltraCompactCell& other) const {
    return !(*this == other);
}

uint8_t* UltraCompactCell::getTypeField() {
    // 不能获取位域的地址，返回nullptr或使用其他方式
    return nullptr;
}

const uint8_t* UltraCompactCell::getTypeField() const {
    // 不能获取位域的地址，返回nullptr或使用其他方式
    return nullptr;
}

uint8_t* UltraCompactCell::getStyleField() {
    // 不能获取位域的地址，返回nullptr或使用其他方式
    return nullptr;
}

const uint8_t* UltraCompactCell::getStyleField() const {
    // 不能获取位域的地址，返回nullptr或使用其他方式
    return nullptr;
}

} // namespace TinaXlsx
