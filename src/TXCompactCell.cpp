//
// @file TXCompactCell.cpp
// @brief 紧凑型单元格实现
//

#include "TinaXlsx/TXCompactCell.hpp"
#include "TinaXlsx/TXFormula.hpp"
#include "TinaXlsx/TXNumberFormat.hpp"
#include "TinaXlsx/TXCoordinate.hpp"
#include <algorithm>
#include <sstream>
#include <variant>
#include <unordered_map>

namespace TinaXlsx {

// ==================== ExtendedData 定义 ====================

struct TXCompactCell::ExtendedData {
    std::unique_ptr<TXFormula> formula;
    std::unique_ptr<TXNumberFormat> number_format;
    uint32_t style_index = 0;
};

// ==================== TXCompactCell 实现 ====================

TXCompactCell::TXCompactCell() 
    : value_(std::monostate{})
    , extended_data_(nullptr) {
    // 初始化位域结构
    flags_.type_ = static_cast<uint8_t>(CellType::Empty);
    flags_.has_style_ = 0;
    flags_.is_merged_ = 0;
    flags_.is_master_cell_ = 0;
    flags_.is_locked_ = 1; // 默认锁定
    flags_.reserved_ = 0;
    flags_.master_row_high_ = 0;
    flags_.master_row_low_ = 0;
    flags_.master_col_ = 0;
}

TXCompactCell::TXCompactCell(const cell_value_t& value)
    : TXCompactCell() {
    setValue(value);
}

TXCompactCell::~TXCompactCell() = default;

TXCompactCell::TXCompactCell(TXCompactCell&& other) noexcept 
    : value_(std::move(other.value_))
    , flags_(other.flags_)
    , extended_data_(std::move(other.extended_data_)) {
    // 重置源对象
    other.flags_.type_ = static_cast<uint8_t>(CellType::Empty);
    other.value_ = std::monostate{};
}

TXCompactCell& TXCompactCell::operator=(TXCompactCell&& other) noexcept {
    if (this != &other) {
        value_ = std::move(other.value_);
        flags_ = other.flags_;
        extended_data_ = std::move(other.extended_data_);
        
        // 重置源对象
        other.flags_.type_ = static_cast<uint8_t>(CellType::Empty);
        other.value_ = std::monostate{};
    }
    return *this;
}

TXCompactCell::TXCompactCell(const TXCompactCell& other) 
    : value_(other.value_)
    , flags_(other.flags_)
    , extended_data_(nullptr) {
    // 深拷贝扩展数据
    if (other.extended_data_) {
        extended_data_ = std::make_unique<ExtendedData>();
        extended_data_->style_index = other.extended_data_->style_index;
        
        if (other.extended_data_->formula) {
            extended_data_->formula = std::make_unique<TXFormula>(*other.extended_data_->formula);
        }
        
        if (other.extended_data_->number_format) {
            extended_data_->number_format = std::make_unique<TXNumberFormat>(*other.extended_data_->number_format);
        }
    }
}

TXCompactCell& TXCompactCell::operator=(const TXCompactCell& other) {
    if (this != &other) {
        value_ = other.value_;
        flags_ = other.flags_;
        extended_data_.reset();
        
        // 深拷贝扩展数据
        if (other.extended_data_) {
            extended_data_ = std::make_unique<ExtendedData>();
            extended_data_->style_index = other.extended_data_->style_index;
            
            if (other.extended_data_->formula) {
                extended_data_->formula = std::make_unique<TXFormula>(*other.extended_data_->formula);
            }
            
            if (other.extended_data_->number_format) {
                extended_data_->number_format = std::make_unique<TXNumberFormat>(*other.extended_data_->number_format);
            }
        }
    }
    return *this;
}

// ==================== 值操作 ====================

void TXCompactCell::setValue(const cell_value_t& value) {
    value_ = value;
    flags_.type_ = static_cast<uint8_t>(inferType(value));
}

cell_value_t TXCompactCell::getValue() const {
    return value_;
}

// ==================== 样式操作 ====================

void TXCompactCell::setStyleIndex(uint32_t index) {
    if (index == 0) {
        flags_.has_style_ = 0;
        if (extended_data_) {
            extended_data_->style_index = 0;
            cleanupExtendedData();
        }
    } else {
        flags_.has_style_ = 1;
        ensureExtendedData();
        extended_data_->style_index = index;
    }
}

uint32_t TXCompactCell::getStyleIndex() const {
    if (!flags_.has_style_ || !extended_data_) {
        return 0;
    }
    return extended_data_->style_index;
}

// ==================== 合并单元格操作 ====================

void TXCompactCell::setMerged(bool is_master, uint16_t master_row, uint16_t master_col) {
    flags_.is_merged_ = 1;
    flags_.is_master_cell_ = is_master ? 1 : 0;
    
    if (!is_master) {
        // 存储主单元格坐标（限制在16位范围内）
        flags_.master_row_high_ = (master_row >> 8) & 0xFF;
        flags_.master_row_low_ = master_row & 0xFF;
        flags_.master_col_ = std::min(master_col, static_cast<uint16_t>(255));
    }
}

// ==================== 扩展数据操作 ====================

void TXCompactCell::setFormula(std::unique_ptr<TXFormula> formula) {
    if (formula) {
        ensureExtendedData();
        extended_data_->formula = std::move(formula);
        flags_.type_ = static_cast<uint8_t>(CellType::Formula);
    } else {
        if (extended_data_) {
            extended_data_->formula.reset();
            cleanupExtendedData();
        }
        // 如果移除公式，恢复为原始值类型
        flags_.type_ = static_cast<uint8_t>(inferType(value_));
    }
}

const TXFormula* TXCompactCell::getFormula() const {
    if (!extended_data_) {
        return nullptr;
    }
    return extended_data_->formula.get();
}

void TXCompactCell::setNumberFormat(std::unique_ptr<TXNumberFormat> format) {
    if (format) {
        ensureExtendedData();
        extended_data_->number_format = std::move(format);
    } else {
        if (extended_data_) {
            extended_data_->number_format.reset();
            cleanupExtendedData();
        }
    }
}

const TXNumberFormat* TXCompactCell::getNumberFormat() const {
    if (!extended_data_) {
        return nullptr;
    }
    return extended_data_->number_format.get();
}

// ==================== 兼容性接口 ====================

std::string TXCompactCell::getFormattedValue() const {
    // 如果有数字格式，使用它来格式化
    if (extended_data_ && extended_data_->number_format) {
        return extended_data_->number_format->format(value_);
    }
    
    // 否则使用默认格式化
    return std::visit([](const auto& val) -> std::string {
        using T = std::decay_t<decltype(val)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            return "";
        } else if constexpr (std::is_same_v<T, std::string>) {
            return val;
        } else if constexpr (std::is_same_v<T, double>) {
            std::ostringstream oss;
            oss << val;
            return oss.str();
        } else if constexpr (std::is_same_v<T, int64_t>) {
            return std::to_string(val);
        } else if constexpr (std::is_same_v<T, bool>) {
            return val ? "TRUE" : "FALSE";
        }
        return "";
    }, value_);
}

bool TXCompactCell::hasFormula() const {
    return extended_data_ && extended_data_->formula != nullptr;
}

void TXCompactCell::setFormula(const std::string& formulaText) {
    if (formulaText.empty()) {
        // 清除公式
        if (extended_data_) {
            extended_data_->formula.reset();
            cleanupExtendedData();
        }
        // 如果移除公式，恢复为原始值类型
        flags_.type_ = static_cast<uint8_t>(inferType(value_));
        return;
    }

    // 确保扩展数据存在
    ensureExtendedData();

    // 创建TXFormula对象
    extended_data_->formula = std::make_unique<TXFormula>(formulaText);

    // 更新类型
    flags_.type_ = static_cast<uint8_t>(CellType::Formula);
}

std::string TXCompactCell::getFormulaText() const {
    if (hasFormula()) {
        return extended_data_->formula->getFormulaString();
    }
    return "";
}

void TXCompactCell::setNumberFormatObject(std::unique_ptr<TXNumberFormat> format) {
    setNumberFormat(std::move(format));
}

const TXNumberFormat* TXCompactCell::getNumberFormatObject() const {
    return getNumberFormat();
}

std::string TXCompactCell::getStringValue() const {
    if (std::holds_alternative<std::string>(value_)) {
        return std::get<std::string>(value_);
    }
    return "";
}

double TXCompactCell::getNumberValue() const {
    if (std::holds_alternative<double>(value_)) {
        return std::get<double>(value_);
    }
    return 0.0;
}

int64_t TXCompactCell::getIntegerValue() const {
    if (std::holds_alternative<int64_t>(value_)) {
        return std::get<int64_t>(value_);
    }
    return 0;
}

bool TXCompactCell::getBooleanValue() const {
    if (std::holds_alternative<bool>(value_)) {
        return std::get<bool>(value_);
    }
    return false;
}

bool TXCompactCell::isFormula() const {
    return hasFormula();
}

const TXFormula* TXCompactCell::getFormulaObject() const {
    if (hasFormula()) {
        return extended_data_->formula.get();
    }
    return nullptr;
}

// ==================== 内存统计 ====================

size_t TXCompactCell::getMemoryUsage() const {
    size_t baseSize = sizeof(TXCompactCell);
    
    if (extended_data_) {
        baseSize += sizeof(ExtendedData);
        
        if (extended_data_->formula) {
            baseSize += sizeof(TXFormula); // 简化计算
        }
        
        if (extended_data_->number_format) {
            baseSize += sizeof(TXNumberFormat); // 简化计算
        }
    }
    
    // 字符串值的额外内存
    if (std::holds_alternative<std::string>(value_)) {
        baseSize += std::get<std::string>(value_).capacity();
    }
    
    return baseSize;
}

double TXCompactCell::getCompactRatio() {
    // 相比原始TXCell的内存节省比例
    // 原始TXCell约80-120字节，TXCompactCell约44字节
    return 0.6; // 约60%的内存节省
}

// ==================== 辅助方法 ====================

void TXCompactCell::ensureExtendedData() {
    if (!extended_data_) {
        extended_data_ = std::make_unique<ExtendedData>();
    }
}

void TXCompactCell::cleanupExtendedData() {
    if (extended_data_ && 
        !extended_data_->formula && 
        !extended_data_->number_format && 
        extended_data_->style_index == 0) {
        extended_data_.reset();
    }
}

TXCompactCell::CellType TXCompactCell::inferType(const cell_value_t& value) {
    return std::visit([](const auto& val) -> CellType {
        using T = std::decay_t<decltype(val)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            return CellType::Empty;
        } else if constexpr (std::is_same_v<T, std::string>) {
            return CellType::String;
        } else if constexpr (std::is_same_v<T, double>) {
            return CellType::Number;
        } else if constexpr (std::is_same_v<T, int64_t>) {
            return CellType::Integer;
        } else if constexpr (std::is_same_v<T, bool>) {
            return CellType::Boolean;
        }
        return CellType::Empty;
    }, value);
}

// ==================== TXCompactCellManager 实现 ====================

TXCompactCellManager::TXCompactCellManager() {
    // 预分配一些空间
    cells_.reserve(1000);
}

// ==================== 单元格访问 ====================

const TXCompactCell* TXCompactCellManager::getCell(const Coordinate& coord) const {
    auto it = cells_.find(coord);
    return (it != cells_.end()) ? &it->second : nullptr;
}

TXCompactCell* TXCompactCellManager::getOrCreateCell(const Coordinate& coord) {
    stats_dirty_ = true;
    return &cells_[coord]; // 使用operator[]自动创建
}

bool TXCompactCellManager::setCellValue(const Coordinate& coord, const CellValue& value) {
    try {
        auto* cell = getOrCreateCell(coord);
        cell->setValue(value);
        return true;
    } catch (const std::exception&) {
        return false;
    }
}

// ==================== 批量操作 ====================

size_t TXCompactCellManager::setCellValues(const std::vector<std::pair<Coordinate, CellValue>>& values) {
    size_t successCount = 0;

    // 预分配空间
    if (cells_.size() + values.size() > cells_.bucket_count() * cells_.max_load_factor()) {
        cells_.reserve(cells_.size() + values.size());
    }

    for (const auto& [coord, value] : values) {
        if (setCellValue(coord, value)) {
            ++successCount;
        }
    }

    return successCount;
}

size_t TXCompactCellManager::setRangeValues(row_t startRow, column_t startCol,
                                          const std::vector<std::vector<CellValue>>& values) {
    if (values.empty()) {
        return 0;
    }

    size_t successCount = 0;
    size_t totalCells = 0;

    // 计算总单元格数
    for (const auto& row : values) {
        totalCells += row.size();
    }

    // 预分配空间
    if (cells_.size() + totalCells > cells_.bucket_count() * cells_.max_load_factor()) {
        cells_.reserve(cells_.size() + totalCells);
    }

    for (size_t rowIdx = 0; rowIdx < values.size(); ++rowIdx) {
        row_t currentRow = row_t(startRow.index() + rowIdx);

        for (size_t colIdx = 0; colIdx < values[rowIdx].size(); ++colIdx) {
            column_t currentCol = column_t(startCol.index() + colIdx);
            Coordinate coord(currentRow, currentCol);

            if (setCellValue(coord, values[rowIdx][colIdx])) {
                ++successCount;
            }
        }
    }

    return successCount;
}

// ==================== 内存管理 ====================

TXCompactCellManager::MemoryStats TXCompactCellManager::getMemoryStats() const {
    if (!stats_dirty_) {
        return cached_stats_;
    }

    MemoryStats stats;
    stats.total_cells = cells_.size();

    for (const auto& [coord, cell] : cells_) {
        size_t cellMemory = cell.getMemoryUsage();
        stats.memory_used += cellMemory;

        // 估算原始TXCell的内存使用（约100字节/单元格）
        stats.memory_saved += (100 - cellMemory);
    }

    if (stats.total_cells > 0) {
        stats.compact_ratio = static_cast<double>(stats.memory_saved) /
                             (stats.memory_used + stats.memory_saved);
    }

    cached_stats_ = stats;
    stats_dirty_ = false;

    return stats;
}

size_t TXCompactCellManager::compactMemory() {
    size_t freedMemory = 0;

    for (auto& [coord, cell] : cells_) {
        // 这里可以实现具体的内存压缩逻辑
        // 例如：清理不需要的扩展数据
        // 暂时返回0，具体实现可以后续添加
    }

    stats_dirty_ = true;
    return freedMemory;
}

void TXCompactCellManager::reserve(size_t expected_cells) {
    cells_.reserve(expected_cells);
}

} // namespace TinaXlsx
