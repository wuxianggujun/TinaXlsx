//
// @file TXSheet.cpp
// @brief 🚀 用户层工作表类实现
//

#include "TinaXlsx/user/TXSheet.hpp"
#include "TinaXlsx/TXCoordUtils.hpp"
#include "TinaXlsx/TXHighPerformanceLogger.hpp"
#include <sstream>
#include <algorithm>
#include <limits>

namespace TinaXlsx {

// ==================== 构造和析构 ====================

TXSheet::TXSheet(
    const std::string& name,
    TXUnifiedMemoryManager& memory_manager,
    TXGlobalStringPool& string_pool
) : sheet_(std::make_unique<TXInMemorySheet>(name, memory_manager, string_pool)) {
    TX_LOG_DEBUG("TXSheet '{}' 创建完成", name);
}

TXSheet::TXSheet(std::unique_ptr<TXInMemorySheet> sheet)
    : sheet_(std::move(sheet)) {
    TX_LOG_DEBUG("TXSheet '{}' 从TXInMemorySheet创建完成", sheet_->getName());
}

TXSheet::~TXSheet() {
    if (sheet_) {
        TX_LOG_DEBUG("TXSheet '{}' 析构", sheet_->getName());
    }
}

TXSheet::TXSheet(TXSheet&& other) noexcept
    : sheet_(std::move(other.sheet_)) {
}

TXSheet& TXSheet::operator=(TXSheet&& other) noexcept {
    if (this != &other) {
        sheet_ = std::move(other.sheet_);
    }
    return *this;
}

// ==================== 单元格访问 ====================

TXCell TXSheet::cell(const std::string& address) {
    return TXCell(*sheet_, address);
}

TXCell TXSheet::cell(uint32_t row, uint32_t col) {
    return TXCell(*sheet_, TXCoordinate{row_t(row + 1), column_t(col + 1)});
}

TXCell TXSheet::cell(const TXCoordinate& coord) {
    return TXCell(*sheet_, coord);
}

// ==================== 范围操作 ====================

TXRange TXSheet::range(const std::string& range_address) {
    auto result = parseRangeAddress(range_address);
    if (result.isOk()) {
        return result.value();
    } else {
        handleError("解析范围地址", result.error());
        // 返回无效范围
        return TXRange();
    }
}

TXRange TXSheet::range(uint32_t start_row, uint32_t start_col,
                      uint32_t end_row, uint32_t end_col) {
    TXCoordinate start{row_t(start_row + 1), column_t(start_col + 1)};
    TXCoordinate end{row_t(end_row + 1), column_t(end_col + 1)};
    return TXRange(start, end);
}

TXRange TXSheet::range(const TXCoordinate& start, const TXCoordinate& end) {
    return TXRange(start, end);
}

// ==================== 批量数据操作 ====================

TXResult<void> TXSheet::setValues(const std::string& range_address,
                                 const TXVector<TXVector<TXVariant>>& values) {
    auto range_result = parseRangeAddress(range_address);
    if (range_result.isError()) {
        return TXResult<void>(range_result.error());
    }
    
    auto range = range_result.value();
    auto start = range.getStart();
    
    try {
        // 准备坐标和值数组
        std::vector<TXCoordinate> coords;
        std::vector<TXVariant> variants;
        
        for (size_t row = 0; row < values.size(); ++row) {
            for (size_t col = 0; col < values[row].size(); ++col) {
                uint32_t target_row = start.getRow().index() + static_cast<uint32_t>(row) - 1;
                uint32_t target_col = start.getCol().index() + static_cast<uint32_t>(col) - 1;
                
                coords.emplace_back(row_t(target_row + 1), column_t(target_col + 1));
                variants.push_back(values[row][col]);
            }
        }
        
        // 使用底层批量设置
        auto result = sheet_->setBatchMixed(coords, variants);
        if (result.isError()) {
            return TXResult<void>(result.error());
        }
        
        TX_LOG_DEBUG("批量设置值完成: {} 个单元格", coords.size());
        return TXResult<void>();
        
    } catch (const std::exception& e) {
        return TXResult<void>(TXError(TXErrorCode::InvalidOperation, 
                                     std::string("批量设置值失败: ") + e.what()));
    }
}

TXResult<TXVector<TXVector<TXVariant>>> TXSheet::getValues(const std::string& range_address) {
    auto range_result = parseRangeAddress(range_address);
    if (range_result.isError()) {
        return TXResult<TXVector<TXVector<TXVariant>>>(range_result.error());
    }

    auto range = range_result.value();
    auto start = range.getStart();
    auto end = range.getEnd();

    try {
        auto& memory_manager = GlobalUnifiedMemoryManager::getInstance();
        TXVector<TXVector<TXVariant>> result(memory_manager);

        uint32_t rows = end.getRow().index() - start.getRow().index() + 1;
        uint32_t cols = end.getCol().index() - start.getCol().index() + 1;

        // 手动构造每一行，避免resize()的默认构造问题
        result.reserve(rows);
        for (uint32_t i = 0; i < rows; ++i) {
            result.emplace_back(memory_manager, cols, TXVariant());
        }

        // 逐个获取单元格值
        for (uint32_t row = 0; row < rows; ++row) {
            for (uint32_t col = 0; col < cols; ++col) {
                uint32_t target_row = start.getRow().index() + row - 1;
                uint32_t target_col = start.getCol().index() + col - 1;

                TXCoordinate coord{row_t(target_row + 1), column_t(target_col + 1)};
                auto value_result = sheet_->getValue(coord);

                if (value_result.isOk()) {
                    result[row][col] = value_result.value();
                } else {
                    result[row][col] = TXVariant(); // 空值
                }
            }
        }

        TX_LOG_DEBUG("批量获取值完成: {}x{} 范围", rows, cols);
        return TXResult<TXVector<TXVector<TXVariant>>>(std::move(result));

    } catch (const std::exception& e) {
        return TXResult<TXVector<TXVector<TXVariant>>>(
            TXError(TXErrorCode::InvalidOperation,
                   std::string("批量获取值失败: ") + e.what()));
    }
}

TXResult<void> TXSheet::fillRange(const std::string& range_address, const TXVariant& value) {
    auto range_result = parseRangeAddress(range_address);
    if (range_result.isError()) {
        return TXResult<void>(range_result.error());
    }

    auto range = range_result.value();

    // 根据值类型选择合适的批量操作
    if (value.getType() == TXVariant::Type::Number) {
        auto result = sheet_->fillRange(range, value.getNumber());
        if (result.isOk()) {
            TX_LOG_DEBUG("数值填充完成: {} 个单元格", result.value());
            return TXResult<void>();
        } else {
            return TXResult<void>(result.error());
        }
    } else if (value.getType() == TXVariant::Type::String) {
        // 使用字符串版本的fillRange
        auto result = sheet_->fillRange(range, value.getString());
        if (result.isOk()) {
            TX_LOG_DEBUG("字符串填充完成: {} 个单元格", result.value());
            return TXResult<void>();
        } else {
            return TXResult<void>(result.error());
        }
    } else {
        // 对于其他类型，使用通用方法
        auto start = range.getStart();
        auto end = range.getEnd();

        std::vector<TXCoordinate> coords;
        std::vector<TXVariant> values;

        for (uint32_t row = start.getRow().index(); row <= end.getRow().index(); ++row) {
            for (uint32_t col = start.getCol().index(); col <= end.getCol().index(); ++col) {
                coords.emplace_back(row_t(row), column_t(col));
                values.push_back(value);
            }
        }

        auto result = sheet_->setBatchMixed(coords, values);
        if (result.isOk()) {
            TX_LOG_DEBUG("填充完成: {} 个单元格", result.value());
            return TXResult<void>();
        } else {
            return TXResult<void>(result.error());
        }
    }
}

TXResult<void> TXSheet::clearRange(const std::string& range_address) {
    auto range_result = parseRangeAddress(range_address);
    if (range_result.isError()) {
        return TXResult<void>(range_result.error());
    }
    
    auto range = range_result.value();
    auto result = sheet_->clearRange(range);
    
    if (result.isOk()) {
        TX_LOG_DEBUG("清除范围完成: {} 个单元格", result.value());
    }
    
    if (result.isError()) {
        return TXResult<void>(result.error());
    } else {
        return TXResult<void>(); // 成功的TXResult<void>
    }
}

// ==================== 工作表属性 ====================

const std::string& TXSheet::getName() const {
    return sheet_->getName();
}

void TXSheet::setName(const std::string& name) {
    sheet_->setName(name);
    TX_LOG_DEBUG("工作表名称已更改为: {}", name);
}

TXRange TXSheet::getUsedRange() const {
    return sheet_->getUsedRange();
}

size_t TXSheet::getCellCount() const {
    return sheet_->getCellCount();
}

bool TXSheet::isEmpty() const {
    return getCellCount() == 0;
}

// ==================== 性能优化 ====================

void TXSheet::reserve(size_t estimated_cells) {
    sheet_->reserve(estimated_cells);
    TX_LOG_DEBUG("预分配内存: {} 个单元格", estimated_cells);
}

void TXSheet::optimize() {
    sheet_->optimizeMemoryLayout();
    TX_LOG_DEBUG("内存布局优化完成");
}

size_t TXSheet::compress() {
    size_t compressed = sheet_->compressSparseData();
    TX_LOG_DEBUG("压缩稀疏数据: {} 个单元格", compressed);
    return compressed;
}

void TXSheet::shrinkToFit() {
    sheet_->shrink_to_fit();
    TX_LOG_DEBUG("内存收缩完成");
}

// ==================== 查找和统计 ====================

TXVector<TXCoordinate> TXSheet::findValue(const TXVariant& value,
                                         const std::string& range_address) {
    auto& memory_manager = GlobalUnifiedMemoryManager::getInstance();

    if (range_address.empty()) {
        // 在整个工作表中查找
        if (value.getType() == TXVariant::Type::Number) {
            auto std_result = sheet_->findValue(value.getNumber());
            // 转换std::vector到TXVector
            TXVector<TXCoordinate> result(memory_manager);
            result.reserve(std_result.size());
            for (const auto& coord : std_result) {
                result.push_back(coord);
            }
            return result;
        } else if (value.getType() == TXVariant::Type::String) {
            auto std_result = sheet_->findString(value.getString());
            // 转换std::vector到TXVector
            TXVector<TXCoordinate> result(memory_manager);
            result.reserve(std_result.size());
            for (const auto& coord : std_result) {
                result.push_back(coord);
            }
            return result;
        }
        return TXVector<TXCoordinate>(memory_manager);
    } else {
        // 在指定范围中查找
        auto range_result = parseRangeAddress(range_address);
        if (range_result.isError()) {
            handleError("解析查找范围", range_result.error());
            return TXVector<TXCoordinate>(memory_manager);
        }

        auto range = range_result.value();
        if (value.getType() == TXVariant::Type::Number) {
            auto std_result = sheet_->findValue(value.getNumber(), &range);
            // 转换std::vector到TXVector
            TXVector<TXCoordinate> result(memory_manager);
            result.reserve(std_result.size());
            for (const auto& coord : std_result) {
                result.push_back(coord);
            }
            return result;
        } else if (value.getType() == TXVariant::Type::String) {
            auto std_result = sheet_->findString(value.getString(), &range);
            // 转换std::vector到TXVector
            TXVector<TXCoordinate> result(memory_manager);
            result.reserve(std_result.size());
            for (const auto& coord : std_result) {
                result.push_back(coord);
            }
            return result;
        }
        return TXVector<TXCoordinate>(memory_manager);
    }
}

TXResult<double> TXSheet::sum(const std::string& range_address) {
    auto range_result = parseRangeAddress(range_address);
    if (range_result.isError()) {
        return TXResult<double>(range_result.error());
    }

    auto range = range_result.value();
    auto result = sheet_->sum(range);

    if (result.isOk()) {
        TX_LOG_DEBUG("求和完成: 范围={}, 结果={}", range_address, result.value());
    }

    return result;
}

TXResult<double> TXSheet::average(const std::string& range_address) {
    auto sum_result = sum(range_address);
    if (sum_result.isError()) {
        return sum_result;
    }

    auto range_result = parseRangeAddress(range_address);
    if (range_result.isError()) {
        return TXResult<double>(range_result.error());
    }

    auto range = range_result.value();
    uint64_t cell_count = range.getCellCount();

    if (cell_count == 0) {
        return TXResult<double>(TXError(TXErrorCode::InvalidOperation, "范围为空"));
    }

    double avg = sum_result.value() / static_cast<double>(cell_count);
    TX_LOG_DEBUG("平均值计算完成: 范围={}, 结果={}", range_address, avg);

    return TXResult<double>(avg);
}

TXResult<double> TXSheet::max(const std::string& range_address) {
    auto range_result = parseRangeAddress(range_address);
    if (range_result.isError()) {
        return TXResult<double>(range_result.error());
    }

    auto range = range_result.value();

    // 遍历范围找最大值
    double max_value = std::numeric_limits<double>::lowest();
    bool found_number = false;

    auto start = range.getStart();
    auto end = range.getEnd();

    for (uint32_t row = start.getRow().index(); row <= end.getRow().index(); ++row) {
        for (uint32_t col = start.getCol().index(); col <= end.getCol().index(); ++col) {
            TXCoordinate coord{row_t(row), column_t(col)};
            auto value_result = sheet_->getValue(coord);

            if (value_result.isOk() && value_result.value().getType() == TXVariant::Type::Number) {
                double val = value_result.value().getNumber();
                if (!found_number || val > max_value) {
                    max_value = val;
                    found_number = true;
                }
            }
        }
    }

    if (!found_number) {
        return TXResult<double>(TXError(TXErrorCode::InvalidOperation, "范围中没有数值"));
    }

    TX_LOG_DEBUG("最大值计算完成: 范围={}, 结果={}", range_address, max_value);
    return TXResult<double>(max_value);
}

TXResult<double> TXSheet::min(const std::string& range_address) {
    auto range_result = parseRangeAddress(range_address);
    if (range_result.isError()) {
        return TXResult<double>(range_result.error());
    }

    auto range = range_result.value();

    // 遍历范围找最小值
    double min_value = std::numeric_limits<double>::max();
    bool found_number = false;

    auto start = range.getStart();
    auto end = range.getEnd();

    for (uint32_t row = start.getRow().index(); row <= end.getRow().index(); ++row) {
        for (uint32_t col = start.getCol().index(); col <= end.getCol().index(); ++col) {
            TXCoordinate coord{row_t(row), column_t(col)};
            auto value_result = sheet_->getValue(coord);

            if (value_result.isOk() && value_result.value().getType() == TXVariant::Type::Number) {
                double val = value_result.value().getNumber();
                if (!found_number || val < min_value) {
                    min_value = val;
                    found_number = true;
                }
            }
        }
    }

    if (!found_number) {
        return TXResult<double>(TXError(TXErrorCode::InvalidOperation, "范围中没有数值"));
    }

    TX_LOG_DEBUG("最小值计算完成: 范围={}, 结果={}", range_address, min_value);
    return TXResult<double>(min_value);
}

// ==================== 调试和诊断 ====================

std::string TXSheet::toString() const {
    std::ostringstream oss;
    oss << "TXSheet{";
    oss << "名称=\"" << getName() << "\"";
    oss << ", 单元格数=" << getCellCount();
    oss << ", 使用范围=" << getUsedRange().toAddress();
    oss << ", 是否为空=" << (isEmpty() ? "是" : "否");
    oss << "}";
    return oss.str();
}

bool TXSheet::isValid() const {
    return sheet_ != nullptr && !getName().empty();
}

std::string TXSheet::getPerformanceStats() const {
    auto stats = sheet_->getPerformanceStats();
    std::ostringstream oss;
    oss << "TXSheet性能统计:\n";
    oss << "  总单元格数: " << stats.total_cells << "\n";
    oss << "  批量操作次数: " << stats.batch_operations << "\n";
    oss << "  平均操作时间: " << stats.avg_operation_time << "ms\n";
    oss << "  缓存命中率: " << (stats.cache_hit_ratio * 100) << "%\n";
    oss << "  内存使用量: " << (stats.memory_usage / 1024.0 / 1024.0) << "MB\n";
    oss << "  压缩比: " << (stats.compression_ratio * 100) << "%";
    return oss.str();
}

// ==================== 内部辅助方法 ====================

void TXSheet::handleError(const std::string& operation, const TXError& error) const {
    TX_LOG_WARN("TXSheet操作失败: {} - 工作表={}, 错误={}",
                operation, getName(), error.getMessage());
}

TXResult<TXRange> TXSheet::parseRangeAddress(const std::string& range_address) {
    auto result = TXCoordUtils::parseRange(range_address);
    if (result.isOk()) {
        auto [start, end] = result.value();
        return TXResult<TXRange>(TXRange(start, end));
    } else {
        return TXResult<TXRange>(result.error());
    }
}

} // namespace TinaXlsx
