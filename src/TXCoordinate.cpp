#include "TinaXlsx/TXCoordinate.hpp"
#include <algorithm>
#include <cctype>
#include <regex>
#include <sstream>

namespace TinaXlsx {

// ==================== TXCoordinate 构造函数实现 ====================

TXCoordinate::TXCoordinate(const std::string& address) {
    // 使用正则表达式解析地址
    std::regex addr_regex(R"(([A-Z]+)(\d+))");
    std::smatch match;
    
    if (std::regex_match(address, match, addr_regex)) {
        std::string col_str = match[1].str();
        std::string row_str = match[2].str();
        
        col_ = column_t(col_str);
        try {
            u32 row_index = static_cast<u32>(std::stoul(row_str));
            row_ = row_t(row_index);
        } catch (const std::exception&) {
            // 解析失败，创建无效坐标
            row_ = row_t(0);
            col_ = column_t(static_cast<u32>(0));
        }
    } else {
        // 解析失败，创建无效坐标
        row_ = row_t(0);
        col_ = column_t(static_cast<u32>(0));
    }
}

// ==================== TXCoordinate 设置器实现 ====================

TXCoordinate& TXCoordinate::setRow(const row_t& row) {
    row_ = row;
    return *this;
}

TXCoordinate& TXCoordinate::setCol(const column_t& col) {
    col_ = col;
    return *this;
}

TXCoordinate& TXCoordinate::set(const row_t& row, const column_t& col) {
    row_ = row;
    col_ = col;
    return *this;
}

// ==================== TXCoordinate 验证方法实现 ====================

bool TXCoordinate::isValid() const {
    return row_.is_valid() && col_.is_valid();
}

bool TXCoordinate::isValidRow() const {
    return row_.is_valid();
}

bool TXCoordinate::isValidCol() const {
    return col_.is_valid();
}

// ==================== TXCoordinate 转换方法实现 ====================

std::string TXCoordinate::toAddress() const {
    return col_.column_string() + std::to_string(row_.index());
}

std::string TXCoordinate::getColName() const {
    return col_.column_string();
}

// ==================== TXCoordinate 偏移操作实现 ====================

TXCoordinate TXCoordinate::offsetRow(int offset) const {
    u32 newRow = static_cast<u32>(
        std::max(1, static_cast<int>(row_.index()) + offset)
    );
    return TXCoordinate(row_t(newRow), col_);
}

TXCoordinate TXCoordinate::offsetCol(int offset) const {
    u32 newCol = static_cast<u32>(
        std::max(1, static_cast<int>(col_.index()) + offset)
    );
    return TXCoordinate(row_, column_t(newCol));
}

TXCoordinate TXCoordinate::offset(int row_offset, int col_offset) const {
    u32 newRow = static_cast<u32>(
        std::max(1, static_cast<int>(row_.index()) + row_offset)
    );
    u32 newCol = static_cast<u32>(
        std::max(1, static_cast<int>(col_.index()) + col_offset)
    );
    return TXCoordinate(row_t(newRow), column_t(newCol));
}

// ==================== TXCoordinate 比较操作实现 ====================

bool TXCoordinate::operator==(const TXCoordinate& other) const {
    return row_ == other.row_ && col_ == other.col_;
}

bool TXCoordinate::operator<(const TXCoordinate& other) const {
    if (row_ != other.row_) {
        return row_ < other.row_;
    }
    return col_ < other.col_;
}

bool TXCoordinate::operator<=(const TXCoordinate& other) const {
    return *this < other || *this == other;
}

bool TXCoordinate::operator>(const TXCoordinate& other) const {
    return !(*this <= other);
}

bool TXCoordinate::operator>=(const TXCoordinate& other) const {
    return !(*this < other);
}

// ==================== TXCoordinate 运算符重载实现 ====================

TXCoordinate TXCoordinate::operator+(const TXCoordinate& other) const {
    return offset(other.row_.index() - 1, other.col_.index() - 1); // 减1因为坐标是1-based
}

TXCoordinate TXCoordinate::operator-(const TXCoordinate& other) const {
    return offset(-(static_cast<int>(other.row_.index()) - 1), -(static_cast<int>(other.col_.index()) - 1));
}

TXCoordinate& TXCoordinate::operator+=(const TXCoordinate& other) {
    *this = *this + other;
    return *this;
}

TXCoordinate& TXCoordinate::operator-=(const TXCoordinate& other) {
    *this = *this - other;
    return *this;
}

// ==================== TXCoordinate 静态工厂方法实现 ====================

TXCoordinate TXCoordinate::fromAddress(const std::string& address) {
    return TXCoordinate(address);
}

TXCoordinate TXCoordinate::fromColNameRow(const std::string& col_name, const row_t& row) {
    column_t col(col_name);
    return TXCoordinate(row, col);
}

// ==================== TXCoordinate 静态工具方法实现 ====================

bool TXCoordinate::isValidCoordinate(const row_t& row, const column_t& col) {
    return row.is_valid() && col.is_valid();
}

} // namespace TinaXlsx 