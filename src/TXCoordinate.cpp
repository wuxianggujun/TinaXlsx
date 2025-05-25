#include "TinaXlsx/TXCoordinate.hpp"
#include <algorithm>
#include <cctype>
#include <regex>
#include <sstream>

namespace TinaXlsx {

// ==================== TXCoordinate 构造函数实现 ====================

TXCoordinate::TXCoordinate(const std::string& address) {
    auto [row, col] = addressToCoordinate(address);
    row_ = row;
    col_ = col;
}

// ==================== TXCoordinate 设置器实现 ====================

TXCoordinate& TXCoordinate::setRow(TXTypes::RowIndex row) {
    row_ = row;
    return *this;
}

TXCoordinate& TXCoordinate::setCol(TXTypes::ColIndex col) {
    col_ = col;
    return *this;
}

TXCoordinate& TXCoordinate::set(TXTypes::RowIndex row, TXTypes::ColIndex col) {
    row_ = row;
    col_ = col;
    return *this;
}

TXCoordinate& TXCoordinate::setFromAddress(const std::string& address) {
    auto [row, col] = addressToCoordinate(address);
    row_ = row;
    col_ = col;
    return *this;
}

// ==================== TXCoordinate 验证方法实现 ====================

bool TXCoordinate::isValid() const {
    return isValidCoordinate(row_, col_);
}

bool TXCoordinate::isValidRow() const {
    return TXTypes::isValidRow(row_);
}

bool TXCoordinate::isValidCol() const {
    return TXTypes::isValidCol(col_);
}

// ==================== TXCoordinate 转换方法实现 ====================

std::string TXCoordinate::toAddress() const {
    return coordinateToAddress(row_, col_);
}

std::string TXCoordinate::getColName() const {
    return colIndexToName(col_);
}

// ==================== TXCoordinate 偏移操作实现 ====================

TXCoordinate TXCoordinate::offsetRow(int offset) const {
    TXTypes::RowIndex newRow = static_cast<TXTypes::RowIndex>(
        std::max(1, static_cast<int>(row_) + offset)
    );
    return TXCoordinate(newRow, col_);
}

TXCoordinate TXCoordinate::offsetCol(int offset) const {
    TXTypes::ColIndex newCol = static_cast<TXTypes::ColIndex>(
        std::max(1, static_cast<int>(col_) + offset)
    );
    return TXCoordinate(row_, newCol);
}

TXCoordinate TXCoordinate::offset(int row_offset, int col_offset) const {
    TXTypes::RowIndex newRow = static_cast<TXTypes::RowIndex>(
        std::max(1, static_cast<int>(row_) + row_offset)
    );
    TXTypes::ColIndex newCol = static_cast<TXTypes::ColIndex>(
        std::max(1, static_cast<int>(col_) + col_offset)
    );
    return TXCoordinate(newRow, newCol);
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
    return offset(other.row_ - 1, other.col_ - 1); // 减1因为坐标是1-based
}

TXCoordinate TXCoordinate::operator-(const TXCoordinate& other) const {
    return offset(-(static_cast<int>(other.row_) - 1), -(static_cast<int>(other.col_) - 1));
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

TXCoordinate TXCoordinate::fromRowCol(TXTypes::RowIndex row, TXTypes::ColIndex col) {
    return TXCoordinate(row, col);
}

TXCoordinate TXCoordinate::fromColNameRow(const std::string& col_name, TXTypes::RowIndex row) {
    TXTypes::ColIndex col = colNameToIndex(col_name);
    return TXCoordinate(row, col);
}

// ==================== TXCoordinate 静态工具方法实现 ====================

std::string TXCoordinate::colIndexToName(TXTypes::ColIndex col) {
    if (col == 0 || col > TXTypes::MAX_COLS) {
        return "";
    }
    
    std::string result;
    while (col > 0) {
        col--; // 转换为0-based
        result = static_cast<char>('A' + (col % 26)) + result;
        col /= 26;
    }
    return result;
}

TXTypes::ColIndex TXCoordinate::colNameToIndex(const std::string& name) {
    if (name.empty()) {
        return TXTypes::INVALID_COL;
    }
    
    TXTypes::ColIndex result = 0;
    for (char c : name) {
        if (!std::isalpha(c)) {
            return TXTypes::INVALID_COL;
        }
        c = std::toupper(c);
        if (c < 'A' || c > 'Z') {
            return TXTypes::INVALID_COL;
        }
        result = result * 26 + (c - 'A' + 1);
    }
    
    return (result <= TXTypes::MAX_COLS) ? result : TXTypes::INVALID_COL;
}

std::string TXCoordinate::coordinateToAddress(TXTypes::RowIndex row, TXTypes::ColIndex col) {
    if (!isValidCoordinate(row, col)) {
        return "";
    }
    
    return colIndexToName(col) + std::to_string(row);
}

std::pair<TXTypes::RowIndex, TXTypes::ColIndex> TXCoordinate::addressToCoordinate(const std::string& address) {
    if (address.empty()) {
        return {TXTypes::INVALID_ROW, TXTypes::INVALID_COL};
    }
    
    // 分离字母和数字部分
    std::size_t i = 0;
    while (i < address.size() && std::isalpha(address[i])) {
        ++i;
    }
    
    if (i == 0 || i == address.size()) {
        return {TXTypes::INVALID_ROW, TXTypes::INVALID_COL};
    }
    
    std::string colPart = address.substr(0, i);
    std::string rowPart = address.substr(i);
    
    // 验证数字部分
    for (char c : rowPart) {
        if (!std::isdigit(c)) {
            return {TXTypes::INVALID_ROW, TXTypes::INVALID_COL};
        }
    }
    
    TXTypes::ColIndex col = colNameToIndex(colPart);
    TXTypes::RowIndex row = 0;
    
    try {
        unsigned long rowLong = std::stoul(rowPart);
        if (rowLong > TXTypes::MAX_ROWS) {
            return {TXTypes::INVALID_ROW, TXTypes::INVALID_COL};
        }
        row = static_cast<TXTypes::RowIndex>(rowLong);
    } catch (...) {
        return {TXTypes::INVALID_ROW, TXTypes::INVALID_COL};
    }
    
    if (row == 0 || col == TXTypes::INVALID_COL) {
        return {TXTypes::INVALID_ROW, TXTypes::INVALID_COL};
    }
    
    return {row, col};
}

bool TXCoordinate::isValidCoordinate(TXTypes::RowIndex row, TXTypes::ColIndex col) {
    return TXTypes::isValidRow(row) && TXTypes::isValidCol(col);
}

// ==================== TXRange 实现 ====================

TXRange::TXRange(const TXCoordinate& start, const TXCoordinate& end) 
    : start_(start), end_(end) {
    normalize();
}

TXRange::TXRange(TXTypes::RowIndex start_row, TXTypes::ColIndex start_col,
                 TXTypes::RowIndex end_row, TXTypes::ColIndex end_col)
    : start_(start_row, start_col), end_(end_row, end_col) {
    normalize();
}

TXRange::TXRange(const std::string& range_address) {
    *this = fromAddress(range_address);
}

// ==================== TXRange 访问器实现 ====================

TXTypes::RowIndex TXRange::getRowCount() const {
    return end_.getRow() - start_.getRow() + 1;
}

TXTypes::ColIndex TXRange::getColCount() const {
    return end_.getCol() - start_.getCol() + 1;
}

uint64_t TXRange::getCellCount() const {
    return static_cast<uint64_t>(getRowCount()) * static_cast<uint64_t>(getColCount());
}

// ==================== TXRange 设置器实现 ====================

TXRange& TXRange::setStart(const TXCoordinate& start) {
    start_ = start;
    normalize();
    return *this;
}

TXRange& TXRange::setEnd(const TXCoordinate& end) {
    end_ = end;
    normalize();
    return *this;
}

TXRange& TXRange::set(const TXCoordinate& start, const TXCoordinate& end) {
    start_ = start;
    end_ = end;
    normalize();
    return *this;
}

// ==================== TXRange 验证和操作实现 ====================

bool TXRange::isValid() const {
    return start_.isValid() && end_.isValid() && start_ <= end_;
}

bool TXRange::contains(const TXCoordinate& coord) const {
    return coord.getRow() >= start_.getRow() && coord.getRow() <= end_.getRow() &&
           coord.getCol() >= start_.getCol() && coord.getCol() <= end_.getCol();
}

bool TXRange::contains(const TXRange& other) const {
    return contains(other.start_) && contains(other.end_);
}

bool TXRange::intersects(const TXRange& other) const {
    return !(end_.getRow() < other.start_.getRow() || 
             start_.getRow() > other.end_.getRow() ||
             end_.getCol() < other.start_.getCol() || 
             start_.getCol() > other.end_.getCol());
}

TXRange TXRange::intersection(const TXRange& other) const {
    if (!intersects(other)) {
        return TXRange(); // 返回无效范围
    }
    
    TXTypes::RowIndex start_row = std::max(start_.getRow(), other.start_.getRow());
    TXTypes::ColIndex start_col = std::max(start_.getCol(), other.start_.getCol());
    TXTypes::RowIndex end_row = std::min(end_.getRow(), other.end_.getRow());
    TXTypes::ColIndex end_col = std::min(end_.getCol(), other.end_.getCol());
    
    return TXRange(start_row, start_col, end_row, end_col);
}

TXRange& TXRange::expand(const TXCoordinate& coord) {
    start_ = TXCoordinate(
        std::min(start_.getRow(), coord.getRow()),
        std::min(start_.getCol(), coord.getCol())
    );
    end_ = TXCoordinate(
        std::max(end_.getRow(), coord.getRow()),
        std::max(end_.getCol(), coord.getCol())
    );
    return *this;
}

TXRange& TXRange::expand(const TXRange& other) {
    expand(other.start_);
    expand(other.end_);
    return *this;
}

// ==================== TXRange 转换方法实现 ====================

std::string TXRange::toAddress() const {
    if (start_ == end_) {
        return start_.toAddress();
    }
    return start_.toAddress() + ":" + end_.toAddress();
}

std::vector<TXCoordinate> TXRange::getAllCoordinates() const {
    std::vector<TXCoordinate> coords;
    coords.reserve(getCellCount());
    
    for (TXTypes::RowIndex row = start_.getRow(); row <= end_.getRow(); ++row) {
        for (TXTypes::ColIndex col = start_.getCol(); col <= end_.getCol(); ++col) {
            coords.emplace_back(row, col);
        }
    }
    
    return coords;
}

// ==================== TXRange 比较操作实现 ====================

bool TXRange::operator==(const TXRange& other) const {
    return start_ == other.start_ && end_ == other.end_;
}

// ==================== TXRange 静态工厂方法实现 ====================

TXRange TXRange::fromAddress(const std::string& range_address) {
    // 查找冒号分隔符
    std::size_t colon_pos = range_address.find(':');
    
    if (colon_pos == std::string::npos) {
        // 单个单元格
        TXCoordinate coord(range_address);
        return TXRange(coord, coord);
    }
    
    std::string start_addr = range_address.substr(0, colon_pos);
    std::string end_addr = range_address.substr(colon_pos + 1);
    
    TXCoordinate start(start_addr);
    TXCoordinate end(end_addr);
    
    return TXRange(start, end);
}

TXRange TXRange::singleCell(const TXCoordinate& coord) {
    return TXRange(coord, coord);
}

TXRange TXRange::entireRow(TXTypes::RowIndex row) {
    return TXRange(row, 1, row, TXTypes::MAX_COLS);
}

TXRange TXRange::entireCol(TXTypes::ColIndex col) {
    return TXRange(1, col, TXTypes::MAX_ROWS, col);
}

// ==================== TXRange 辅助方法实现 ====================

void TXRange::normalize() {
    if (start_.getRow() > end_.getRow() || start_.getCol() > end_.getCol()) {
        TXTypes::RowIndex min_row = std::min(start_.getRow(), end_.getRow());
        TXTypes::RowIndex max_row = std::max(start_.getRow(), end_.getRow());
        TXTypes::ColIndex min_col = std::min(start_.getCol(), end_.getCol());
        TXTypes::ColIndex max_col = std::max(start_.getCol(), end_.getCol());
        
        start_.set(min_row, min_col);
        end_.set(max_row, max_col);
    }
}

} // namespace TinaXlsx 