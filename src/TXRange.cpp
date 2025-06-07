#include "TinaXlsx/TXRange.hpp"
#include "TinaXlsx/TXCoordUtils.hpp"  // 🚀 使用统一的坐标转换工具
#include <algorithm>
#include <sstream>

namespace TinaXlsx {

// ==================== TXRange 构造函数实现 ====================

TXRange::TXRange() : start_(row_t(1), column_t(1)), end_(row_t(1), column_t(1)) {}

TXRange::TXRange(const TXCoordinate& start, const TXCoordinate& end)
    : start_(start), end_(end) {
    normalize();
}

TXRange::TXRange(const std::string& range_address) {
    // 🚀 使用统一的TXCoordUtils解析范围地址
    auto colon_pos = range_address.find(':');
    if (colon_pos == std::string::npos) {
        // 单个单元格
        auto result = TXCoordUtils::parseCoord(range_address);
        if (result.isOk()) {
            start_ = result.value();
            end_ = result.value();
        } else {
            // 解析失败，创建无效范围
            start_ = TXCoordinate(row_t(static_cast<uint32_t>(0)), column_t(static_cast<uint32_t>(0)));
            end_ = TXCoordinate(row_t(static_cast<uint32_t>(0)), column_t(static_cast<uint32_t>(0)));
        }
    } else {
        // 使用TXCoordUtils解析范围
        auto range_result = TXCoordUtils::parseRange(range_address);
        if (range_result.isOk()) {
            auto [start, end] = range_result.value();
            start_ = start;
            end_ = end;
            normalize();
        } else {
            // 解析失败，创建无效范围
            start_ = TXCoordinate(row_t(static_cast<uint32_t>(0)), column_t(static_cast<uint32_t>(0)));
            end_ = TXCoordinate(row_t(static_cast<uint32_t>(0)), column_t(static_cast<uint32_t>(0)));
        }
    }
}

// ==================== TXRange 访问器实现 ====================

row_t TXRange::getRowCount() const {
    return row_t(end_.getRow().index() - start_.getRow().index() + 1);
}

column_t TXRange::getColCount() const {
    return column_t(end_.getCol().index() - start_.getCol().index() + 1);
}

uint64_t TXRange::getCellCount() const {
    return static_cast<uint64_t>(getRowCount().index()) * static_cast<uint64_t>(getColCount().index());
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
    return start_.isValid() && end_.isValid() &&
           start_.getRow().index() <= end_.getRow().index() &&
           start_.getCol().index() <= end_.getCol().index();
}

bool TXRange::isEmpty() const {
    // 范围为空的条件：无效或者起始和结束坐标都是(0,0)
    return !isValid() ||
           (start_.getRow().index() == 0 && start_.getCol().index() == 0 &&
            end_.getRow().index() == 0 && end_.getCol().index() == 0);
}

bool TXRange::contains(const TXCoordinate& coord) const {
    return coord.getRow().index() >= start_.getRow().index() && 
           coord.getRow().index() <= end_.getRow().index() &&
           coord.getCol().index() >= start_.getCol().index() && 
           coord.getCol().index() <= end_.getCol().index();
}

bool TXRange::contains(const TXRange& other) const {
    return contains(other.start_) && contains(other.end_);
}

bool TXRange::intersects(const TXRange& other) const {
    return !(end_.getRow().index() < other.start_.getRow().index() || 
             start_.getRow().index() > other.end_.getRow().index() ||
             end_.getCol().index() < other.start_.getCol().index() || 
             start_.getCol().index() > other.end_.getCol().index());
}

TXRange TXRange::intersection(const TXRange& other) const {
    if (!intersects(other)) {
        return TXRange(); // 无效范围
    }
    
    u32 start_row = std::max(start_.getRow().index(), other.start_.getRow().index());
    u32 start_col = std::max(start_.getCol().index(), other.start_.getCol().index());
    u32 end_row = std::min(end_.getRow().index(), other.end_.getRow().index());
    u32 end_col = std::min(end_.getCol().index(), other.end_.getCol().index());
    
    return TXRange(TXCoordinate(row_t(start_row), column_t(start_col)), 
                   TXCoordinate(row_t(end_row), column_t(end_col)));
}

TXRange& TXRange::expand(const TXCoordinate& coord) {
    if (!start_.isValid()) {
        start_ = coord;
        end_ = coord;
    } else {
        u32 start_row = std::min(start_.getRow().index(), coord.getRow().index());
        u32 start_col = std::min(start_.getCol().index(), coord.getCol().index());
        u32 end_row = std::max(end_.getRow().index(), coord.getRow().index());
        u32 end_col = std::max(end_.getCol().index(), coord.getCol().index());
        
        start_ = TXCoordinate(row_t(start_row), column_t(start_col));
        end_ = TXCoordinate(row_t(end_row), column_t(end_col));
    }
    return *this;
}

TXRange& TXRange::expand(const TXRange& other) {
    if (!other.isValid()) {
        return *this;
    }
    
    if (!isValid()) {
        *this = other;
    } else {
        expand(other.start_);
        expand(other.end_);
    }
    return *this;
}

// ==================== TXRange 转换方法实现 ====================

std::string TXRange::toAddress() const {
    if (start_ == end_) {
        return start_.toAddress();
    }
    return start_.toAddress() + ":" + end_.toAddress();
}

std::string TXRange::toAbsoluteAddress() const {
    // 需要先检查TXCoordinate是否有toAbsoluteAddress方法
    // 如果没有，我们手动添加$符号
    std::string startAddr = start_.toAddress();
    std::string endAddr = end_.toAddress();

    // 为地址添加$符号
    auto addDollarSigns = [](const std::string& addr) -> std::string {
        std::string result;
        bool foundLetter = false;
        for (char c : addr) {
            if (std::isalpha(c) && !foundLetter) {
                result += "$";
                foundLetter = true;
            } else if (std::isdigit(c) && foundLetter) {
                result += "$";
                foundLetter = false;
            }
            result += c;
        }
        return result;
    };

    std::string absStart = addDollarSigns(startAddr);
    std::string absEnd = addDollarSigns(endAddr);

    if (start_ == end_) {
        return absStart;
    }
    return absStart + ":" + absEnd;
}

std::vector<TXCoordinate> TXRange::getAllCoordinates() const {
    std::vector<TXCoordinate> coords;
    
    for (u32 row = start_.getRow().index(); row <= end_.getRow().index(); ++row) {
        for (u32 col = start_.getCol().index(); col <= end_.getCol().index(); ++col) {
            coords.emplace_back(row_t(row), column_t(col));
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
    return TXRange(range_address);
}

TXRange TXRange::singleCell(const TXCoordinate& coord) {
    return TXRange(coord, coord);
}

TXRange TXRange::entireRow(const row_t& row) {
    return TXRange(TXCoordinate(row, column_t(1)), 
                   TXCoordinate(row, column_t(column_t::MAX_COLUMNS)));
}

TXRange TXRange::entireCol(const column_t& col) {
    return TXRange(TXCoordinate(row_t(1), col), 
                   TXCoordinate(row_t(row_t::MAX_ROWS), col));
}

// ==================== TXRange 私有方法实现 ====================

void TXRange::normalize() {
    if (start_.getRow().index() > end_.getRow().index()) {
        std::swap(start_, end_);
    }
    if (start_.getCol().index() > end_.getCol().index()) {
        u32 start_row = start_.getRow().index();
        u32 start_col = start_.getCol().index();
        u32 end_row = end_.getRow().index();
        u32 end_col = end_.getCol().index();
        
        start_ = TXCoordinate(row_t(start_row), column_t(end_col));
        end_ = TXCoordinate(row_t(end_row), column_t(start_col));
    }
}

} // namespace TinaXlsx 