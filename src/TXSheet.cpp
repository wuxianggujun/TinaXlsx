#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include <unordered_map>
#include <regex>
#include <algorithm>

namespace TinaXlsx {

// 哈希函数特化
struct CellCoordinateHash {
    std::size_t operator()(const TXSheet::CellCoordinate& coord) const {
        return std::hash<std::size_t>()(coord.row) ^ (std::hash<std::size_t>()(coord.col) << 1);
    }
};

class TXSheet::Impl {
public:
    explicit Impl(const std::string& name) : name_(name), last_error_("") {}

    const std::string& getName() const {
        return name_;
    }

    void setName(const std::string& name) {
        name_ = name;
    }

    TXSheet::CellValue getCellValue(const CellCoordinate& coord) const {
        auto it = cells_.find(coord);
        if (it != cells_.end()) {
            return it->second.getValue();
        }
        return std::string("");  // 默认返回空字符串
    }

    bool setCellValue(const CellCoordinate& coord, const CellValue& value) {
        if (coord.row == 0 || coord.col == 0) {
            last_error_ = "Invalid cell coordinate (row and column must be >= 1)";
            return false;
        }

        cells_[coord].setValue(value);
        last_error_.clear();
        return true;
    }

    TXCell* getCell(const CellCoordinate& coord) {
        if (coord.row == 0 || coord.col == 0) {
            return nullptr;
        }

        auto it = cells_.find(coord);
        if (it != cells_.end()) {
            return &it->second;
        }

        // 创建新的单元格
        cells_[coord] = TXCell();
        return &cells_[coord];
    }

    const TXCell* getCell(const CellCoordinate& coord) const {
        auto it = cells_.find(coord);
        if (it != cells_.end()) {
            return &it->second;
        }
        return nullptr;
    }

    bool insertRows(std::size_t row, std::size_t count) {
        if (row == 0) {
            last_error_ = "Invalid row number (must be >= 1)";
            return false;
        }

        // 需要移动的单元格
        std::unordered_map<CellCoordinate, TXCell, CellCoordinateHash> new_cells;
        
        for (auto& pair : cells_) {
            const auto& coord = pair.first;
            if (coord.row >= row) {
                // 向下移动
                CellCoordinate new_coord{coord.row + count, coord.col};
                new_cells[new_coord] = std::move(pair.second);
            } else {
                new_cells[coord] = std::move(pair.second);
            }
        }

        cells_ = std::move(new_cells);
        last_error_.clear();
        return true;
    }

    bool deleteRows(std::size_t row, std::size_t count) {
        if (row == 0) {
            last_error_ = "Invalid row number (must be >= 1)";
            return false;
        }

        std::unordered_map<CellCoordinate, TXCell, CellCoordinateHash> new_cells;
        
        for (auto& pair : cells_) {
            const auto& coord = pair.first;
            if (coord.row < row) {
                // 保持不变
                new_cells[coord] = std::move(pair.second);
            } else if (coord.row >= row + count) {
                // 向上移动
                CellCoordinate new_coord{coord.row - count, coord.col};
                new_cells[new_coord] = std::move(pair.second);
            }
            // 在删除范围内的单元格被丢弃
        }

        cells_ = std::move(new_cells);
        last_error_.clear();
        return true;
    }

    bool insertColumns(std::size_t col, std::size_t count) {
        if (col == 0) {
            last_error_ = "Invalid column number (must be >= 1)";
            return false;
        }

        std::unordered_map<CellCoordinate, TXCell, CellCoordinateHash> new_cells;
        
        for (auto& pair : cells_) {
            const auto& coord = pair.first;
            if (coord.col >= col) {
                // 向右移动
                CellCoordinate new_coord{coord.row, coord.col + count};
                new_cells[new_coord] = std::move(pair.second);
            } else {
                new_cells[coord] = std::move(pair.second);
            }
        }

        cells_ = std::move(new_cells);
        last_error_.clear();
        return true;
    }

    bool deleteColumns(std::size_t col, std::size_t count) {
        if (col == 0) {
            last_error_ = "Invalid column number (must be >= 1)";
            return false;
        }

        std::unordered_map<CellCoordinate, TXCell, CellCoordinateHash> new_cells;
        
        for (auto& pair : cells_) {
            const auto& coord = pair.first;
            if (coord.col < col) {
                // 保持不变
                new_cells[coord] = std::move(pair.second);
            } else if (coord.col >= col + count) {
                // 向左移动
                CellCoordinate new_coord{coord.row, coord.col - count};
                new_cells[new_coord] = std::move(pair.second);
            }
            // 在删除范围内的单元格被丢弃
        }

        cells_ = std::move(new_cells);
        last_error_.clear();
        return true;
    }

    std::size_t getUsedRowCount() const {
        std::size_t max_row = 0;
        for (const auto& pair : cells_) {
            if (!pair.second.isEmpty()) {
                max_row = std::max(max_row, pair.first.row);
            }
        }
        return max_row;
    }

    std::size_t getUsedColumnCount() const {
        std::size_t max_col = 0;
        for (const auto& pair : cells_) {
            if (!pair.second.isEmpty()) {
                max_col = std::max(max_col, pair.first.col);
            }
        }
        return max_col;
    }

    TXSheet::CellRange getUsedRange() const {
        if (cells_.empty()) {
            return CellRange{CellCoordinate{1, 1}, CellCoordinate{1, 1}};
        }

        std::size_t min_row = SIZE_MAX, max_row = 0;
        std::size_t min_col = SIZE_MAX, max_col = 0;

        for (const auto& pair : cells_) {
            if (!pair.second.isEmpty()) {
                min_row = std::min(min_row, pair.first.row);
                max_row = std::max(max_row, pair.first.row);
                min_col = std::min(min_col, pair.first.col);
                max_col = std::max(max_col, pair.first.col);
            }
        }

        if (min_row == SIZE_MAX) {
            return CellRange{CellCoordinate{1, 1}, CellCoordinate{1, 1}};
        }

        return CellRange{CellCoordinate{min_row, min_col}, CellCoordinate{max_row, max_col}};
    }

    void clear() {
        cells_.clear();
        last_error_.clear();
    }

    std::size_t setCellValues(const std::vector<std::pair<CellCoordinate, CellValue>>& values) {
        std::size_t success_count = 0;
        for (const auto& pair : values) {
            if (setCellValue(pair.first, pair.second)) {
                ++success_count;
            }
        }
        return success_count;
    }

    std::vector<std::pair<TXSheet::CellCoordinate, TXSheet::CellValue>> 
    getCellValues(const std::vector<CellCoordinate>& coords) const {
        std::vector<std::pair<CellCoordinate, CellValue>> results;
        results.reserve(coords.size());
        
        for (const auto& coord : coords) {
            results.emplace_back(coord, getCellValue(coord));
        }
        
        return results;
    }

    bool setRangeValues(const TXSheet::CellRange& range, const std::vector<std::vector<CellValue>>& values) {
        if (values.empty()) {
            last_error_ = "Values array is empty";
            return false;
        }

        std::size_t row_count = range.end.row - range.start.row + 1;
        std::size_t col_count = range.end.col - range.start.col + 1;

        if (values.size() != row_count) {
            last_error_ = "Values row count doesn't match range";
            return false;
        }

        for (std::size_t i = 0; i < values.size(); ++i) {
            if (values[i].size() != col_count) {
                last_error_ = "Values column count doesn't match range";
                return false;
            }
        }

        // 设置值
        for (std::size_t row = 0; row < row_count; ++row) {
            for (std::size_t col = 0; col < col_count; ++col) {
                CellCoordinate coord{range.start.row + row, range.start.col + col};
                setCellValue(coord, values[row][col]);
            }
        }

        last_error_.clear();
        return true;
    }

    std::vector<std::vector<TXSheet::CellValue>> getRangeValues(const TXSheet::CellRange& range) const {
        std::size_t row_count = range.end.row - range.start.row + 1;
        std::size_t col_count = range.end.col - range.start.col + 1;

        std::vector<std::vector<CellValue>> result(row_count);
        for (auto& row : result) {
            row.resize(col_count);
        }

        for (std::size_t row = 0; row < row_count; ++row) {
            for (std::size_t col = 0; col < col_count; ++col) {
                CellCoordinate coord{range.start.row + row, range.start.col + col};
                result[row][col] = getCellValue(coord);
            }
        }

        return result;
    }

    const std::string& getLastError() const {
        return last_error_;
    }

private:
    std::string name_;
    std::unordered_map<CellCoordinate, TXCell, CellCoordinateHash> cells_;
    mutable std::string last_error_;
};

// TXSheet 实现
TXSheet::TXSheet(const std::string& name) : pImpl(std::make_unique<Impl>(name)) {}

TXSheet::~TXSheet() = default;

TXSheet::TXSheet(TXSheet&& other) noexcept : pImpl(std::move(other.pImpl)) {}

TXSheet& TXSheet::operator=(TXSheet&& other) noexcept {
    if (this != &other) {
        pImpl = std::move(other.pImpl);
    }
    return *this;
}

const std::string& TXSheet::getName() const {
    return pImpl->getName();
}

void TXSheet::setName(const std::string& name) {
    pImpl->setName(name);
}

TXSheet::CellValue TXSheet::getCellValue(std::size_t row, std::size_t col) const {
    return pImpl->getCellValue(CellCoordinate{row, col});
}

TXSheet::CellValue TXSheet::getCellValue(const CellCoordinate& coord) const {
    return pImpl->getCellValue(coord);
}

TXSheet::CellValue TXSheet::getCellValue(const std::string& address) const {
    return getCellValue(addressToCoordinate(address));
}

bool TXSheet::setCellValue(std::size_t row, std::size_t col, const CellValue& value) {
    return pImpl->setCellValue(CellCoordinate{row, col}, value);
}

bool TXSheet::setCellValue(const CellCoordinate& coord, const CellValue& value) {
    return pImpl->setCellValue(coord, value);
}

bool TXSheet::setCellValue(const std::string& address, const CellValue& value) {
    return setCellValue(addressToCoordinate(address), value);
}

TXCell* TXSheet::getCell(std::size_t row, std::size_t col) {
    return pImpl->getCell(CellCoordinate{row, col});
}

const TXCell* TXSheet::getCell(std::size_t row, std::size_t col) const {
    return pImpl->getCell(CellCoordinate{row, col});
}

bool TXSheet::insertRows(std::size_t row, std::size_t count) {
    return pImpl->insertRows(row, count);
}

bool TXSheet::deleteRows(std::size_t row, std::size_t count) {
    return pImpl->deleteRows(row, count);
}

bool TXSheet::insertColumns(std::size_t col, std::size_t count) {
    return pImpl->insertColumns(col, count);
}

bool TXSheet::deleteColumns(std::size_t col, std::size_t count) {
    return pImpl->deleteColumns(col, count);
}

std::size_t TXSheet::getUsedRowCount() const {
    return pImpl->getUsedRowCount();
}

std::size_t TXSheet::getUsedColumnCount() const {
    return pImpl->getUsedColumnCount();
}

TXSheet::CellRange TXSheet::getUsedRange() const {
    return pImpl->getUsedRange();
}

void TXSheet::clear() {
    pImpl->clear();
}

std::size_t TXSheet::setCellValues(const std::vector<std::pair<CellCoordinate, CellValue>>& values) {
    return pImpl->setCellValues(values);
}

std::vector<std::pair<TXSheet::CellCoordinate, TXSheet::CellValue>> 
TXSheet::getCellValues(const std::vector<CellCoordinate>& coords) const {
    return pImpl->getCellValues(coords);
}

bool TXSheet::setRangeValues(const CellRange& range, const std::vector<std::vector<CellValue>>& values) {
    return pImpl->setRangeValues(range, values);
}

std::vector<std::vector<TXSheet::CellValue>> TXSheet::getRangeValues(const CellRange& range) const {
    return pImpl->getRangeValues(range);
}

// 静态方法实现
TXSheet::CellCoordinate TXSheet::addressToCoordinate(const std::string& address) {
    std::regex addr_regex("^([A-Z]+)(\\d+)$");
    std::smatch match;
    
    if (!std::regex_match(address, match, addr_regex)) {
        return CellCoordinate{1, 1};  // 默认返回A1
    }

    std::string col_str = match[1].str();
    std::size_t row = std::stoull(match[2].str());

    // 转换列字母为数字 (A=1, B=2, ..., Z=26, AA=27, ...)
    std::size_t col = 0;
    for (char c : col_str) {
        col = col * 26 + (c - 'A' + 1);
    }

    return CellCoordinate{row, col};
}

std::string TXSheet::coordinateToAddress(const CellCoordinate& coord) {
    std::string col_str;
    std::size_t col = coord.col;
    
    while (col > 0) {
        col--;  // 转换为0基
        col_str = char('A' + (col % 26)) + col_str;
        col /= 26;
    }
    
    return col_str + std::to_string(coord.row);
}

const std::string& TXSheet::getLastError() const {
    return pImpl->getLastError();
}

} // namespace TinaXlsx 