#include "TinaXlsx/TXDataFilter.hpp"
#include <algorithm>
#include <cctype>

namespace TinaXlsx {

// ==================== TXAutoFilter 实现 ====================

TXAutoFilter::TXAutoFilter(const TXRange& range)
    : range_(range)
    , showFilterButtons_(true)
{
}

void TXAutoFilter::addFilterCondition(const FilterCondition& condition) {
    // 移除同一列的现有条件
    removeFilterCondition(condition.columnIndex);
    
    // 添加新条件
    filterConditions_.push_back(condition);
}

void TXAutoFilter::removeFilterCondition(u32 columnIndex) {
    filterConditions_.erase(
        std::remove_if(filterConditions_.begin(), filterConditions_.end(),
                      [columnIndex](const FilterCondition& condition) {
                          return condition.columnIndex == columnIndex;
                      }),
        filterConditions_.end()
    );
}

void TXAutoFilter::clearFilterConditions() {
    filterConditions_.clear();
    customFilters_.clear();
}

void TXAutoFilter::setTextFilter(u32 columnIndex, const std::string& text, 
                                FilterOperator operator_, bool caseSensitive) {
    FilterCondition condition(columnIndex, operator_, text, "", caseSensitive);
    addFilterCondition(condition);
}

void TXAutoFilter::setNumberFilter(u32 columnIndex, double value, FilterOperator operator_) {
    FilterCondition condition(columnIndex, operator_, std::to_string(value));
    addFilterCondition(condition);
}

void TXAutoFilter::setRangeFilter(u32 columnIndex, double minValue, double maxValue) {
    FilterCondition condition(columnIndex, FilterOperator::GreaterThanOrEqual, 
                            std::to_string(minValue), std::to_string(maxValue));
    addFilterCondition(condition);
    
    // 添加第二个条件（小于等于最大值）
    FilterCondition condition2(columnIndex, FilterOperator::LessThanOrEqual, 
                              std::to_string(maxValue));
    filterConditions_.push_back(condition2);
}

void TXAutoFilter::setTopNFilter(u32 columnIndex, int count, bool isTop) {
    FilterOperator op = isTop ? FilterOperator::Top10 : FilterOperator::Bottom10;
    FilterCondition condition(columnIndex, op, std::to_string(count));
    addFilterCondition(condition);
}

void TXAutoFilter::setCustomFilter(u32 columnIndex, std::function<bool(const std::string&)> customFunction) {
    // 确保自定义筛选器向量足够大
    if (customFilters_.size() <= columnIndex) {
        customFilters_.resize(columnIndex + 1);
    }
    customFilters_[columnIndex] = customFunction;
}

// ==================== TXDataSorter 实现 ====================

TXDataSorter::TXDataSorter(const TXRange& range)
    : range_(range)
    , hasHeader_(true)
{
}

void TXDataSorter::addSortCondition(const SortCondition& condition) {
    sortConditions_.push_back(condition);
}

void TXDataSorter::clearSortConditions() {
    sortConditions_.clear();
}

void TXDataSorter::sortByColumn(u32 columnIndex, SortOrder order, bool caseSensitive) {
    clearSortConditions();
    addSortCondition(SortCondition(columnIndex, order, caseSensitive));
}

void TXDataSorter::sortByMultipleColumns(const std::vector<SortCondition>& conditions) {
    sortConditions_ = conditions;
}

void TXDataSorter::customSort(std::function<bool(const std::vector<std::string>&, const std::vector<std::string>&)> compareFunction) {
    // 自定义排序的实现将在实际应用时处理
    // 这里保存比较函数供后续使用
    (void)compareFunction; // 避免未使用参数警告
}

// ==================== TXDataTable 实现 ====================

TXDataTable::TXDataTable(const TXRange& range, bool hasHeader)
    : range_(range)
    , hasHeader_(hasHeader)
    , sorter_(range)
    , alternateRowColors_(false)
{
    sorter_.setHasHeader(hasHeader);
}

TXAutoFilter& TXDataTable::enableAutoFilter() {
    if (!autoFilter_) {
        autoFilter_ = std::make_unique<TXAutoFilter>(range_);
    }
    return *autoFilter_;
}

void TXDataTable::disableAutoFilter() {
    autoFilter_.reset();
}

void TXDataTable::setTableStyle(bool showFilterButtons, bool alternateRowColors) {
    alternateRowColors_ = alternateRowColors;
    
    if (showFilterButtons) {
        enableAutoFilter();
        autoFilter_->setShowFilterButtons(true);
    } else {
        disableAutoFilter();
    }
}

void TXDataTable::applyTableFormat(const std::string& headerStyle, 
                                  const std::string& dataStyle, 
                                  const std::string& alternateStyle) {
    // 表格格式应用的实现将在实际的工作表操作中处理
    // 这里保存样式信息供后续使用
    (void)headerStyle;
    (void)dataStyle;
    (void)alternateStyle;
}

} // namespace TinaXlsx
