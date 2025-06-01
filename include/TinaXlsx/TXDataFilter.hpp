#pragma once

#include "TXTypes.hpp"
#include "TXRange.hpp"
#include <string>
#include <vector>
#include <functional>
#include <memory>

namespace TinaXlsx {

/**
 * @brief 排序方向
 */
enum class SortOrder {
    Ascending,      ///< 升序
    Descending      ///< 降序
};

/**
 * @brief 筛选操作符
 */
enum class FilterOperator {
    Equal,          ///< 等于
    NotEqual,       ///< 不等于
    GreaterThan,    ///< 大于
    LessThan,       ///< 小于
    GreaterThanOrEqual, ///< 大于等于
    LessThanOrEqual,    ///< 小于等于
    Contains,       ///< 包含
    NotContains,    ///< 不包含
    BeginsWith,     ///< 开始于
    EndsWith,       ///< 结束于
    IsEmpty,        ///< 为空
    IsNotEmpty,     ///< 不为空
    Top10,          ///< 前10项
    Bottom10        ///< 后10项
};

/**
 * @brief 排序条件
 */
struct SortCondition {
    u32 columnIndex;        ///< 列索引（从0开始）
    SortOrder order;        ///< 排序方向
    bool caseSensitive;     ///< 是否区分大小写
    
    SortCondition(u32 col, SortOrder ord, bool caseSens = false)
        : columnIndex(col), order(ord), caseSensitive(caseSens) {}
};

/**
 * @brief 筛选条件
 */
struct FilterCondition {
    u32 columnIndex;            ///< 列索引（从0开始）
    FilterOperator operator_;   ///< 筛选操作符
    std::string value1;         ///< 第一个值
    std::string value2;         ///< 第二个值（用于Between等操作）
    bool caseSensitive;         ///< 是否区分大小写
    
    FilterCondition(u32 col, FilterOperator op, const std::string& val1 = "", 
                   const std::string& val2 = "", bool caseSens = false)
        : columnIndex(col), operator_(op), value1(val1), value2(val2), caseSensitive(caseSens) {}
};

/**
 * @brief 自动筛选设置
 */
class TXAutoFilter {
public:
    /**
     * @brief 构造函数
     * @param range 筛选范围
     */
    explicit TXAutoFilter(const TXRange& range);

    // ==================== 基本属性 ====================

    /**
     * @brief 获取筛选范围
     */
    const TXRange& getRange() const { return range_; }

    /**
     * @brief 设置筛选范围
     * @param range 筛选范围
     */
    void setRange(const TXRange& range) { range_ = range; }

    /**
     * @brief 设置是否显示筛选按钮
     * @param show 是否显示
     */
    void setShowFilterButtons(bool show) { showFilterButtons_ = show; }

    /**
     * @brief 获取是否显示筛选按钮
     */
    bool getShowFilterButtons() const { return showFilterButtons_; }

    // ==================== 筛选条件管理 ====================

    /**
     * @brief 添加筛选条件
     * @param condition 筛选条件
     */
    void addFilterCondition(const FilterCondition& condition);

    /**
     * @brief 移除指定列的筛选条件
     * @param columnIndex 列索引
     */
    void removeFilterCondition(u32 columnIndex);

    /**
     * @brief 清除所有筛选条件
     */
    void clearFilterConditions();

    /**
     * @brief 获取筛选条件列表
     */
    const std::vector<FilterCondition>& getFilterConditions() const { return filterConditions_; }

    // ==================== 便捷筛选方法 ====================

    /**
     * @brief 设置文本筛选
     * @param columnIndex 列索引
     * @param text 筛选文本
     * @param operator_ 操作符
     * @param caseSensitive 是否区分大小写
     */
    void setTextFilter(u32 columnIndex, const std::string& text, 
                      FilterOperator operator_ = FilterOperator::Contains, 
                      bool caseSensitive = false);

    /**
     * @brief 设置数值筛选
     * @param columnIndex 列索引
     * @param value 筛选值
     * @param operator_ 操作符
     */
    void setNumberFilter(u32 columnIndex, double value, FilterOperator operator_);

    /**
     * @brief 设置范围筛选
     * @param columnIndex 列索引
     * @param minValue 最小值
     * @param maxValue 最大值
     */
    void setRangeFilter(u32 columnIndex, double minValue, double maxValue);

    /**
     * @brief 设置前N项筛选
     * @param columnIndex 列索引
     * @param count 项数
     * @param isTop 是否为前N项（false为后N项）
     */
    void setTopNFilter(u32 columnIndex, int count, bool isTop = true);

    /**
     * @brief 设置自定义筛选
     * @param columnIndex 列索引
     * @param customFunction 自定义筛选函数
     */
    void setCustomFilter(u32 columnIndex, std::function<bool(const std::string&)> customFunction);

private:
    TXRange range_;                                 ///< 筛选范围
    bool showFilterButtons_;                        ///< 是否显示筛选按钮
    std::vector<FilterCondition> filterConditions_; ///< 筛选条件列表
    std::vector<std::function<bool(const std::string&)>> customFilters_; ///< 自定义筛选函数
};

/**
 * @brief 数据排序器
 */
class TXDataSorter {
public:
    /**
     * @brief 构造函数
     * @param range 排序范围
     */
    explicit TXDataSorter(const TXRange& range);

    // ==================== 基本属性 ====================

    /**
     * @brief 获取排序范围
     */
    const TXRange& getRange() const { return range_; }

    /**
     * @brief 设置排序范围
     * @param range 排序范围
     */
    void setRange(const TXRange& range) { range_ = range; }

    /**
     * @brief 设置是否包含标题行
     * @param hasHeader 是否包含标题行
     */
    void setHasHeader(bool hasHeader) { hasHeader_ = hasHeader; }

    /**
     * @brief 获取是否包含标题行
     */
    bool getHasHeader() const { return hasHeader_; }

    // ==================== 排序条件管理 ====================

    /**
     * @brief 添加排序条件
     * @param condition 排序条件
     */
    void addSortCondition(const SortCondition& condition);

    /**
     * @brief 清除所有排序条件
     */
    void clearSortConditions();

    /**
     * @brief 获取排序条件列表
     */
    const std::vector<SortCondition>& getSortConditions() const { return sortConditions_; }

    // ==================== 便捷排序方法 ====================

    /**
     * @brief 按单列排序
     * @param columnIndex 列索引
     * @param order 排序方向
     * @param caseSensitive 是否区分大小写
     */
    void sortByColumn(u32 columnIndex, SortOrder order, bool caseSensitive = false);

    /**
     * @brief 按多列排序
     * @param conditions 排序条件列表
     */
    void sortByMultipleColumns(const std::vector<SortCondition>& conditions);

    /**
     * @brief 自定义排序
     * @param compareFunction 自定义比较函数
     */
    void customSort(std::function<bool(const std::vector<std::string>&, const std::vector<std::string>&)> compareFunction);

private:
    TXRange range_;                             ///< 排序范围
    bool hasHeader_;                            ///< 是否包含标题行
    std::vector<SortCondition> sortConditions_; ///< 排序条件列表
};

/**
 * @brief 数据表格管理器
 * 
 * 集成排序和筛选功能的数据表格管理器
 */
class TXDataTable {
public:
    /**
     * @brief 构造函数
     * @param range 数据表格范围
     * @param hasHeader 是否包含标题行
     */
    TXDataTable(const TXRange& range, bool hasHeader = true);

    // ==================== 基本属性 ====================

    /**
     * @brief 获取数据范围
     */
    const TXRange& getRange() const { return range_; }

    /**
     * @brief 设置数据范围
     * @param range 数据范围
     */
    void setRange(const TXRange& range) { range_ = range; }

    /**
     * @brief 获取是否包含标题行
     */
    bool getHasHeader() const { return hasHeader_; }

    /**
     * @brief 设置是否包含标题行
     * @param hasHeader 是否包含标题行
     */
    void setHasHeader(bool hasHeader) { hasHeader_ = hasHeader; }

    // ==================== 筛选功能 ====================

    /**
     * @brief 启用自动筛选
     * @return 自动筛选对象引用
     */
    TXAutoFilter& enableAutoFilter();

    /**
     * @brief 禁用自动筛选
     */
    void disableAutoFilter();

    /**
     * @brief 获取自动筛选对象
     */
    TXAutoFilter* getAutoFilter() { return autoFilter_.get(); }

    /**
     * @brief 检查是否启用了自动筛选
     */
    bool hasAutoFilter() const { return autoFilter_ != nullptr; }

    // ==================== 排序功能 ====================

    /**
     * @brief 获取数据排序器
     * @return 数据排序器引用
     */
    TXDataSorter& getSorter() { return sorter_; }

    /**
     * @brief 获取数据排序器（只读）
     */
    const TXDataSorter& getSorter() const { return sorter_; }

    // ==================== 便捷方法 ====================

    /**
     * @brief 快速设置表格样式
     * @param showFilterButtons 是否显示筛选按钮
     * @param alternateRowColors 是否使用交替行颜色
     */
    void setTableStyle(bool showFilterButtons = true, bool alternateRowColors = false);

    /**
     * @brief 应用表格格式
     * @param headerStyle 标题行样式
     * @param dataStyle 数据行样式
     * @param alternateStyle 交替行样式（可选）
     */
    void applyTableFormat(const std::string& headerStyle, 
                         const std::string& dataStyle, 
                         const std::string& alternateStyle = "");

private:
    TXRange range_;                             ///< 数据表格范围
    bool hasHeader_;                            ///< 是否包含标题行
    std::unique_ptr<TXAutoFilter> autoFilter_;  ///< 自动筛选器
    TXDataSorter sorter_;                       ///< 数据排序器
    bool alternateRowColors_;                   ///< 是否使用交替行颜色
};

} // namespace TinaXlsx
