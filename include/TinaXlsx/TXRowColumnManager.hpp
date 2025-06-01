#pragma once

#include <unordered_map>
#include <functional>
#include "TXTypes.hpp"
#include "TXCoordinate.hpp"

namespace TinaXlsx {

// 前向声明
class TXCellManager;

/**
 * @brief 行列管理器
 * 
 * 专门负责行列的操作和尺寸管理
 * 职责：
 * - 行列的插入和删除
 * - 行高和列宽的设置
 * - 自动调整行高列宽
 * - 行列的隐藏和显示
 */
class TXRowColumnManager {
public:
    /**
     * @brief 行列操作类型
     */
    enum class OperationType {
        InsertRows,
        DeleteRows,
        InsertColumns,
        DeleteColumns
    };

    TXRowColumnManager() = default;
    ~TXRowColumnManager() = default;

    // 禁用拷贝，支持移动
    TXRowColumnManager(const TXRowColumnManager&) = delete;
    TXRowColumnManager& operator=(const TXRowColumnManager&) = delete;
    TXRowColumnManager(TXRowColumnManager&&) = default;
    TXRowColumnManager& operator=(TXRowColumnManager&&) = default;

    // ==================== 行操作 ====================

    /**
     * @brief 插入行
     * @param row 插入位置
     * @param count 插入数量
     * @param cellManager 单元格管理器（用于移动单元格）
     * @return 成功返回true
     */
    bool insertRows(row_t row, row_t count, TXCellManager& cellManager);

    /**
     * @brief 删除行
     * @param row 删除位置
     * @param count 删除数量
     * @param cellManager 单元格管理器（用于移动单元格）
     * @return 成功返回true
     */
    bool deleteRows(row_t row, row_t count, TXCellManager& cellManager);

    /**
     * @brief 设置行高
     * @param row 行号
     * @param height 高度（磅数）
     * @return 成功返回true
     */
    bool setRowHeight(row_t row, double height);

    /**
     * @brief 获取行高
     * @param row 行号
     * @return 行高（磅数），默认15.0
     */
    double getRowHeight(row_t row) const;

    /**
     * @brief 隐藏行
     * @param row 行号
     * @param hidden 是否隐藏
     * @return 成功返回true
     */
    bool setRowHidden(row_t row, bool hidden);

    /**
     * @brief 检查行是否隐藏
     * @param row 行号
     * @return 隐藏返回true
     */
    bool isRowHidden(row_t row) const;

    // ==================== 列操作 ====================

    /**
     * @brief 插入列
     * @param col 插入位置
     * @param count 插入数量
     * @param cellManager 单元格管理器（用于移动单元格）
     * @return 成功返回true
     */
    bool insertColumns(column_t col, column_t count, TXCellManager& cellManager);

    /**
     * @brief 删除列
     * @param col 删除位置
     * @param count 删除数量
     * @param cellManager 单元格管理器（用于移动单元格）
     * @return 成功返回true
     */
    bool deleteColumns(column_t col, column_t count, TXCellManager& cellManager);

    /**
     * @brief 设置列宽
     * @param col 列号
     * @param width 宽度（字符单位）
     * @return 成功返回true
     */
    bool setColumnWidth(column_t col, double width);

    /**
     * @brief 获取列宽
     * @param col 列号
     * @return 列宽（字符单位），默认8.43
     */
    double getColumnWidth(column_t col) const;

    /**
     * @brief 隐藏列
     * @param col 列号
     * @param hidden 是否隐藏
     * @return 成功返回true
     */
    bool setColumnHidden(column_t col, bool hidden);

    /**
     * @brief 检查列是否隐藏
     * @param col 列号
     * @return 隐藏返回true
     */
    bool isColumnHidden(column_t col) const;

    // ==================== 自动调整 ====================

    /**
     * @brief 自动调整列宽
     * @param col 列号
     * @param cellManager 单元格管理器（用于计算内容宽度）
     * @param minWidth 最小宽度
     * @param maxWidth 最大宽度
     * @return 调整后的列宽
     */
    double autoFitColumnWidth(column_t col, const TXCellManager& cellManager, 
                             double minWidth = 1.0, double maxWidth = 255.0);

    /**
     * @brief 自动调整行高
     * @param row 行号
     * @param cellManager 单元格管理器（用于计算内容高度）
     * @param minHeight 最小高度
     * @param maxHeight 最大高度
     * @return 调整后的行高
     */
    double autoFitRowHeight(row_t row, const TXCellManager& cellManager,
                           double minHeight = 12.0, double maxHeight = 409.0);

    /**
     * @brief 自动调整所有列宽
     * @param cellManager 单元格管理器
     * @param minWidth 最小宽度
     * @param maxWidth 最大宽度
     * @return 调整的列数
     */
    std::size_t autoFitAllColumnWidths(const TXCellManager& cellManager,
                                      double minWidth = 1.0, double maxWidth = 255.0);

    /**
     * @brief 自动调整所有行高
     * @param cellManager 单元格管理器
     * @param minHeight 最小高度
     * @param maxHeight 最大高度
     * @return 调整的行数
     */
    std::size_t autoFitAllRowHeights(const TXCellManager& cellManager,
                                    double minHeight = 12.0, double maxHeight = 409.0);

    // ==================== 批量操作 ====================

    /**
     * @brief 批量设置行高
     * @param heights 行号-高度对列表
     * @return 成功设置的数量
     */
    std::size_t setRowHeights(const std::vector<std::pair<row_t, double>>& heights);

    /**
     * @brief 批量设置列宽
     * @param widths 列号-宽度对列表
     * @return 成功设置的数量
     */
    std::size_t setColumnWidths(const std::vector<std::pair<column_t, double>>& widths);

    // ==================== 查询方法 ====================

    /**
     * @brief 获取所有自定义行高
     * @return 行号-高度映射
     */
    const std::unordered_map<row_t::index_t, double>& getCustomRowHeights() const { return rowHeights_; }

    /**
     * @brief 获取所有自定义列宽
     * @return 列号-宽度映射
     */
    const std::unordered_map<column_t::index_t, double>& getCustomColumnWidths() const { return columnWidths_; }

    /**
     * @brief 清空所有自定义尺寸
     */
    void clear();

private:
    // 存储自定义的行高和列宽
    std::unordered_map<row_t::index_t, double> rowHeights_;
    std::unordered_map<column_t::index_t, double> columnWidths_;
    
    // 存储隐藏状态
    std::unordered_map<row_t::index_t, bool> hiddenRows_;
    std::unordered_map<column_t::index_t, bool> hiddenColumns_;

    // 默认值
    static constexpr double DEFAULT_ROW_HEIGHT = 15.0;
    static constexpr double DEFAULT_COLUMN_WIDTH = 8.43;

    /**
     * @brief 计算文本显示宽度
     * @param text 文本内容
     * @param fontSize 字体大小
     * @param fontName 字体名称
     * @return 显示宽度（字符单位）
     */
    double calculateTextWidth(const std::string& text, double fontSize = 11.0, 
                             const std::string& fontName = "Calibri") const;

    /**
     * @brief 计算文本显示高度
     * @param text 文本内容
     * @param fontSize 字体大小
     * @param columnWidth 列宽（用于换行计算）
     * @return 显示高度（磅数）
     */
    double calculateTextHeight(const std::string& text, double fontSize = 11.0, 
                              double columnWidth = 8.43) const;

    /**
     * @brief 验证行号有效性
     * @param row 行号
     * @return 有效返回true
     */
    bool isValidRow(row_t row) const;

    /**
     * @brief 验证列号有效性
     * @param col 列号
     * @return 有效返回true
     */
    bool isValidColumn(column_t col) const;

    /**
     * @brief 验证尺寸有效性
     * @param size 尺寸值
     * @param isWidth 是否为宽度（否则为高度）
     * @return 有效返回true
     */
    bool isValidSize(double size, bool isWidth) const;
};

} // namespace TinaXlsx
