#pragma once

#include <cstdint>
#include <string>
#include <limits>

namespace TinaXlsx {

/**
 * @brief 统一类型定义
 * 
 * 提供TinaXlsx库中使用的所有基础类型定义和常量
 */
namespace TXTypes {

    // ==================== 基础索引类型 ====================
    
    using RowIndex = std::uint32_t;      ///< 行索引类型（从1开始）
    using ColIndex = std::uint32_t;      ///< 列索引类型（从1开始）
    using SheetIndex = std::uint32_t;    ///< 工作表索引类型（从0开始）
    
    // ==================== 样式相关类型 ====================
    
    using ColorValue = std::uint32_t;    ///< 颜色值类型（ARGB格式）
    using FontSize = std::uint32_t;      ///< 字体大小类型
    using BorderWidth = std::uint32_t;   ///< 边框宽度类型
    
    // ==================== 数值类型 ====================
    
    using CellDouble = double;           ///< 单元格浮点数类型
    using CellInteger = std::int64_t;    ///< 单元格整数类型
    
    // ==================== 常量定义 ====================
    
    /// Excel 限制常量
    constexpr RowIndex MAX_ROWS = 1048576;    ///< Excel最大行数
    constexpr ColIndex MAX_COLS = 16384;      ///< Excel最大列数
    constexpr std::size_t MAX_SHEET_NAME = 31; ///< 工作表名称最大长度
    
    /// 无效值常量
    constexpr RowIndex INVALID_ROW = 0;       ///< 无效行号
    constexpr ColIndex INVALID_COL = 0;       ///< 无效列号
    constexpr SheetIndex INVALID_SHEET = std::numeric_limits<SheetIndex>::max();
    
    /// 默认值常量
    constexpr FontSize DEFAULT_FONT_SIZE = 11;    ///< 默认字体大小
    constexpr ColorValue DEFAULT_COLOR = 0xFF000000; ///< 默认颜色（黑色）
    constexpr BorderWidth DEFAULT_BORDER_WIDTH = 1;  ///< 默认边框宽度
    
    // ==================== 文件格式常量 ====================
    
    /// 支持的文件扩展名
    constexpr const char* XLSX_EXTENSION = ".xlsx";
    constexpr const char* XLS_EXTENSION = ".xls";
    constexpr const char* CSV_EXTENSION = ".csv";
    
    /// MIME类型
    constexpr const char* XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    constexpr const char* XLS_MIME_TYPE = "application/vnd.ms-excel";
    constexpr const char* CSV_MIME_TYPE = "text/csv";
    
    // ==================== 验证函数 ====================
    
    /**
     * @brief 检查行号是否有效
     * @param row 行号
     * @return 有效返回true，否则返回false
     */
    constexpr bool isValidRow(RowIndex row) {
        return row >= 1 && row <= MAX_ROWS;
    }
    
    /**
     * @brief 检查列号是否有效
     * @param col 列号
     * @return 有效返回true，否则返回false
     */
    constexpr bool isValidCol(ColIndex col) {
        return col >= 1 && col <= MAX_COLS;
    }
    
    /**
     * @brief 检查坐标是否有效
     * @param row 行号
     * @param col 列号
     * @return 有效返回true，否则返回false
     */
    constexpr bool isValidCoordinate(RowIndex row, ColIndex col) {
        return isValidRow(row) && isValidCol(col);
    }
    
    /**
     * @brief 检查字体大小是否有效
     * @param size 字体大小
     * @return 有效返回true，否则返回false
     */
    constexpr bool isValidFontSize(FontSize size) {
        return size >= 1 && size <= 72; // 通常字体大小范围
    }
    
    /**
     * @brief 检查工作表名称是否有效
     * @param name 工作表名称
     * @return 有效返回true，否则返回false
     */
    bool isValidSheetName(const std::string& name);
    
    // ==================== 工具函数 ====================
    
    /**
     * @brief 计算曼哈顿距离
     * @param row1 起始行
     * @param col1 起始列
     * @param row2 结束行
     * @param col2 结束列
     * @return 曼哈顿距离
     */
    constexpr std::size_t manhattanDistance(RowIndex row1, ColIndex col1, 
                                           RowIndex row2, ColIndex col2) {
        return static_cast<std::size_t>(
            (row1 > row2 ? row1 - row2 : row2 - row1) + 
            (col1 > col2 ? col1 - col2 : col2 - col1)
        );
    }
    
    /**
     * @brief 获取文件扩展名
     * @param filename 文件名
     * @return 扩展名（含点号）
     */
    std::string getFileExtension(const std::string& filename);
    
    /**
     * @brief 检查是否为Excel文件
     * @param filename 文件名
     * @return 是Excel文件返回true，否则返回false
     */
    bool isExcelFile(const std::string& filename);

} // namespace TXTypes

} // namespace TinaXlsx 