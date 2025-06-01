#pragma once

#include <string>
#include <vector>
#include <memory>
#include <variant>
#include <unordered_map>

#include "TXCoordinate.hpp"
#include "TXRange.hpp"
#include "TXCell.hpp"
#include "TXTypes.hpp"
#include "TXMergedCells.hpp"
#include "TXWorkbook.hpp"

// 新增管理器头文件
#include "TXCellManager.hpp"
#include "TXRowColumnManager.hpp"
#include "TXSheetProtectionManager.hpp"
#include "TXFormulaManager.hpp"

namespace TinaXlsx {

// Forward declarations
// class TXCell; // TXCell.hpp 已包含
class TXWorkbook;
class TXCellStyle; // Forward declaration for TXCellStyle if its full definition is not needed here

/**
 * @brief Excel工作表类
 *
 * 表示Excel中的一个Sheet，负责管理Sheet内的数据操作
 */
class TXSheet {
public:
    // ==================== 类型别名 ====================
    using CellValue = cell_value_t;
    using Coordinate = TXCoordinate;
    using Range = TXRange;

public:
    explicit TXSheet(const std::string& name, TXWorkbook* parentWorkbook);
    ~TXSheet();
    
    // 禁用拷贝构造和赋值
    TXSheet(const TXSheet&) = delete;
    TXSheet& operator=(const TXSheet&) = delete;

    // 支持移动构造和赋值
    TXSheet(TXSheet&& other) noexcept;
    TXSheet& operator=(TXSheet&& other) noexcept;

    /**
     * @brief 获取工作表名称
     * @return 工作表名称
     */
    const std::string& getName() const;

    /**
     * @brief 设置工作表名称
     * @param name 新名称
     */
    void setName(const std::string& name);

    // ==================== 单元格值操作 ====================

    /**
     * @brief 获取单元格值
     * @param row 行号 (1-based)
     * @param col 列号 (1-based)
     * @return 单元格值
     */
    CellValue getCellValue(row_t row, column_t col) const;

    /**
     * @brief 获取单元格值
     * @param coord 单元格坐标
     * @return 单元格值
     */
    CellValue getCellValue(const Coordinate& coord) const;

    /**
     * @brief 获取单元格值（使用A1格式）
     * @param address 单元格地址，如"A1", "B2"
     * @return 单元格值
     */
    CellValue getCellValue(const std::string& address) const;

    /**
     * @brief 设置单元格值
     * @param row 行号（1开始）
     * @param col 列号（1开始）
     * @param value 单元格值
     * @return 成功返回true，失败返回false
     */
    bool setCellValue(row_t row, column_t col, const CellValue& value);

    /**
     * @brief 设置单元格值
     * @param coord 单元格坐标
     * @param value 单元格值
     * @return 成功返回true，失败返回false
     */
    bool setCellValue(const Coordinate& coord, const CellValue& value);

    /**
     * @brief 设置单元格值（使用A1格式）
     * @param address 单元格地址，如"A1", "B2"
     * @param value 单元格值
     * @return 成功返回true，失败返回false
     */
    bool setCellValue(const std::string& address, const CellValue& value);

    // ==================== 单元格对象访问 ====================

    /**
     * @brief 获取单元格
     * @param row 行号（1开始）
     * @param col 列号（1开始）
     * @return 单元格指针，如果不存在返回nullptr
     */
    TXCell* getCell(row_t row, column_t col);

    /**
     * @brief 获取单元格（const版本）
     * @param row 行号（1开始）
     * @param col 列号（1开始）
     * @return 单元格指针，如果不存在返回nullptr
     */
    const TXCell* getCell(row_t row, column_t col) const;

    /**
     * @brief 获取单元格
     * @param coord 单元格坐标
     * @return 单元格指针，如果不存在返回nullptr
     */
    TXCell* getCell(const Coordinate& coord);

    /**
     * @brief 获取单元格（const版本）
     * @param coord 单元格坐标
     * @return 单元格指针，如果不存在返回nullptr
     */
    const TXCell* getCell(const Coordinate& coord) const;

    /**
     * @brief 获取单元格（使用A1格式）
     * @param address 单元格地址，如"A1", "B2"
     * @return 单元格指针，如果不存在返回nullptr
     */
    TXCell* getCell(const std::string& address);

    /**
     * @brief 获取单元格（const版本，使用A1格式）
     * @param address 单元格地址，如"A1", "B2"
     * @return 单元格指针，如果不存在返回nullptr
     */
    const TXCell* getCell(const std::string& address) const;

    // ==================== 行列操作 ====================

    /**
     * @brief 插入行
     * @param row 插入位置（1开始）
     * @param count 插入行数
     * @return 成功返回true，失败返回false
     */
    bool insertRows(row_t row, row_t count = row_t(1));

    /**
     * @brief 删除行
     * @param row 删除位置（1开始）
     * @param count 删除行数
     * @return 成功返回true，失败返回false
     */
    bool deleteRows(row_t row, row_t count = row_t(1));

    /**
     * @brief 插入列
     * @param col 插入位置（1开始）
     * @param count 插入列数
     * @return 成功返回true，失败返回false
     */
    bool insertColumns(column_t col, column_t count = column_t(1));

    /**
     * @brief 删除列
     * @param col 删除位置（1开始）
     * @param count 删除列数
     * @return 成功返回true，失败返回false
     */
    bool deleteColumns(column_t col, column_t count = column_t(1));

    // ==================== 范围信息 ====================

    /**
     * @brief 获取使用的行数
     * @return 最大使用行号
     */
    row_t getUsedRowCount() const;

    /**
     * @brief 获取使用的列数
     * @return 最大使用列号
     */
    column_t getUsedColumnCount() const;

    /**
     * @brief 获取使用的范围
     * @return 使用的单元格范围
     */
    Range getUsedRange() const;

    /**
     * @brief 清空工作表
     */
    void clear();

    // ==================== 批量操作 ====================

    /**
     * @brief 批量设置单元格值（高性能版本）
     * @param values 坐标到值的映射
     * @return 成功设置的单元格数量
     */
    std::size_t setCellValues(const std::vector<std::pair<Coordinate, CellValue>>& values);

    /**
     * @brief 批量获取单元格值（高性能版本）
     * @param coords 坐标列表
     * @return 坐标到值的映射
     */
    std::vector<std::pair<Coordinate, CellValue>> getCellValues(const std::vector<Coordinate>& coords) const;

    /**
     * @brief 设置范围内的值
     * @param range 单元格范围
     * @param values 值的二维数组
     * @return 成功返回true，失败返回false
     */
    bool setRangeValues(const Range& range, const std::vector<std::vector<CellValue>>& values);

    /**
     * @brief 获取范围内的值
     * @param range 单元格范围
     * @return 值的二维数组
     */
    std::vector<std::vector<CellValue>> getRangeValues(const Range& range) const;

    // ==================== 合并单元格操作 ====================

    /**
     * @brief 合并单元格区域
     * @param startRow 起始行
     * @param startCol 起始列
     * @param endRow 结束行
     * @param endCol 结束列
     * @return 成功返回true，失败返回false
     */
    bool mergeCells(row_t startRow, column_t startCol,
                   row_t endRow, column_t endCol);

    /**
     * @brief 合并单元格区域
     * @param range 要合并的范围
     * @return 成功返回true，失败返回false
     */
    bool mergeCells(const Range& range);

    /**
     * @brief 合并单元格区域（使用A1格式）
     * @param rangeStr 范围字符串，如"A1:C3"
     * @return 成功返回true，失败返回false
     */
    bool mergeCells(const std::string& rangeStr);

    /**
     * @brief 拆分包含指定单元格的合并区域
     * @param row 单元格行号
     * @param col 单元格列号
     * @return 成功返回true，失败返回false
     */
    bool unmergeCells(row_t row, column_t col);

    /**
     * @brief 拆分指定范围的所有合并区域
     * @param range 范围
     * @return 拆分的区域数量
     */
    std::size_t unmergeCellsInRange(const Range& range);

    /**
     * @brief 检查单元格是否被合并
     * @param row 行号
     * @param col 列号
     * @return 被合并返回true，否则返回false
     */
    bool isCellMerged(row_t row, column_t col) const;

    /**
     * @brief 获取包含指定单元格的合并区域
     * @param row 行号
     * @param col 列号
     * @return 合并区域，如果不存在返回空Range
     */
    Range getMergeRegion(row_t row, column_t col) const;

    /**
     * @brief 获取所有合并区域
     * @return 合并区域列表
     */
    std::vector<Range> getAllMergeRegions() const;

    /**
     * @brief 获取合并区域数量
     * @return 合并区域数量
     */
    std::size_t getMergeCount() const;

    // ==================== 公式计算 ====================

    /**
     * @brief 计算所有公式
     * @return 成功计算的公式数量
     */
    std::size_t calculateAllFormulas();

    /**
     * @brief 计算指定范围内的公式
     * @param range 范围
     * @return 成功计算的公式数量
     */
    std::size_t calculateFormulasInRange(const Range& range);

    /**
     * @brief 设置单元格公式
     * @param row 行号
     * @param col 列号
     * @param formula 公式字符串
     * @return 成功返回true，失败返回false
     */
    bool setCellFormula(row_t row, column_t col, const std::string& formula);

    /**
     * @brief 获取单元格公式
     * @param row 行号
     * @param col 列号
     * @return 公式字符串，如果不是公式返回空字符串
     */
    std::string getCellFormula(row_t row, column_t col) const;

    /**
     * @brief 批量设置公式（高性能版本）
     * @param formulas 坐标到公式的映射
     * @return 成功设置的公式数量
     */
    std::size_t setCellFormulas(const std::vector<std::pair<Coordinate, std::string>>& formulas);

    // ==================== 格式化操作 ====================

    /**
     * @brief 设置单元格数字格式
     * @param row 行号
     * @param col 列号
     * @param formatType 格式类型 (修正此处)
     * @param decimalPlaces 小数位数
     * @return 成功返回true，失败返回false
     */
    bool setCellNumberFormat(row_t row, column_t col,
                           TXNumberFormat::FormatType formatType, int decimalPlaces = 2); // TXCell::NumberFormat -> TXNumberFormat::FormatType

    /**
     * @brief 设置单元格自定义格式
     * @param row 行号
     * @param col 列号
     * @param formatString 格式字符串
     * @return 成功返回true，失败返回false
     */
    bool setCellCustomFormat(row_t row, column_t col,
                           const std::string& formatString);

    /**
     * @brief 设置范围内单元格的格式
     * @param range 范围
     * @param formatType 格式类型
     * @param decimalPlaces 小数位数
     * @return 成功设置的单元格数量
     */
    std::size_t setRangeNumberFormat(const Range& range, TXNumberFormat::FormatType formatType, // TXCell::NumberFormat -> TXNumberFormat::FormatType
                                   int decimalPlaces = 2);

    /**
     * @brief 获取单元格格式化后的值
     * @param row 行号
     * @param col 列号
     * @return 格式化后的字符串
     */
    std::string getCellFormattedValue(row_t row, column_t col) const;

    /**
     * @brief 批量设置格式（高性能版本）
     * @param formats 坐标到格式的映射
     * @return 成功设置的格式数量
     */
    std::size_t setCellFormats(const std::vector<std::pair<Coordinate, TXNumberFormat::FormatType>>& formats); // TXCell::NumberFormat -> TXNumberFormat::FormatType

    // ==================== 样式操作 ====================
    
    /**
     * @brief 设置单元格样式
     * @param row 行号
     * @param col 列号
     * @param style 单元格样式
     * @return 成功返回true，失败返回false
     */
    bool setCellStyle(row_t row, column_t col, const TXCellStyle& style);

    /**
     * @brief 设置单元格样式（使用A1格式）
     * @param address 单元格地址，如"A1"
     * @param style 单元格样式
     * @return 成功返回true，失败返回false
     */
    bool setCellStyle(const std::string& address, const TXCellStyle& style);

    /**
     * @brief 设置范围内单元格的样式
     * @param range 范围
     * @param style 单元格样式
     * @return 成功设置的单元格数量
     */
    std::size_t setRangeStyle(const Range& range, const TXCellStyle& style);

    /**
     * @brief 批量设置样式（高性能版本）
     * @param styles 坐标到样式的映射
     * @return 成功设置的样式数量
     */
    std::size_t setCellStyles(const std::vector<std::pair<Coordinate, TXCellStyle>>& styles);

    /**
     * @brief 批量设置数字格式（高性能版本）
     * @param formats 坐标到数字格式的映射
     * @return 成功设置的格式数量
     */
    std::size_t setBatchNumberFormats(const std::vector<std::pair<Coordinate, TXCellStyle::NumberFormatDefinition>>& formats);

    // ==================== 列宽和行高操作 ====================
    
    /**
     * @brief 设置列宽
     * @param col 列号（1开始）
     * @param width 宽度（字符单位）
     * @return 成功返回true，失败返回false
     */
    bool setColumnWidth(column_t col, double width);
    
    /**
     * @brief 获取列宽
     * @param col 列号（1开始）
     * @return 列宽（字符单位），默认宽度为8.43
     */
    double getColumnWidth(column_t col) const;
    
    /**
     * @brief 设置行高
     * @param row 行号（1开始）
     * @param height 高度（磅数）
     * @return 成功返回true，失败返回false
     */
    bool setRowHeight(row_t row, double height);
    
    /**
     * @brief 获取行高
     * @param row 行号（1开始）
     * @return 行高（磅数），默认高度为15.0
     */
    double getRowHeight(row_t row) const;
    
    /**
     * @brief 自动调整列宽以适应内容
     * @param col 列号（1开始）
     * @param minWidth 最小宽度（字符单位）
     * @param maxWidth 最大宽度（字符单位）
     * @return 调整后的列宽
     */
    double autoFitColumnWidth(column_t col, double minWidth = 1.0, double maxWidth = 255.0);
    
    /**
     * @brief 自动调整行高以适应内容
     * @param row 行号（1开始）
     * @param minHeight 最小高度（磅数）
     * @param maxHeight 最大高度（磅数）
     * @return 调整后的行高
     */
    double autoFitRowHeight(row_t row, double minHeight = 12.0, double maxHeight = 409.0);
    
    /**
     * @brief 自动调整所有列宽
     * @param minWidth 最小宽度
     * @param maxWidth 最大宽度
     * @return 调整的列数
     */
    std::size_t autoFitAllColumnWidths(double minWidth = 1.0, double maxWidth = 255.0);
    
    /**
     * @brief 自动调整所有行高
     * @param minHeight 最小高度
     * @param maxHeight 最大高度
     * @return 调整的行数
     */
    std::size_t autoFitAllRowHeights(double minHeight = 12.0, double maxHeight = 409.0);

    // ==================== 工作表保护功能 ====================
    
    /**
     * @brief 工作表保护选项（兼容性别名）
     */
    using SheetProtection = TXSheetProtectionManager::SheetProtection;
    
    /**
     * @brief 保护工作表
     * @param password 保护密码（空字符串表示无密码）
     * @param protection 保护选项
     * @return 成功返回true，失败返回false
     */
    bool protectSheet(const std::string& password = "", const SheetProtection& protection = SheetProtection{});
    
    /**
     * @brief 取消工作表保护
     * @param password 解除保护的密码
     * @return 成功返回true，失败返回false
     */
    bool unprotectSheet(const std::string& password = "");
    
    /**
     * @brief 检查工作表是否受保护
     * @return 受保护返回true，否则返回false
     */
    bool isSheetProtected() const;
    
    /**
     * @brief 获取工作表保护设置
     * @return 保护设置
     */
    const SheetProtection& getSheetProtection() const;
    
    /**
     * @brief 设置单元格锁定状态
     * @param row 行号
     * @param col 列号
     * @param locked 是否锁定
     * @return 成功返回true，失败返回false
     */
    bool setCellLocked(row_t row, column_t col, bool locked);
    
    /**
     * @brief 检查单元格是否锁定
     * @param row 行号
     * @param col 列号
     * @return 锁定返回true，否则返回false
     */
    bool isCellLocked(row_t row, column_t col) const;
    
    /**
     * @brief 设置范围内单元格的锁定状态
     * @param range 范围
     * @param locked 是否锁定
     * @return 成功设置的单元格数量
     */
    std::size_t setRangeLocked(const Range& range, bool locked);

    // ==================== 增强公式支持 ====================
    
    /**
     * @brief 公式计算选项（兼容性别名）
     */
    using FormulaCalculationOptions = TXFormulaManager::FormulaCalculationOptions;
    
    /**
     * @brief 设置公式计算选项
     * @param options 计算选项
     */
    void setFormulaCalculationOptions(const FormulaCalculationOptions& options);
    
    /**
     * @brief 获取公式计算选项
     * @return 计算选项
     */
    const FormulaCalculationOptions& getFormulaCalculationOptions() const;
    
    /**
     * @brief 添加命名范围
     * @param name 名称
     * @param range 范围
     * @param comment 注释
     * @return 成功返回true，失败返回false
     */
    bool addNamedRange(const std::string& name, const Range& range, const std::string& comment = "");
    
    /**
     * @brief 删除命名范围
     * @param name 名称
     * @return 成功返回true，失败返回false
     */
    bool removeNamedRange(const std::string& name);
    
    /**
     * @brief 获取命名范围
     * @param name 名称
     * @return 范围，如果不存在返回无效范围
     */
    Range getNamedRange(const std::string& name) const;
    
    /**
     * @brief 获取所有命名范围
     * @return 名称到范围的映射
     */
    std::unordered_map<std::string, Range> getAllNamedRanges() const;
    
    /**
     * @brief 检测循环引用
     * @return 发现循环引用返回true，否则返回false
     */
    bool detectCircularReferences() const;
    
    // ==================== 哈希函数 ====================
    
    /**
     * @brief 坐标哈希函数
     */
    struct CoordinateHash {
        std::size_t operator()(const TXCoordinate& coord) const {
            return std::hash<u32>()(coord.getRow().index()) ^
                (std::hash<u32>()(coord.getCol().index()) << 1);
        }
    };
    
    /**
     * @brief 获取公式依赖关系图
     * @return 单元格坐标到其依赖的单元格列表的映射
     */
    std::unordered_map<Coordinate, std::vector<Coordinate>, CoordinateHash> getFormulaDependencies() const;

    // ==================== 便捷方法 ====================

    /**
     * @brief 从A1格式地址创建坐标
     * @param address A1格式地址，如"A1", "B2"
     * @return 坐标对象
     */
    static Coordinate addressToCoordinate(const std::string& address) {
        return Coordinate::fromAddress(address);
    }

    /**
     * @brief 坐标转换为A1格式地址
     * @param coord 坐标
     * @return A1格式地址
     */
    static std::string coordinateToAddress(const Coordinate& coord) {
        return coord.toAddress();
    }

    /**
     * @brief 从A1格式范围地址创建范围
     * @param rangeAddress 范围地址，如"A1:B5"
     * @return 范围对象
     */
    static Range addressToRange(const std::string& rangeAddress) {
        return Range::fromAddress(rangeAddress);
    }

    /**
     * @brief 范围转换为A1格式地址
     * @param range 范围
     * @return A1格式范围地址
     */
    static std::string rangeToAddress(const Range& range) {
        return range.toAddress();
    }

    /**
     * @brief 获取最后的错误信息
     * @return 错误信息字符串
     */
    const std::string& getLastError() const;

    /**
    * @brief 获取父工作簿对象
    * @return TXWorkbook 指针
    */
    TXWorkbook* getWorkbook() const { return workbook_; }

    // ==================== 管理器访问接口（高级用法）====================

    /**
     * @brief 获取单元格管理器
     * @return 单元格管理器引用
     */
    TXCellManager& getCellManager() { return cellManager_; }
    const TXCellManager& getCellManager() const { return cellManager_; }

    /**
     * @brief 获取行列管理器
     * @return 行列管理器引用
     */
    TXRowColumnManager& getRowColumnManager() { return rowColumnManager_; }
    const TXRowColumnManager& getRowColumnManager() const { return rowColumnManager_; }

    /**
     * @brief 获取保护管理器
     * @return 保护管理器引用
     */
    TXSheetProtectionManager& getProtectionManager() { return protectionManager_; }
    const TXSheetProtectionManager& getProtectionManager() const { return protectionManager_; }

    /**
     * @brief 获取公式管理器
     * @return 公式管理器引用
     */
    TXFormulaManager& getFormulaManager() { return formulaManager_; }
    const TXFormulaManager& getFormulaManager() const { return formulaManager_; }

    /**
     * @brief 获取合并单元格管理器
     * @return 合并单元格管理器引用
     */
    TXMergedCells& getMergedCells() { return mergedCells_; }
    const TXMergedCells& getMergedCells() const { return mergedCells_; }


private:
    // ==================== 基本属性 ====================
    std::string name_;                              ///< 工作表名称
    TXWorkbook* workbook_ = nullptr;                ///< 父工作簿指针
    mutable std::string lastError_;                 ///< 最后的错误信息

    // ==================== 管理器组件 ====================
    TXCellManager cellManager_;                     ///< 单元格管理器
    TXRowColumnManager rowColumnManager_;           ///< 行列管理器
    TXSheetProtectionManager protectionManager_;   ///< 保护管理器
    TXFormulaManager formulaManager_;               ///< 公式管理器
    TXMergedCells mergedCells_;                     ///< 合并单元格管理器

    // ==================== 兼容性保留 ====================
    // 为了保持API兼容性，保留一些原有的结构定义
    // 但实际数据存储在管理器中
    
    // ==================== 私有辅助方法 ====================
    
    /**
     * @brief 获取单元格（内部实现）
     */
    TXCell* getCellInternal(const Coordinate& coord);
    const TXCell* getCellInternal(const Coordinate& coord) const;
    
    /**
     * @brief 应用数字格式到单元格（内部辅助方法）
     * @param cell 目标单元格
     * @param numFmtId 数字格式ID
     * @return 成功返回true，失败返回false
     */
    bool applyCellNumberFormat(TXCell* cell, u32 numFmtId);
    
    /**
     * @brief 获取单元格当前的有效样式
     * @param cell 目标单元格
     * @return 当前的完整样式对象
     */
    TXCellStyle getCellEffectiveStyle(TXCell* cell);
    
    /**
     * @brief 更新已使用范围
     */
    void updateUsedRange();
    
    /**
     * @brief 计算文本内容的显示宽度
     * @param text 文本内容
     * @param fontSize 字体大小
     * @param fontName 字体名称
     * @return 显示宽度（字符单位）
     */
    double calculateTextWidth(const std::string& text, double fontSize = 11.0, const std::string& fontName = "Calibri") const;
    
    /**
     * @brief 计算文本内容的显示高度
     * @param text 文本内容
     * @param fontSize 字体大小
     * @param columnWidth 列宽（用于自动换行计算）
     * @return 显示高度（磅数）
     */
    double calculateTextHeight(const std::string& text, double fontSize = 11.0, double columnWidth = 8.43) const;
    
    /**
     * @brief 生成密码哈希
     * @param password 原始密码
     * @return MD5哈希值
     */
    std::string generatePasswordHash(const std::string& password) const;
    
    /**
     * @brief 验证密码
     * @param password 输入密码
     * @param hash 存储的哈希值
     * @return 验证成功返回true，否则返回false
     */
    bool verifyPassword(const std::string& password, const std::string& hash) const;
    
    /**
     * @brief 检查操作是否被保护设置阻止
     * @param operation 操作类型
     * @return 被阻止返回true，否则返回false
     */
    bool isOperationBlocked(const std::string& operation) const;
    

    
    /**
     * @brief 解析公式中的单元格引用
     * @param formula 公式字符串
     * @return 引用的单元格坐标列表
     */
    std::vector<Coordinate> parseFormulaReferences(const std::string& formula) const;

    /**
     * @brief 设置错误信息
     * @param error 错误信息
     */
    void setError(const std::string& error) const { lastError_ = error; }

    /**
     * @brief 清除错误信息
     */
    void clearError() const { lastError_.clear(); }

    /**
     * @brief 通知组件变化
     * @param component 变化的组件
     */
    void notifyComponentChange(ExcelComponent component) const;
};

} // namespace TinaXlsx 
