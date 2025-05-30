#pragma once

#include "TXTypes.hpp"
#include "TXStyle.hpp"
#include "TXRange.hpp"
#include "TXConditionalFormat.hpp"
#include "TXStyleTemplate.hpp"
#include <string>
#include <vector>
#include <memory>
#include <functional>

namespace TinaXlsx {

// Forward declarations
class TXSheet;
class TXCell;

/**
 * @brief 格式应用模式枚举
 */
enum class FormatApplyMode : uint8_t {
    Replace = 0,        ///< 替换现有格式
    Merge = 1,          ///< 合并格式（只更新指定属性）
    Overlay = 2         ///< 覆盖格式（保留未指定的属性）
};

/**
 * @brief 格式应用选项
 */
struct FormatApplyOptions {
    FormatApplyMode mode;           ///< 应用模式
    bool applyFont;                 ///< 应用字体
    bool applyAlignment;            ///< 应用对齐
    bool applyBorder;               ///< 应用边框
    bool applyFill;                 ///< 应用填充
    bool applyNumberFormat;         ///< 应用数字格式
    bool applyConditionalFormat;    ///< 应用条件格式
    bool preserveFormulas;          ///< 保留公式
    bool preserveValues;            ///< 保留值
    
    FormatApplyOptions() 
        : mode(FormatApplyMode::Replace)
        , applyFont(true)
        , applyAlignment(true)
        , applyBorder(true)
        , applyFill(true)
        , applyNumberFormat(true)
        , applyConditionalFormat(true)
        , preserveFormulas(true)
        , preserveValues(true) {}
};

/**
 * @brief 格式应用过滤器函数类型
 */
using FormatFilter = std::function<bool(const TXCell& cell, row_t row, column_t col)>;

/**
 * @brief 批量格式应用器
 */
class TXBatchFormatApplicator {
public:
    TXBatchFormatApplicator();
    ~TXBatchFormatApplicator();
    
    // ==================== 基本格式应用 ====================
    
    /**
     * @brief 应用样式到范围
     * @param sheet 工作表
     * @param range 范围
     * @param style 样式
     * @param options 应用选项
     * @return 成功应用的单元格数量
     */
    static size_t applyStyleToRange(TXSheet& sheet, const TXRange& range, const TXCellStyle& style, const FormatApplyOptions& options = FormatApplyOptions{});
    
    /**
     * @brief 应用样式到多个范围
     * @param sheet 工作表
     * @param ranges 范围列表
     * @param style 样式
     * @param options 应用选项
     * @return 成功应用的单元格数量
     */
    static size_t applyStyleToRanges(TXSheet& sheet, const std::vector<TXRange>& ranges, const TXCellStyle& style, const FormatApplyOptions& options = FormatApplyOptions{});
    
    /**
     * @brief 应用样式到行
     * @param sheet 工作表
     * @param startRow 开始行
     * @param endRow 结束行
     * @param style 样式
     * @param options 应用选项
     * @return 成功应用的单元格数量
     */
    static size_t applyStyleToRows(TXSheet& sheet, row_t startRow, row_t endRow, const TXCellStyle& style, const FormatApplyOptions& options = FormatApplyOptions{});
    
    /**
     * @brief 应用样式到列
     * @param sheet 工作表
     * @param startCol 开始列
     * @param endCol 结束列
     * @param style 样式
     * @param options 应用选项
     * @return 成功应用的单元格数量
     */
    static size_t applyStyleToColumns(TXSheet& sheet, column_t startCol, column_t endCol, const TXCellStyle& style, const FormatApplyOptions& options = FormatApplyOptions{});
    
    // ==================== 条件格式应用 ====================
    
    /**
     * @brief 应用条件格式到范围
     * @param sheet 工作表
     * @param range 范围
     * @param conditionalFormat 条件格式管理器
     * @return 成功应用的单元格数量
     */
    static size_t applyConditionalFormatToRange(TXSheet& sheet, const TXRange& range, const TXConditionalFormatManager& conditionalFormat);
    
    // ==================== 模板应用 ====================
    
    /**
     * @brief 应用样式模板到范围
     * @param sheet 工作表
     * @param range 范围
     * @param styleTemplate 样式模板
     * @param options 应用选项
     * @return 成功应用的单元格数量
     */
    static size_t applyTemplateToRange(TXSheet& sheet, const TXRange& range, const TXStyleTemplate& styleTemplate, const FormatApplyOptions& options = FormatApplyOptions{});
    
    /**
     * @brief 应用模板中的命名样式
     * @param sheet 工作表
     * @param range 范围
     * @param styleTemplate 样式模板
     * @param styleName 样式名称
     * @param options 应用选项
     * @return 成功应用的单元格数量
     */
    static size_t applyNamedStyleToRange(TXSheet& sheet, const TXRange& range, const TXStyleTemplate& styleTemplate, const std::string& styleName, const FormatApplyOptions& options = FormatApplyOptions{});
    
    // ==================== 智能格式应用 ====================
    
    /**
     * @brief 自动格式化表格
     * @param sheet 工作表
     * @param range 表格范围
     * @param hasHeader 是否有标题行
     * @param templateId 模板ID
     * @return 成功返回true
     */
    static bool autoFormatTable(TXSheet& sheet, const TXRange& range, bool hasHeader = true, const std::string& templateId = "data_table");
    
    /**
     * @brief 应用交替行颜色
     * @param sheet 工作表
     * @param range 范围
     * @param color1 第一种颜色
     * @param color2 第二种颜色
     * @param startWithColor1 是否从第一种颜色开始
     * @return 成功应用的行数
     */
    static size_t applyAlternatingRowColors(TXSheet& sheet, const TXRange& range, const TXColor& color1, const TXColor& color2, bool startWithColor1 = true);
    
    /**
     * @brief 应用交替列颜色
     * @param sheet 工作表
     * @param range 范围
     * @param color1 第一种颜色
     * @param color2 第二种颜色
     * @param startWithColor1 是否从第一种颜色开始
     * @return 成功应用的列数
     */
    static size_t applyAlternatingColumnColors(TXSheet& sheet, const TXRange& range, const TXColor& color1, const TXColor& color2, bool startWithColor1 = true);
    
    // ==================== 高级应用方法 ====================
    
    /**
     * @brief 带过滤器的样式应用
     * @param sheet 工作表
     * @param range 范围
     * @param style 样式
     * @param filter 过滤器函数
     * @param options 应用选项
     * @return 成功应用的单元格数量
     */
    static size_t applyStyleWithFilter(TXSheet& sheet, const TXRange& range, const TXCellStyle& style, FormatFilter filter, const FormatApplyOptions& options = FormatApplyOptions{});
    
    /**
     * @brief 基于值的条件样式应用
     * @param sheet 工作表
     * @param range 范围
     * @param valueStyleMap 值到样式的映射
     * @param defaultStyle 默认样式
     * @param options 应用选项
     * @return 成功应用的单元格数量
     */
    static size_t applyStyleByValue(TXSheet& sheet, const TXRange& range, const std::unordered_map<std::string, TXCellStyle>& valueStyleMap, const TXCellStyle& defaultStyle, const FormatApplyOptions& options = FormatApplyOptions{});
    
    // ==================== 格式复制与粘贴 ====================
    
    /**
     * @brief 复制格式
     * @param sheet 工作表
     * @param sourceRange 源范围
     * @return 格式信息
     */
    static std::vector<std::vector<TXCellStyle>> copyFormat(const TXSheet& sheet, const TXRange& sourceRange);
    
    /**
     * @brief 粘贴格式
     * @param sheet 工作表
     * @param targetRange 目标范围
     * @param formats 格式信息
     * @param options 应用选项
     * @return 成功应用的单元格数量
     */
    static size_t pasteFormat(TXSheet& sheet, const TXRange& targetRange, const std::vector<std::vector<TXCellStyle>>& formats, const FormatApplyOptions& options = FormatApplyOptions{});
    
    /**
     * @brief 复制格式并粘贴
     * @param sheet 工作表
     * @param sourceRange 源范围
     * @param targetRange 目标范围
     * @param options 应用选项
     * @return 成功应用的单元格数量
     */
    static size_t copyPasteFormat(TXSheet& sheet, const TXRange& sourceRange, const TXRange& targetRange, const FormatApplyOptions& options = FormatApplyOptions{});
    
    // ==================== 格式清除 ====================
    
    /**
     * @brief 清除格式
     * @param sheet 工作表
     * @param range 范围
     * @param clearOptions 清除选项
     * @return 成功清除的单元格数量
     */
    static size_t clearFormat(TXSheet& sheet, const TXRange& range, const FormatApplyOptions& clearOptions = FormatApplyOptions{});
    
    /**
     * @brief 重置为默认格式
     * @param sheet 工作表
     * @param range 范围
     * @return 成功重置的单元格数量
     */
    static size_t resetToDefaultFormat(TXSheet& sheet, const TXRange& range);

private:
    // 辅助方法
    static void applySingleCellStyle(TXCell& cell, const TXCellStyle& style, const FormatApplyOptions& options);
    static void mergeCellStyle(TXCellStyle& target, const TXCellStyle& source, const FormatApplyOptions& options);
    static bool shouldApplyToCell(const TXCell& cell, row_t row, column_t col, FormatFilter filter);
};

/**
 * @brief 格式批处理任务
 */
class TXFormatBatchTask {
public:
    /**
     * @brief 任务类型枚举
     */
    enum class TaskType {
        ApplyStyle,             ///< 应用样式
        ApplyTemplate,          ///< 应用模板
        ApplyConditionalFormat, ///< 应用条件格式
        ClearFormat,            ///< 清除格式
        CopyFormat,             ///< 复制格式
        AutoFormat              ///< 自动格式
    };
    
    explicit TXFormatBatchTask(TaskType type);
    ~TXFormatBatchTask();
    
    /**
     * @brief 设置目标范围
     * @param range 范围
     */
    void setTargetRange(const TXRange& range);
    
    /**
     * @brief 设置应用选项
     * @param options 选项
     */
    void setOptions(const FormatApplyOptions& options);
    
    /**
     * @brief 设置样式
     * @param style 样式
     */
    void setStyle(const TXCellStyle& style);
    
    /**
     * @brief 设置模板
     * @param styleTemplate 样式模板
     */
    void setTemplate(const TXStyleTemplate& styleTemplate);
    
    /**
     * @brief 执行任务
     * @param sheet 工作表
     * @return 成功返回true
     */
    bool execute(TXSheet& sheet);
    
    /**
     * @brief 获取任务类型
     * @return 任务类型
     */
    TaskType getType() const;

private:
    TaskType type_;
    TXRange targetRange_;
    FormatApplyOptions options_;
    TXCellStyle style_;
};

/**
 * @brief 批量格式任务管理器
 */
class TXBatchFormatTaskManager {
public:
    TXBatchFormatTaskManager();
    ~TXBatchFormatTaskManager();
    
    /**
     * @brief 添加任务
     * @param task 任务
     */
    void addTask(std::unique_ptr<TXFormatBatchTask> task);
    
    /**
     * @brief 清除所有任务
     */
    void clearTasks();
    
    /**
     * @brief 执行所有任务
     * @param sheet 工作表
     * @return 成功执行的任务数量
     */
    size_t executeAllTasks(TXSheet& sheet);
    
    /**
     * @brief 获取任务数量
     * @return 任务数量
     */
    size_t getTaskCount() const;

private:
    std::vector<std::unique_ptr<TXFormatBatchTask>> tasks_;
    size_t currentTaskIndex_;
};

} // namespace TinaXlsx