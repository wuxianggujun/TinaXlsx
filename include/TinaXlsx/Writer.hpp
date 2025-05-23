/**
 * @file Writer.hpp
 * @brief 高性能Excel写入器
 */

#pragma once

#include "Types.hpp"
#include "Exception.hpp"
#include "Format.hpp"
#include <xlsxwriter.h>
#include <memory>
#include <string>
#include <vector>
#include <optional>

namespace TinaXlsx {

// 前向声明
class Worksheet;

/**
 * @brief 高性能Excel写入器
 * 
 * 基于libxlsxwriter的高性能写入器，支持批量写入和流式写入
 * 使用RAII管理资源，保证异常安全
 */
class Writer {
private:
    struct Impl;
    std::unique_ptr<Impl> pImpl_;

public:
    /**
     * @brief 构造函数
     * @param filePath Excel文件路径
     * @throws FileException 如果文件无法创建
     */
    explicit Writer(const std::string& filePath);
    
    /**
     * @brief 移动构造函数
     */
    Writer(Writer&& other) noexcept;
    
    /**
     * @brief 移动赋值操作符
     */
    Writer& operator=(Writer&& other) noexcept;
    
    /**
     * @brief 析构函数
     */
    ~Writer();
    
    // 禁用拷贝构造和拷贝赋值
    Writer(const Writer&) = delete;
    Writer& operator=(const Writer&) = delete;
    
    /**
     * @brief 创建新的工作表
     * @param sheetName 工作表名称，如果为空则使用默认名称
     * @return std::shared_ptr<Worksheet> 工作表对象
     */
    [[nodiscard]] std::shared_ptr<Worksheet> createWorksheet(const std::string& sheetName = "");
    
    /**
     * @brief 获取指定名称的工作表
     * @param sheetName 工作表名称
     * @return std::shared_ptr<Worksheet> 工作表对象，如果不存在则返回nullptr
     */
    [[nodiscard]] std::shared_ptr<Worksheet> getWorksheet(const std::string& sheetName);
    
    /**
     * @brief 获取指定索引的工作表
     * @param index 工作表索引（0基于）
     * @return std::shared_ptr<Worksheet> 工作表对象，如果不存在则返回nullptr
     */
    [[nodiscard]] std::shared_ptr<Worksheet> getWorksheet(SheetIndex index);
    
    /**
     * @brief 获取所有工作表名称
     * @return std::vector<std::string> 工作表名称列表
     */
    [[nodiscard]] std::vector<std::string> getWorksheetNames() const;
    
    /**
     * @brief 获取工作表数量
     * @return size_t 工作表数量
     */
    [[nodiscard]] size_t getWorksheetCount() const;
    
    /**
     * @brief 创建格式对象
     * @return std::unique_ptr<Format> 格式对象
     */
    [[nodiscard]] std::unique_ptr<Format> createFormat();
    
    /**
     * @brief 获取格式构建器
     * @return FormatBuilder 格式构建器
     */
    [[nodiscard]] FormatBuilder getFormatBuilder();
    
    /**
     * @brief 设置文档属性
     * @param title 标题
     * @param subject 主题
     * @param author 作者
     * @param manager 管理者
     * @param company 公司
     * @param category 类别
     * @param keywords 关键词
     * @param comments 备注
     */
    void setDocumentProperties(
        const std::string& title = "",
        const std::string& subject = "",
        const std::string& author = "",
        const std::string& manager = "",
        const std::string& company = "",
        const std::string& category = "",
        const std::string& keywords = "",
        const std::string& comments = "");
    
    /**
     * @brief 设置自定义属性
     * @param name 属性名称
     * @param value 属性值
     */
    void setCustomProperty(const std::string& name, const std::string& value);
    
    /**
     * @brief 设置自定义属性（数字）
     * @param name 属性名称
     * @param value 属性值
     */
    void setCustomProperty(const std::string& name, double value);
    
    /**
     * @brief 设置自定义属性（布尔值）
     * @param name 属性名称
     * @param value 属性值
     */
    void setCustomProperty(const std::string& name, bool value);
    
    /**
     * @brief 保存文件
     * @return bool 是否成功保存
     */
    bool save();
    
    /**
     * @brief 关闭文件（自动调用save）
     * @return bool 是否成功关闭
     */
    bool close();
    
    /**
     * @brief 获取内部工作簿对象指针（仅供内部使用）
     * @return lxw_workbook* 内部工作簿对象指针
     */
    [[nodiscard]] lxw_workbook* getInternalWorkbook() const;
    
    /**
     * @brief 检查文件是否已关闭
     * @return bool 是否已关闭
     */
    [[nodiscard]] bool isClosed() const;
    
    /**
     * @brief 设置默认日期格式
     * @param format 日期格式字符串
     */
    void setDefaultDateFormat(const std::string& format);
    
    /**
     * @brief 定义命名区域
     * @param name 区域名称
     * @param sheetName 工作表名称
     * @param range 单元格范围
     */
    void defineNamedRange(const std::string& name, const std::string& sheetName, const CellRange& range);
};

} // namespace TinaXlsx 