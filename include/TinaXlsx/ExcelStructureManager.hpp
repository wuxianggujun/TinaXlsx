/**
 * @file ExcelStructureManager.hpp
 * @brief Excel文件结构管理器 - 专门负责Excel文件结构的解析和管理
 */

#pragma once

#include "Types.hpp"
#include "ZipReader.hpp"
#include "XmlParser.hpp"
#include <memory>
#include <vector>
#include <unordered_map>
#include <optional>

namespace TinaXlsx {

/**
 * @brief Excel文件结构管理器
 * 专门负责Excel文件结构的解析和管理，包括工作表信息、共享字符串、关系等
 */
class ExcelStructureManager {
public:
    /**
     * @brief 工作表信息结构
     */
    struct SheetInfo {
        std::string name;           ///< 工作表名称
        std::string relationId;     ///< 关系ID
        std::string filePath;       ///< 工作表文件路径
        RowIndex sheetId;          ///< 工作表ID
        
        /**
         * @brief 生成缓存键
         * @return std::string 用于缓存的键
         */
        std::string getCacheKey() const {
            return name + "|" + filePath;
        }
    };

private:
    std::unique_ptr<ExcelZipReader> zipReader_;
    std::vector<SheetInfo> sheets_;
    std::vector<std::string> sharedStrings_;
    std::unordered_map<std::string, std::string> relationships_;
    bool initialized_ = false;

public:
    /**
     * @brief 构造函数
     * @param zipReader ZIP读取器（接管所有权）
     */
    explicit ExcelStructureManager(std::unique_ptr<ExcelZipReader> zipReader);
    
    /**
     * @brief 析构函数
     */
    ~ExcelStructureManager() = default;
    
    // 禁用拷贝，启用移动
    ExcelStructureManager(const ExcelStructureManager&) = delete;
    ExcelStructureManager& operator=(const ExcelStructureManager&) = delete;
    ExcelStructureManager(ExcelStructureManager&&) = default;
    ExcelStructureManager& operator=(ExcelStructureManager&&) = default;
    
    /**
     * @brief 获取工作表列表
     * @return const std::vector<SheetInfo>& 工作表信息列表
     */
    const std::vector<SheetInfo>& getSheets() const;
    
    /**
     * @brief 获取共享字符串列表
     * @return const std::vector<std::string>& 共享字符串列表
     */
    const std::vector<std::string>& getSharedStrings() const;
    
    /**
     * @brief 获取ZIP读取器
     * @return ExcelZipReader* ZIP读取器指针
     */
    ExcelZipReader* getZipReader() const;
    
    /**
     * @brief 通过名称查找工作表
     * @param name 工作表名称
     * @return std::optional<SheetInfo> 工作表信息，如果不存在则返回空
     */
    std::optional<SheetInfo> findSheetByName(const std::string& name) const;
    
    /**
     * @brief 通过索引获取工作表
     * @param index 工作表索引（0开始）
     * @return std::optional<SheetInfo> 工作表信息，如果索引无效则返回空
     */
    std::optional<SheetInfo> getSheetByIndex(size_t index) const;
    
    /**
     * @brief 获取工作表数量
     * @return size_t 工作表数量
     */
    size_t getSheetCount() const;
    
    /**
     * @brief 获取工作表名称列表
     * @return std::vector<std::string> 工作表名称列表
     */
    std::vector<std::string> getSheetNames() const;
    
    /**
     * @brief 检查Excel文件是否有效
     * @return bool 是否有效
     */
    bool isValidExcelFile() const;

private:
    /**
     * @brief 确保已初始化（延迟初始化）
     */
    void ensureInitialized() const;
    
    /**
     * @brief 解析Excel文件结构
     */
    void parseStructure();
    
    /**
     * @brief 解析工作簿信息
     */
    void parseWorkbook();
    
    /**
     * @brief 解析关系文件
     */
    void parseRelationships();
    
    /**
     * @brief 解析共享字符串
     */
    void parseSharedStrings();
};

} // namespace TinaXlsx 