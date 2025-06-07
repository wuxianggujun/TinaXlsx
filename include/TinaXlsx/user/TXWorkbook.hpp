//
// @file TXWorkbook.hpp
// @brief 🚀 用户层工作簿类 - 多工作表管理和文件操作
//

#pragma once

#include <TinaXlsx/user/TXSheet.hpp>
#include <TinaXlsx/TXResult.hpp>
#include <TinaXlsx/TXError.hpp>
#include <TinaXlsx/TXUnifiedMemoryManager.hpp>
#include <TinaXlsx/TXGlobalStringPool.hpp>
#include <TinaXlsx/TXVector.hpp>
#include <string>
#include <vector>
#include <memory>
#include <unordered_map>

namespace TinaXlsx {

/**
 * @brief 🚀 用户层工作簿类
 * 
 * 设计理念：
 * - 管理多个TXSheet工作表
 * - 提供Excel文件的保存和加载功能
 * - 简单直观的工作簿操作API
 * - 完整的错误处理和资源管理
 * 
 * 使用示例：
 * ```cpp
 * // 创建新工作簿
 * auto workbook = TXWorkbook::create("我的工作簿");
 * 
 * // 添加工作表
 * auto sheet1 = workbook->addSheet("销售数据");
 * auto sheet2 = workbook->addSheet("统计分析");
 * 
 * // 操作工作表
 * sheet1->cell("A1").setValue("产品名称");
 * sheet1->cell("B1").setValue("销售额");
 * 
 * // 保存文件
 * workbook->saveAs("report.xlsx");
 * 
 * // 加载文件
 * auto loaded = TXWorkbook::load("existing.xlsx");
 * ```
 */
class TXWorkbook {
public:
    // ==================== 静态工厂方法 ====================
    
    /**
     * @brief 🚀 创建新工作簿
     * @param name 工作簿名称
     * @return 工作簿智能指针
     */
    static std::unique_ptr<TXWorkbook> create(const std::string& name = "新建工作簿");
    
    /**
     * @brief 🚀 从文件加载工作簿
     * @param file_path 文件路径
     * @return 工作簿智能指针或错误
     */
    static TXResult<std::unique_ptr<TXWorkbook>> load(const std::string& file_path);
    
    /**
     * @brief 🚀 从内存数据加载工作簿
     * @param data 内存数据
     * @param size 数据大小
     * @return 工作簿智能指针或错误
     */
    static TXResult<std::unique_ptr<TXWorkbook>> loadFromMemory(const void* data, size_t size);

    // ==================== 构造和析构 ====================
    
    /**
     * @brief 🚀 构造函数
     * @param name 工作簿名称
     */
    explicit TXWorkbook(const std::string& name);
    
    /**
     * @brief 🚀 析构函数
     */
    ~TXWorkbook();
    
    // 禁用拷贝，支持移动
    TXWorkbook(const TXWorkbook&) = delete;
    TXWorkbook& operator=(const TXWorkbook&) = delete;
    TXWorkbook(TXWorkbook&&) noexcept;
    TXWorkbook& operator=(TXWorkbook&&) noexcept;

    // ==================== 工作簿属性 ====================
    
    /**
     * @brief 🚀 获取工作簿名称
     */
    const std::string& getName() const { return name_; }
    
    /**
     * @brief 🚀 设置工作簿名称
     */
    void setName(const std::string& name) { name_ = name; }
    
    /**
     * @brief 🚀 获取工作表数量
     */
    size_t getSheetCount() const { return sheets_.size(); }
    
    /**
     * @brief 🚀 检查是否为空
     */
    bool isEmpty() const { return sheets_.empty(); }
    
    /**
     * @brief 🚀 获取当前活动工作表索引
     */
    int getActiveSheetIndex() const { return active_sheet_index_; }

    // ==================== 工作表管理 ====================
    
    /**
     * @brief 🚀 添加新工作表
     * @param name 工作表名称
     * @return 工作表指针
     */
    TXSheet* addSheet(const std::string& name);
    
    /**
     * @brief 🚀 插入工作表到指定位置
     * @param index 插入位置
     * @param name 工作表名称
     * @return 工作表指针或错误
     */
    TXResult<TXSheet*> insertSheet(size_t index, const std::string& name);
    
    /**
     * @brief 🚀 删除工作表
     * @param index 工作表索引
     * @return 操作结果
     */
    TXResult<void> removeSheet(size_t index);
    
    /**
     * @brief 🚀 删除工作表
     * @param name 工作表名称
     * @return 操作结果
     */
    TXResult<void> removeSheet(const std::string& name);
    
    /**
     * @brief 🚀 重命名工作表
     * @param index 工作表索引
     * @param new_name 新名称
     * @return 操作结果
     */
    TXResult<void> renameSheet(size_t index, const std::string& new_name);
    
    /**
     * @brief 🚀 重命名工作表
     * @param old_name 旧名称
     * @param new_name 新名称
     * @return 操作结果
     */
    TXResult<void> renameSheet(const std::string& old_name, const std::string& new_name);
    
    /**
     * @brief 🚀 移动工作表
     * @param from_index 源位置
     * @param to_index 目标位置
     * @return 操作结果
     */
    TXResult<void> moveSheet(size_t from_index, size_t to_index);

    // ==================== 工作表访问 ====================
    
    /**
     * @brief 🚀 获取工作表（按索引）
     * @param index 工作表索引
     * @return 工作表指针
     */
    TXSheet* getSheet(size_t index);
    const TXSheet* getSheet(size_t index) const;
    
    /**
     * @brief 🚀 获取工作表（按名称）
     * @param name 工作表名称
     * @return 工作表指针
     */
    TXSheet* getSheet(const std::string& name);
    const TXSheet* getSheet(const std::string& name) const;
    
    /**
     * @brief 🚀 获取当前活动工作表
     * @return 工作表指针
     */
    TXSheet* getActiveSheet();
    const TXSheet* getActiveSheet() const;
    
    /**
     * @brief 🚀 设置活动工作表
     * @param index 工作表索引
     * @return 操作结果
     */
    TXResult<void> setActiveSheet(size_t index);
    
    /**
     * @brief 🚀 设置活动工作表
     * @param name 工作表名称
     * @return 操作结果
     */
    TXResult<void> setActiveSheet(const std::string& name);
    
    /**
     * @brief 🚀 获取所有工作表名称
     * @return 工作表名称列表
     */
    TXVector<std::string> getSheetNames() const;
    
    /**
     * @brief 🚀 检查工作表是否存在
     * @param name 工作表名称
     * @return 是否存在
     */
    bool hasSheet(const std::string& name) const;
    
    /**
     * @brief 🚀 查找工作表索引
     * @param name 工作表名称
     * @return 索引（-1表示未找到）
     */
    int findSheetIndex(const std::string& name) const;

    // ==================== 文件操作 ====================
    
    /**
     * @brief 🚀 保存到文件
     * @param file_path 文件路径
     * @return 操作结果
     */
    TXResult<void> saveAs(const std::string& file_path);
    
    /**
     * @brief 🚀 保存到当前文件
     * @return 操作结果
     */
    TXResult<void> save();
    
    /**
     * @brief 🚀 导出到内存
     * @return 内存数据
     */
    TXResult<TXVector<uint8_t>> exportToMemory();
    
    /**
     * @brief 🚀 获取当前文件路径
     */
    const std::string& getFilePath() const { return file_path_; }
    
    /**
     * @brief 🚀 检查是否有未保存的更改
     */
    bool hasUnsavedChanges() const { return has_unsaved_changes_; }
    
    /**
     * @brief 🚀 标记为已保存
     */
    void markAsSaved() { has_unsaved_changes_ = false; }
    
    /**
     * @brief 🚀 标记为已修改
     */
    void markAsModified() { has_unsaved_changes_ = true; }

    // ==================== 便捷操作 ====================
    
    /**
     * @brief 🚀 操作符[] - 按索引访问工作表
     */
    TXSheet* operator[](size_t index) { return getSheet(index); }
    const TXSheet* operator[](size_t index) const { return getSheet(index); }
    
    /**
     * @brief 🚀 操作符[] - 按名称访问工作表
     */
    TXSheet* operator[](const std::string& name) { return getSheet(name); }
    const TXSheet* operator[](const std::string& name) const { return getSheet(name); }

    // ==================== 性能优化 ====================
    
    /**
     * @brief 🚀 预分配内存
     * @param estimated_sheets 预计工作表数量
     */
    void reserve(size_t estimated_sheets);
    
    /**
     * @brief 🚀 优化所有工作表
     */
    void optimize();
    
    /**
     * @brief 🚀 压缩所有工作表
     * @return 压缩的单元格总数
     */
    size_t compress();
    
    /**
     * @brief 🚀 收缩内存
     */
    void shrinkToFit();

    // ==================== 调试和诊断 ====================
    
    /**
     * @brief 🚀 获取调试信息
     */
    std::string toString() const;
    
    /**
     * @brief 🚀 验证工作簿状态
     */
    bool isValid() const;
    
    /**
     * @brief 🚀 获取性能统计
     */
    std::string getPerformanceStats() const;
    
    /**
     * @brief 🚀 获取内存使用情况
     */
    size_t getMemoryUsage() const;

private:
    std::string name_;                                    // 工作簿名称
    TXVector<std::unique_ptr<TXSheet>> sheets_;          // 工作表列表 (高性能TXVector)
    std::unordered_map<std::string, size_t> sheet_map_;  // 名称到索引的映射
    int active_sheet_index_;                             // 当前活动工作表索引
    std::string file_path_;                              // 文件路径
    bool has_unsaved_changes_;                           // 是否有未保存的更改
    
    TXUnifiedMemoryManager& memory_manager_;             // 内存管理器引用
    TXGlobalStringPool& string_pool_;                    // 字符串池引用
    
    // ==================== 内部辅助方法 ====================
    
    /**
     * @brief 内部错误处理
     */
    void handleError(const std::string& operation, const TXError& error) const;
    
    /**
     * @brief 验证工作表索引
     */
    bool isValidIndex(size_t index) const;
    
    /**
     * @brief 生成唯一工作表名称
     */
    std::string generateUniqueSheetName(const std::string& base_name) const;
    
    /**
     * @brief 更新工作表映射
     */
    void updateSheetMap();
    
    /**
     * @brief 调整活动工作表索引
     */
    void adjustActiveSheetIndex();
};

/**
 * @brief 🚀 便捷的工作簿创建函数
 */
inline std::unique_ptr<TXWorkbook> makeWorkbook(const std::string& name = "新建工作簿") {
    return TXWorkbook::create(name);
}

} // namespace TinaXlsx
