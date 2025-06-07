//
// @file TXSheet.hpp
// @brief 🚀 用户层工作表类 - 高级工作表操作接口
//

#pragma once

#include <TinaXlsx/user/TXCell.hpp>
#include <TinaXlsx/TXInMemorySheet.hpp>
#include <TinaXlsx/TXRange.hpp>
#include <TinaXlsx/TXVariant.hpp>
#include <TinaXlsx/TXResult.hpp>
#include <TinaXlsx/TXError.hpp>
#include <TinaXlsx/TXVector.hpp>
#include <string>
#include <vector>
#include <memory>

namespace TinaXlsx {

// 前向声明
class TXWorkbook;

/**
 * @brief 🚀 用户层工作表类
 * 
 * 设计理念：
 * - 封装TXInMemorySheet，提供用户友好的高级API
 * - 支持单元格、范围、批量操作
 * - 完整的错误处理和类型安全
 * - 高性能优化和内存管理
 * 
 * 使用示例：
 * ```cpp
 * TXSheet sheet("我的工作表");
 * 
 * // 单元格操作
 * sheet.cell("A1").setValue(42.0);
 * sheet.cell(0, 1).setValue("Hello");
 * 
 * // 范围操作
 * auto range = sheet.range("A1:C3");
 * range.fill(100.0);
 * 
 * // 批量操作
 * std::vector<std::vector<TXVariant>> data = {
 *     {1.0, 2.0, 3.0},
 *     {"A", "B", "C"}
 * };
 * sheet.setValues("A1:C2", data);
 * ```
 */
class TXSheet {
public:
    // ==================== 构造和析构 ====================
    
    /**
     * @brief 🚀 构造函数
     * @param name 工作表名称
     * @param memory_manager 内存管理器引用
     * @param string_pool 字符串池引用
     */
    explicit TXSheet(
        const std::string& name,
        TXUnifiedMemoryManager& memory_manager,
        TXGlobalStringPool& string_pool
    );
    
    /**
     * @brief 🚀 从现有TXInMemorySheet构造 (用于TXWorkbook)
     */
    explicit TXSheet(std::unique_ptr<TXInMemorySheet> sheet);
    
    /**
     * @brief 🚀 析构函数
     */
    ~TXSheet();
    
    // 禁用拷贝，支持移动
    TXSheet(const TXSheet&) = delete;
    TXSheet& operator=(const TXSheet&) = delete;
    TXSheet(TXSheet&&) noexcept;
    TXSheet& operator=(TXSheet&&) noexcept;

    // ==================== 单元格访问 ====================
    
    /**
     * @brief 🚀 获取单元格 (Excel格式)
     * @param address Excel格式地址 (如 "A1", "B2")
     * @return TXCell对象
     */
    TXCell cell(const std::string& address);
    
    /**
     * @brief 🚀 获取单元格 (行列索引)
     * @param row 行索引 (0-based)
     * @param col 列索引 (0-based)
     * @return TXCell对象
     */
    TXCell cell(uint32_t row, uint32_t col);
    
    /**
     * @brief 🚀 获取单元格 (坐标对象)
     * @param coord 坐标对象
     * @return TXCell对象
     */
    TXCell cell(const TXCoordinate& coord);

    // ==================== 范围操作 ====================
    
    /**
     * @brief 🚀 获取范围 (Excel格式)
     * @param range_address Excel格式范围 (如 "A1:C3")
     * @return TXRange对象
     */
    TXRange range(const std::string& range_address);
    
    /**
     * @brief 🚀 获取范围 (坐标)
     * @param start_row 起始行 (0-based)
     * @param start_col 起始列 (0-based)
     * @param end_row 结束行 (0-based)
     * @param end_col 结束列 (0-based)
     * @return TXRange对象
     */
    TXRange range(uint32_t start_row, uint32_t start_col, 
                  uint32_t end_row, uint32_t end_col);
    
    /**
     * @brief 🚀 获取范围 (坐标对象)
     * @param start 起始坐标
     * @param end 结束坐标
     * @return TXRange对象
     */
    TXRange range(const TXCoordinate& start, const TXCoordinate& end);

    // ==================== 批量数据操作 ====================
    
    /**
     * @brief 🚀 批量设置值
     * @param range_address 范围地址
     * @param values 二维数据数组
     * @return 操作结果
     */
    TXResult<void> setValues(const std::string& range_address,
                            const TXVector<TXVector<TXVariant>>& values);
    
    /**
     * @brief 🚀 批量获取值
     * @param range_address 范围地址
     * @return 二维数据数组
     */
    TXResult<TXVector<TXVector<TXVariant>>> getValues(const std::string& range_address);
    
    /**
     * @brief 🚀 填充范围
     * @param range_address 范围地址
     * @param value 填充值
     * @return 操作结果
     */
    TXResult<void> fillRange(const std::string& range_address, const TXVariant& value);
    
    /**
     * @brief 🚀 清除范围
     * @param range_address 范围地址
     * @return 操作结果
     */
    TXResult<void> clearRange(const std::string& range_address);

    // ==================== 工作表属性 ====================
    
    /**
     * @brief 🚀 获取工作表名称
     */
    const std::string& getName() const;
    
    /**
     * @brief 🚀 设置工作表名称
     */
    void setName(const std::string& name);
    
    /**
     * @brief 🚀 获取使用范围
     */
    TXRange getUsedRange() const;
    
    /**
     * @brief 🚀 获取单元格数量
     */
    size_t getCellCount() const;
    
    /**
     * @brief 🚀 检查是否为空
     */
    bool isEmpty() const;

    // ==================== 性能优化 ====================
    
    /**
     * @brief 🚀 预分配内存
     * @param estimated_cells 预计单元格数量
     */
    void reserve(size_t estimated_cells);
    
    /**
     * @brief 🚀 优化内存布局
     */
    void optimize();
    
    /**
     * @brief 🚀 压缩稀疏数据
     * @return 压缩的单元格数量
     */
    size_t compress();
    
    /**
     * @brief 🚀 收缩内存到实际使用大小
     */
    void shrinkToFit();

    // ==================== 查找和统计 ====================
    
    /**
     * @brief 🚀 查找值
     * @param value 要查找的值
     * @param range_address 查找范围 (可选)
     * @return 找到的坐标列表
     */
    TXVector<TXCoordinate> findValue(const TXVariant& value,
                                    const std::string& range_address = "");
    
    /**
     * @brief 🚀 统计范围
     * @param range_address 统计范围
     * @return 统计结果
     */
    TXResult<double> sum(const std::string& range_address);
    
    /**
     * @brief 🚀 计算平均值
     * @param range_address 计算范围
     * @return 平均值
     */
    TXResult<double> average(const std::string& range_address);
    
    /**
     * @brief 🚀 获取最大值
     * @param range_address 查找范围
     * @return 最大值
     */
    TXResult<double> max(const std::string& range_address);
    
    /**
     * @brief 🚀 获取最小值
     * @param range_address 查找范围
     * @return 最小值
     */
    TXResult<double> min(const std::string& range_address);

    // ==================== 调试和诊断 ====================
    
    /**
     * @brief 🚀 获取调试信息
     */
    std::string toString() const;
    
    /**
     * @brief 🚀 验证工作表状态
     */
    bool isValid() const;
    
    /**
     * @brief 🚀 获取性能统计
     */
    std::string getPerformanceStats() const;

    // ==================== 内部访问 (友元类用) ====================
    
    /**
     * @brief 🚀 获取底层TXInMemorySheet (仅供TXWorkbook使用)
     */
    TXInMemorySheet& getInternalSheet() { return *sheet_; }
    const TXInMemorySheet& getInternalSheet() const { return *sheet_; }

private:
    std::unique_ptr<TXInMemorySheet> sheet_;  // 底层工作表
    
    // ==================== 内部辅助方法 ====================
    
    /**
     * @brief 内部错误处理
     */
    void handleError(const std::string& operation, const TXError& error) const;
    
    /**
     * @brief 解析范围地址
     */
    TXResult<TXRange> parseRangeAddress(const std::string& range_address);
    
    friend class TXWorkbook;  // 允许TXWorkbook访问内部方法
};

/**
 * @brief 🚀 便捷的工作表创建函数
 */
inline std::unique_ptr<TXSheet> makeSheet(const std::string& name) {
    return std::make_unique<TXSheet>(
        name,
        GlobalUnifiedMemoryManager::getInstance(),
        TXGlobalStringPool::instance()
    );
}

} // namespace TinaXlsx
