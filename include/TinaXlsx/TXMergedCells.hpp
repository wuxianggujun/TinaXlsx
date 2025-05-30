#pragma once

#include "TXTypes.hpp"
#include "TXCoordinate.hpp"
#include "TXRange.hpp"
#include <string>
#include <vector>
#include <set>
#include <memory>
#include <unordered_map>
#include <regex>

namespace TinaXlsx {

/**
 * @brief 单元格合并管理类
 * 
 * 提供高性能的单元格合并、拆分和查询功能
 */
class TXMergedCells {
public:
    /**
     * @brief 合并区域结构
     */
    struct MergeRegion {
        row_t startRow;
        column_t startCol;
        row_t endRow;
        column_t endCol;
        
        MergeRegion() : startRow(1), startCol(1), endRow(1), endCol(1) {}
        
        MergeRegion(row_t sr, column_t sc, 
                   row_t er, column_t ec)
            : startRow(sr), startCol(sc), endRow(er), endCol(ec) {}
        
        MergeRegion(const TXRange& range);
        
        /**
         * @brief 检查指定单元格是否在合并区域内
         * @param row 行号
         * @param col 列号
         * @return 在区域内返回true，否则返回false
         */
        bool contains(row_t row, column_t col) const;
        
        /**
         * @brief 获取合并区域的大小
         * @return {行数, 列数}
         */
        std::pair<row_t, column_t> getSize() const;
        
        /**
         * @brief 获取合并的单元格数量
         */
        std::size_t getCellCount() const;
        
        /**
         * @brief 转换为A1格式字符串
         */
        std::string toString() const;
        
        /**
         * @brief 从A1格式字符串解析
         */
        static MergeRegion fromString(const std::string& rangeStr);
        
        /**
         * @brief 转换为TXRange
         */
        TXRange toRange() const;
        
        /**
         * @brief 检查是否有效
         */
        bool isValid() const;
        
        /**
         * @brief 比较操作符
         */
        bool operator==(const MergeRegion& other) const;
        bool operator!=(const MergeRegion& other) const;
        bool operator<(const MergeRegion& other) const;
    };

public:
    TXMergedCells();
    ~TXMergedCells() = default;
    
    // 禁用拷贝，支持移动
    TXMergedCells(const TXMergedCells&) = delete;
    TXMergedCells& operator=(const TXMergedCells&) = delete;
    TXMergedCells(TXMergedCells&&) noexcept = default;
    TXMergedCells& operator=(TXMergedCells&&) noexcept = default;

    // ==================== 合并操作 ====================

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
    bool mergeCells(const TXRange& range);

    /**
     * @brief 合并单元格区域（使用A1格式）
     * @param rangeStr 范围字符串，如"A1:C3"
     * @return 成功返回true，失败返回false
     */
    bool mergeCells(const std::string& rangeStr);

    // ==================== 拆分操作 ====================

    /**
     * @brief 拆分包含指定单元格的合并区域
     * @param row 单元格行号
     * @param col 单元格列号
     * @return 成功返回true，失败返回false
     */
    bool unmergeCells(row_t row, column_t col);

    /**
     * @brief 拆分指定的合并区域
     * @param region 要拆分的合并区域
     * @return 成功返回true，失败返回false
     */
    bool unmergeCells(const MergeRegion& region);

    /**
     * @brief 拆分指定范围的所有合并区域
     * @param range 范围
     * @return 拆分的区域数量
     */
    std::size_t unmergeCellsInRange(const TXRange& range);

    /**
     * @brief 拆分所有合并区域
     */
    void unmergeAllCells();

    // ==================== 查询操作 ====================

    /**
     * @brief 检查单元格是否被合并
     * @param row 行号
     * @param col 列号
     * @return 被合并返回true，否则返回false
     */
    bool isMerged(row_t row, column_t col) const;

    /**
     * @brief 获取包含指定单元格的合并区域
     * @param row 行号
     * @param col 列号
     * @return 合并区域指针，如果不存在返回nullptr
     */
    const MergeRegion* getMergeRegion(row_t row, column_t col) const;

    /**
     * @brief 获取所有合并区域
     * @return 合并区域列表
     */
    std::vector<MergeRegion> getAllMergeRegions() const;

    /**
     * @brief 获取合并区域数量
     * @return 合并区域数量
     */
    std::size_t getMergeCount() const;

    /**
     * @brief 检查区域是否可以合并
     * @param startRow 起始行
     * @param startCol 起始列
     * @param endRow 结束行
     * @param endCol 结束列
     * @return 可以合并返回true，否则返回false
     */
    bool canMerge(row_t startRow, column_t startCol,
                  row_t endRow, column_t endCol) const;

    /**
     * @brief 获取与指定范围重叠的合并区域
     * @param range 范围
     * @return 重叠的合并区域列表
     */
    std::vector<MergeRegion> getOverlappingRegions(const TXRange& range) const;

    // ==================== 范围操作 ====================

    /**
     * @brief 获取指定范围内的所有合并区域
     * @param range 范围
     * @return 合并区域列表
     */
    std::vector<MergeRegion> getMergeRegionsInRange(const TXRange& range) const;

    /**
     * @brief 检查范围是否包含合并区域
     * @param range 范围
     * @return 包含返回true，否则返回false
     */
    bool hasmergeInRange(const TXRange& range) const;

    // ==================== 批量操作 ====================

    /**
     * @brief 批量合并单元格（高性能版本）
     * @param regions 要合并的区域列表
     * @return 成功合并的区域数量
     */
    std::size_t batchMergeCells(const std::vector<MergeRegion>& regions);

    /**
     * @brief 批量拆分单元格（高性能版本）
     * @param regions 要拆分的区域列表
     * @return 成功拆分的区域数量
     */
    std::size_t batchUnmergeCells(const std::vector<MergeRegion>& regions);

    // ==================== 工具函数 ====================

    /**
     * @brief 清空所有合并信息
     */
    void clear();

    /**
     * @brief 检查是否为空
     * @return 为空返回true，否则返回false
     */
    bool empty() const;

    /**
     * @brief 获取最后的错误信息
     * @return 错误信息字符串
     */
    const std::string& getLastError() const;

    // ==================== 静态工具函数 ====================

    /**
     * @brief 检查两个区域是否重叠
     * @param region1 区域1
     * @param region2 区域2
     * @return 重叠返回true，否则返回false
     */
    static bool isOverlapping(const MergeRegion& region1, const MergeRegion& region2);

    /**
     * @brief 检查区域是否有效
     * @param region 区域
     * @return 有效返回true，否则返回false
     */
    static bool isValidRegion(const MergeRegion& region);

    /**
     * @brief 规范化区域（确保起始位置在结束位置之前）
     * @param region 区域
     * @return 规范化后的区域
     */
    static MergeRegion normalizeRegion(const MergeRegion& region);

private:
    // 使用集合存储合并区域，保证有序和唯一性
    std::set<MergeRegion> mergeRegions_;
    
    // 用于快速查找包含特定单元格的合并区域
    std::unordered_map<uint64_t, const MergeRegion*> cellToRegionMap_;
    
    std::string lastError_;

    // ==================== 私有辅助方法 ====================

    /**
     * @brief 生成单元格的唯一键
     */
    uint64_t getCellKey(row_t row, column_t col) const;
    
    /**
     * @brief 更新单元格到区域的映射
     */
    void updateCellMapping(const MergeRegion& region, const MergeRegion* regionPtr);
    
    /**
     * @brief 内部合并检查（不进行验证）
     */
    bool canMergeInternal(const MergeRegion& region) const;
};

} // namespace TinaXlsx 