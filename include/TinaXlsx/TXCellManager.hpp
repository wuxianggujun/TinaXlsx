#pragma once

#include <unordered_map>
#include <vector>
#include <memory>
#include <functional>
#include "TXCoordinate.hpp"
#include "TXCell.hpp"
#include "TXRange.hpp"
#include "TXTypes.hpp"

namespace TinaXlsx {

/**
 * @brief 单元格管理器
 * 
 * 专门负责单元格的存储、访问和基本操作
 * 职责：
 * - 单元格的创建和销毁
 * - 单元格值的读写
 * - 单元格的查找和访问
 * - 批量单元格操作
 */
class TXCellManager {
public:
    using CellValue = cell_value_t;
    using Coordinate = TXCoordinate;

    /**
     * @brief 坐标哈希函数
     */
    struct CoordinateHash {
        std::size_t operator()(const TXCoordinate& coord) const {
            return std::hash<u32>()(coord.getRow().index()) ^
                (std::hash<u32>()(coord.getCol().index()) << 1);
        }
    };

    using CellContainer = std::unordered_map<Coordinate, TXCell, CoordinateHash>;

    TXCellManager() = default;
    ~TXCellManager() = default;

    // 禁用拷贝，支持移动
    TXCellManager(const TXCellManager&) = delete;
    TXCellManager& operator=(const TXCellManager&) = delete;
    TXCellManager(TXCellManager&&) = default;
    TXCellManager& operator=(TXCellManager&&) = default;

    // ==================== 单元格访问 ====================

    /**
     * @brief 获取单元格
     * @param coord 坐标
     * @return 单元格指针，不存在则创建
     */
    TXCell* getCell(const Coordinate& coord);

    /**
     * @brief 获取单元格（只读）
     * @param coord 坐标
     * @return 单元格指针，不存在返回nullptr
     */
    const TXCell* getCell(const Coordinate& coord) const;

    /**
     * @brief 检查单元格是否存在
     * @param coord 坐标
     * @return 存在返回true
     */
    bool hasCell(const Coordinate& coord) const;

    /**
     * @brief 删除单元格
     * @param coord 坐标
     * @return 成功返回true
     */
    bool removeCell(const Coordinate& coord);

    // ==================== 值操作 ====================

    /**
     * @brief 设置单元格值
     * @param coord 坐标
     * @param value 值
     * @return 成功返回true
     */
    bool setCellValue(const Coordinate& coord, const CellValue& value);

    /**
     * @brief 获取单元格值
     * @param coord 坐标
     * @return 单元格值
     */
    CellValue getCellValue(const Coordinate& coord) const;

    /**
     * @brief 批量设置单元格值
     * @param values 坐标-值对列表
     * @return 成功设置的数量
     */
    std::size_t setCellValues(const std::vector<std::pair<Coordinate, CellValue>>& values);

    /**
     * @brief 批量获取单元格值
     * @param coords 坐标列表
     * @return 坐标-值对列表
     */
    std::vector<std::pair<Coordinate, CellValue>> getCellValues(const std::vector<Coordinate>& coords) const;

    // ==================== 范围操作 ====================

    /**
     * @brief 获取使用的范围
     * @return 包含所有非空单元格的范围
     */
    TXRange getUsedRange() const;

    /**
     * @brief 获取最大使用行
     * @return 最大行号
     */
    row_t getMaxUsedRow() const;

    /**
     * @brief 获取最大使用列
     * @return 最大列号
     */
    column_t getMaxUsedColumn() const;

    /**
     * @brief 清空所有单元格
     */
    void clear();

    /**
     * @brief 获取单元格数量
     * @return 单元格总数
     */
    std::size_t getCellCount() const { return cells_.size(); }

    /**
     * @brief 获取非空单元格数量
     * @return 非空单元格数量
     */
    std::size_t getNonEmptyCellCount() const;

    // ==================== 迭代器支持 ====================

    /**
     * @brief 获取所有单元格的迭代器
     */
    CellContainer::iterator begin() { return cells_.begin(); }
    CellContainer::iterator end() { return cells_.end(); }
    CellContainer::const_iterator begin() const { return cells_.begin(); }
    CellContainer::const_iterator end() const { return cells_.end(); }
    CellContainer::const_iterator cbegin() const { return cells_.cbegin(); }
    CellContainer::const_iterator cend() const { return cells_.cend(); }

    // ==================== 行列移动支持 ====================

    /**
     * @brief 移动单元格（用于行列插入删除）
     * @param transform 坐标变换函数
     */
    void transformCells(std::function<Coordinate(const Coordinate&)> transform);

    /**
     * @brief 删除指定范围内的单元格
     * @param range 要删除的范围
     * @return 删除的单元格数量
     */
    std::size_t removeCellsInRange(const TXRange& range);

private:
    CellContainer cells_;

    /**
     * @brief 验证坐标有效性
     * @param coord 坐标
     * @return 有效返回true
     */
    bool isValidCoordinate(const Coordinate& coord) const;
};

} // namespace TinaXlsx
