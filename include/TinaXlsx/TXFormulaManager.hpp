#pragma once

#include <string>
#include <vector>
#include <unordered_map>
#include <unordered_set>
#include <functional>
#include "TXCoordinate.hpp"
#include "TXRange.hpp"
#include "TXTypes.hpp"

namespace TinaXlsx {

// 前向声明
class TXCellManager;

/**
 * @brief 公式管理器
 * 
 * 专门负责公式的管理和计算
 * 职责：
 * - 公式的设置和获取
 * - 公式依赖关系分析
 * - 公式计算和更新
 * - 循环引用检测
 * - 命名范围管理
 */
class TXFormulaManager {
public:
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
     * @brief 公式计算选项
     */
    struct FormulaCalculationOptions {
        bool autoCalculate = true;                   ///< 自动计算
        bool iterativeCalculation = false;           ///< 迭代计算
        int maxIterations = 100;                     ///< 最大迭代次数
        double maxChange = 0.001;                    ///< 最大变化值
        bool precisionAsDisplayed = false;           ///< 以显示精度计算
        bool use1904DateSystem = false;              ///< 使用1904日期系统

        /**
         * @brief 创建默认计算选项
         */
        static FormulaCalculationOptions createDefault() {
            return FormulaCalculationOptions{};
        }

        /**
         * @brief 创建高精度计算选项
         */
        static FormulaCalculationOptions createHighPrecision() {
            FormulaCalculationOptions options;
            options.maxIterations = 1000;
            options.maxChange = 0.0001;
            options.precisionAsDisplayed = false;
            return options;
        }
    };

    /**
     * @brief 公式依赖关系图
     */
    using DependencyGraph = std::unordered_map<TXCoordinate, std::vector<TXCoordinate>, CoordinateHash>;

    TXFormulaManager() = default;
    ~TXFormulaManager() = default;

    // 禁用拷贝，支持移动
    TXFormulaManager(const TXFormulaManager&) = delete;
    TXFormulaManager& operator=(const TXFormulaManager&) = delete;
    TXFormulaManager(TXFormulaManager&&) = default;
    TXFormulaManager& operator=(TXFormulaManager&&) = default;

    // ==================== 计算选项 ====================

    /**
     * @brief 设置公式计算选项
     * @param options 计算选项
     */
    void setCalculationOptions(const FormulaCalculationOptions& options) { options_ = options; }

    /**
     * @brief 获取公式计算选项
     * @return 计算选项
     */
    const FormulaCalculationOptions& getCalculationOptions() const { return options_; }

    // ==================== 公式操作 ====================

    /**
     * @brief 设置单元格公式
     * @param coord 坐标
     * @param formula 公式字符串
     * @param cellManager 单元格管理器
     * @return 成功返回true
     */
    bool setCellFormula(const TXCoordinate& coord, const std::string& formula, TXCellManager& cellManager);

    /**
     * @brief 获取单元格公式
     * @param coord 坐标
     * @param cellManager 单元格管理器
     * @return 公式字符串，如果不是公式返回空字符串
     */
    std::string getCellFormula(const TXCoordinate& coord, const TXCellManager& cellManager) const;

    /**
     * @brief 批量设置公式
     * @param formulas 坐标-公式对列表
     * @param cellManager 单元格管理器
     * @return 成功设置的公式数量
     */
    std::size_t setCellFormulas(const std::vector<std::pair<TXCoordinate, std::string>>& formulas, 
                               TXCellManager& cellManager);

    /**
     * @brief 检查单元格是否包含公式
     * @param coord 坐标
     * @param cellManager 单元格管理器
     * @return 包含公式返回true
     */
    bool hasFormula(const TXCoordinate& coord, const TXCellManager& cellManager) const;

    // ==================== 公式计算 ====================

    /**
     * @brief 计算所有公式
     * @param cellManager 单元格管理器
     * @return 成功计算的公式数量
     */
    std::size_t calculateAllFormulas(TXCellManager& cellManager);

    /**
     * @brief 计算指定范围内的公式
     * @param range 范围
     * @param cellManager 单元格管理器
     * @return 成功计算的公式数量
     */
    std::size_t calculateFormulasInRange(const TXRange& range, TXCellManager& cellManager);

    /**
     * @brief 计算单个公式
     * @param coord 坐标
     * @param cellManager 单元格管理器
     * @return 成功返回true
     */
    bool calculateFormula(const TXCoordinate& coord, TXCellManager& cellManager);

    /**
     * @brief 重新计算依赖于指定单元格的所有公式
     * @param coord 被依赖的单元格坐标
     * @param cellManager 单元格管理器
     * @return 重新计算的公式数量
     */
    std::size_t recalculateDependents(const TXCoordinate& coord, TXCellManager& cellManager);

    // ==================== 依赖关系分析 ====================

    /**
     * @brief 获取公式依赖关系图
     * @param cellManager 单元格管理器
     * @return 依赖关系图
     */
    DependencyGraph getFormulaDependencies(const TXCellManager& cellManager) const;

    /**
     * @brief 获取单元格的直接依赖
     * @param coord 坐标
     * @param cellManager 单元格管理器
     * @return 依赖的单元格列表
     */
    std::vector<TXCoordinate> getDirectDependencies(const TXCoordinate& coord, 
                                                   const TXCellManager& cellManager) const;

    /**
     * @brief 获取依赖于指定单元格的所有单元格
     * @param coord 坐标
     * @param cellManager 单元格管理器
     * @return 依赖单元格列表
     */
    std::vector<TXCoordinate> getDependents(const TXCoordinate& coord, 
                                           const TXCellManager& cellManager) const;

    /**
     * @brief 检测循环引用
     * @param cellManager 单元格管理器
     * @return 发现循环引用返回true
     */
    bool detectCircularReferences(const TXCellManager& cellManager) const;

    /**
     * @brief 获取循环引用的单元格
     * @param cellManager 单元格管理器
     * @return 循环引用的单元格列表
     */
    std::vector<std::vector<TXCoordinate>> getCircularReferences(const TXCellManager& cellManager) const;

    // ==================== 命名范围 ====================

    /**
     * @brief 添加命名范围
     * @param name 名称
     * @param range 范围
     * @param comment 注释
     * @return 成功返回true
     */
    bool addNamedRange(const std::string& name, const TXRange& range, const std::string& comment = "");

    /**
     * @brief 删除命名范围
     * @param name 名称
     * @return 成功返回true
     */
    bool removeNamedRange(const std::string& name);

    /**
     * @brief 获取命名范围
     * @param name 名称
     * @return 范围，如果不存在返回无效范围
     */
    TXRange getNamedRange(const std::string& name) const;

    /**
     * @brief 获取所有命名范围
     * @return 名称到范围的映射
     */
    std::unordered_map<std::string, TXRange> getAllNamedRanges() const { return namedRanges_; }

    /**
     * @brief 检查命名范围是否存在
     * @param name 名称
     * @return 存在返回true
     */
    bool hasNamedRange(const std::string& name) const;

    /**
     * @brief 重命名命名范围
     * @param oldName 旧名称
     * @param newName 新名称
     * @return 成功返回true
     */
    bool renameNamedRange(const std::string& oldName, const std::string& newName);

    // ==================== 公式验证 ====================

    /**
     * @brief 验证公式语法
     * @param formula 公式字符串
     * @return 有效返回true
     */
    bool validateFormula(const std::string& formula) const;

    /**
     * @brief 获取公式中的错误信息
     * @param formula 公式字符串
     * @return 错误信息列表
     */
    std::vector<std::string> getFormulaErrors(const std::string& formula) const;

    /**
     * @brief 解析公式中的单元格引用
     * @param formula 公式字符串
     * @return 引用的单元格坐标列表
     */
    std::vector<TXCoordinate> parseFormulaReferences(const std::string& formula) const;

    /**
     * @brief 解析公式中的范围引用
     * @param formula 公式字符串
     * @return 引用的范围列表
     */
    std::vector<TXRange> parseFormulaRangeReferences(const std::string& formula) const;

    // ==================== 工具方法 ====================

    /**
     * @brief 清空所有公式和命名范围
     */
    void clear();

    /**
     * @brief 获取公式统计信息
     */
    struct FormulaStats {
        std::size_t totalFormulas = 0;
        std::size_t validFormulas = 0;
        std::size_t invalidFormulas = 0;
        std::size_t circularReferences = 0;
        std::size_t namedRanges = 0;
    };

    /**
     * @brief 获取公式统计信息
     * @param cellManager 单元格管理器
     * @return 统计信息
     */
    FormulaStats getFormulaStats(const TXCellManager& cellManager) const;

private:
    FormulaCalculationOptions options_;
    std::unordered_map<std::string, TXRange> namedRanges_;

    /**
     * @brief 循环引用检测辅助方法
     * @param coord 当前检查的坐标
     * @param visiting 正在访问的坐标集合
     * @param visited 已访问的坐标集合
     * @param cellManager 单元格管理器
     * @return 发现循环引用返回true
     */
    bool detectCircularReferencesHelper(const TXCoordinate& coord,
                                       std::unordered_set<TXCoordinate, CoordinateHash>& visiting,
                                       std::unordered_set<TXCoordinate, CoordinateHash>& visited,
                                       const TXCellManager& cellManager) const;

    /**
     * @brief 验证命名范围名称
     * @param name 名称
     * @return 有效返回true
     */
    bool isValidNamedRangeName(const std::string& name) const;

    /**
     * @brief 查找循环引用路径
     */
    bool findCircularReferencePath(const TXCoordinate& coord,
                                  std::unordered_set<TXCoordinate, CoordinateHash>& visiting,
                                  std::unordered_set<TXCoordinate, CoordinateHash>& visited,
                                  std::vector<TXCoordinate>& path,
                                  const TXCellManager& cellManager) const;

    /**
     * @brief 获取计算顺序
     */
    std::vector<TXCoordinate> getCalculationOrder(const DependencyGraph& dependencies) const;

    /**
     * @brief 拓扑排序
     */
    void topologicalSort(const TXCoordinate& coord,
                        const DependencyGraph& dependencies,
                        std::unordered_set<TXCoordinate, CoordinateHash>& visited,
                        std::unordered_set<TXCoordinate, CoordinateHash>& visiting,
                        std::vector<TXCoordinate>& order) const;

    /**
     * @brief 计算公式的实际实现
     * @param formula 公式字符串
     * @param cellManager 单元格管理器
     * @return 计算结果
     */
    cell_value_t evaluateFormula(const std::string& formula, const TXCellManager& cellManager) const;
};

} // namespace TinaXlsx
