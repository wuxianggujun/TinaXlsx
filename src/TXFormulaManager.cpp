#include "TinaXlsx/TXFormulaManager.hpp"
#include "TinaXlsx/TXCellManager.hpp"
#include "TinaXlsx/TXCell.hpp"
#include <regex>
#include <algorithm>
#include <queue>
#include <unordered_set>
#include <cctype>
#include <stdexcept>

namespace TinaXlsx {

// ==================== 公式操作 ====================

bool TXFormulaManager::setCellFormula(const TXCoordinate& coord, const std::string& formula, TXCellManager& cellManager) {
    if (!validateFormula(formula)) {
        return false;
    }
    
    auto* cell = cellManager.getCell(coord);
    if (!cell) {
        return false;
    }
    
    cell->setFormula(formula);
    return true;
}

std::string TXFormulaManager::getCellFormula(const TXCoordinate& coord, const TXCellManager& cellManager) const {
    const auto* cell = cellManager.getCell(coord);
    if (!cell) {
        return "";
    }
    
    return cell->getFormula();
}

std::size_t TXFormulaManager::setCellFormulas(const std::vector<std::pair<TXCoordinate, std::string>>& formulas, 
                                             TXCellManager& cellManager) {
    std::size_t count = 0;
    for (const auto& pair : formulas) {
        if (setCellFormula(pair.first, pair.second, cellManager)) {
            ++count;
        }
    }
    return count;
}

bool TXFormulaManager::hasFormula(const TXCoordinate& coord, const TXCellManager& cellManager) const {
    const auto* cell = cellManager.getCell(coord);
    return cell ? cell->hasFormula() : false;
}

// ==================== 公式计算 ====================

std::size_t TXFormulaManager::calculateAllFormulas(TXCellManager& cellManager) {
    std::size_t count = 0;
    
    if (!options_.autoCalculate) {
        return count;
    }
    
    // 获取依赖关系图
    auto dependencies = getFormulaDependencies(cellManager);
    
    // 拓扑排序计算顺序
    std::vector<TXCoordinate> calculationOrder = getCalculationOrder(dependencies);
    
    // 按顺序计算公式
    for (const auto& coord : calculationOrder) {
        if (calculateFormula(coord, cellManager)) {
            ++count;
        }
    }
    
    return count;
}

std::size_t TXFormulaManager::calculateFormulasInRange(const TXRange& range, TXCellManager& cellManager) {
    if (!range.isValid()) {
        return 0;
    }
    
    std::size_t count = 0;
    auto start = range.getStart();
    auto end = range.getEnd();
    
    for (row_t row = start.getRow(); row <= end.getRow(); ++row) {
        for (column_t col = start.getCol(); col <= end.getCol(); ++col) {
            TXCoordinate coord(row, col);
            if (hasFormula(coord, cellManager)) {
                if (calculateFormula(coord, cellManager)) {
                    ++count;
                }
            }
        }
    }
    
    return count;
}

bool TXFormulaManager::calculateFormula(const TXCoordinate& coord, TXCellManager& cellManager) {
    auto* cell = cellManager.getCell(coord);
    if (!cell || !cell->hasFormula()) {
        return false;
    }
    
    std::string formula = cell->getFormula();
    
    try {
        // 计算公式结果
        cell_value_t result = evaluateFormula(formula, cellManager);
        cell->setValue(result);
        return true;
    } catch (...) {
        // 计算失败，设置错误值
        cell->setValue(std::string("#ERROR!"));
        return false;
    }
}

std::size_t TXFormulaManager::recalculateDependents(const TXCoordinate& coord, TXCellManager& cellManager) {
    std::size_t count = 0;
    
    // 获取依赖关系图
    auto dependencies = getFormulaDependencies(cellManager);
    
    // 找到所有依赖于此单元格的单元格
    std::vector<TXCoordinate> dependents = getDependents(coord, cellManager);
    
    // 递归计算所有依赖单元格
    for (const auto& dependent : dependents) {
        if (calculateFormula(dependent, cellManager)) {
            ++count;
        }
        // 递归计算依赖的依赖
        count += recalculateDependents(dependent, cellManager);
    }
    
    return count;
}

// ==================== 依赖关系分析 ====================

TXFormulaManager::DependencyGraph TXFormulaManager::getFormulaDependencies(const TXCellManager& cellManager) const {
    DependencyGraph dependencies;
    
    for (auto it = cellManager.begin(); it != cellManager.end(); ++it) {
        const auto& coord = it->first;
        const auto& cell = it->second;
        
        if (cell.hasFormula()) {
            std::vector<TXCoordinate> refs = parseFormulaReferences(cell.getFormula());
            dependencies[coord] = refs;
        }
    }
    
    return dependencies;
}

std::vector<TXCoordinate> TXFormulaManager::getDirectDependencies(const TXCoordinate& coord, 
                                                                 const TXCellManager& cellManager) const {
    const auto* cell = cellManager.getCell(coord);
    if (!cell || !cell->hasFormula()) {
        return {};
    }
    
    return parseFormulaReferences(cell->getFormula());
}

std::vector<TXCoordinate> TXFormulaManager::getDependents(const TXCoordinate& coord, 
                                                         const TXCellManager& cellManager) const {
    std::vector<TXCoordinate> dependents;
    
    // 遍历所有单元格，找到引用了指定单元格的公式
    for (auto it = cellManager.begin(); it != cellManager.end(); ++it) {
        const auto& cellCoord = it->first;
        const auto& cell = it->second;
        
        if (cell.hasFormula()) {
            std::vector<TXCoordinate> refs = parseFormulaReferences(cell.getFormula());
            if (std::find(refs.begin(), refs.end(), coord) != refs.end()) {
                dependents.push_back(cellCoord);
            }
        }
    }
    
    return dependents;
}

bool TXFormulaManager::detectCircularReferences(const TXCellManager& cellManager) const {
    std::unordered_set<TXCoordinate, CoordinateHash> visiting;
    std::unordered_set<TXCoordinate, CoordinateHash> visited;
    
    for (auto it = cellManager.begin(); it != cellManager.end(); ++it) {
        const auto& coord = it->first;
        const auto& cell = it->second;
        
        if (cell.hasFormula() && visited.find(coord) == visited.end()) {
            if (detectCircularReferencesHelper(coord, visiting, visited, cellManager)) {
                return true;
            }
        }
    }
    
    return false;
}

std::vector<std::vector<TXCoordinate>> TXFormulaManager::getCircularReferences(const TXCellManager& cellManager) const {
    std::vector<std::vector<TXCoordinate>> circularRefs;
    std::unordered_set<TXCoordinate, CoordinateHash> globalVisited;
    
    for (auto it = cellManager.begin(); it != cellManager.end(); ++it) {
        const auto& coord = it->first;
        const auto& cell = it->second;
        
        if (cell.hasFormula() && globalVisited.find(coord) == globalVisited.end()) {
            std::unordered_set<TXCoordinate, CoordinateHash> visiting;
            std::unordered_set<TXCoordinate, CoordinateHash> visited;
            std::vector<TXCoordinate> path;
            
            if (findCircularReferencePath(coord, visiting, visited, path, cellManager)) {
                circularRefs.push_back(path);
                // 将路径中的所有坐标标记为已访问
                for (const auto& pathCoord : path) {
                    globalVisited.insert(pathCoord);
                }
            }
        }
    }
    
    return circularRefs;
}

// ==================== 命名范围 ====================

bool TXFormulaManager::addNamedRange(const std::string& name, const TXRange& range, const std::string& comment) {
    if (!isValidNamedRangeName(name) || !range.isValid()) {
        return false;
    }
    
    namedRanges_[name] = range;
    return true;
}

bool TXFormulaManager::removeNamedRange(const std::string& name) {
    auto it = namedRanges_.find(name);
    if (it != namedRanges_.end()) {
        namedRanges_.erase(it);
        return true;
    }
    return false;
}

TXRange TXFormulaManager::getNamedRange(const std::string& name) const {
    auto it = namedRanges_.find(name);
    return (it != namedRanges_.end()) ? it->second : TXRange();
}

bool TXFormulaManager::hasNamedRange(const std::string& name) const {
    return namedRanges_.find(name) != namedRanges_.end();
}

bool TXFormulaManager::renameNamedRange(const std::string& oldName, const std::string& newName) {
    if (!isValidNamedRangeName(newName) || !hasNamedRange(oldName) || hasNamedRange(newName)) {
        return false;
    }
    
    TXRange range = getNamedRange(oldName);
    removeNamedRange(oldName);
    addNamedRange(newName, range);
    
    return true;
}

// ==================== 公式验证 ====================

bool TXFormulaManager::validateFormula(const std::string& formula) const {
    if (formula.empty()) {
        return false;
    }
    
    // 公式必须以等号开始
    if (formula[0] != '=') {
        return false;
    }
    
    // 简单的括号匹配检查
    int parenthesesCount = 0;
    for (char c : formula) {
        if (c == '(') {
            ++parenthesesCount;
        } else if (c == ')') {
            --parenthesesCount;
            if (parenthesesCount < 0) {
                return false; // 右括号多于左括号
            }
        }
    }
    
    return parenthesesCount == 0; // 括号必须匹配
}

std::vector<std::string> TXFormulaManager::getFormulaErrors(const std::string& formula) const {
    std::vector<std::string> errors;
    
    if (formula.empty()) {
        errors.push_back("Formula is empty");
        return errors;
    }
    
    if (formula[0] != '=') {
        errors.push_back("Formula must start with '='");
    }
    
    // 检查括号匹配
    int parenthesesCount = 0;
    for (size_t i = 0; i < formula.length(); ++i) {
        char c = formula[i];
        if (c == '(') {
            ++parenthesesCount;
        } else if (c == ')') {
            --parenthesesCount;
            if (parenthesesCount < 0) {
                errors.push_back("Unmatched closing parenthesis at position " + std::to_string(i));
            }
        }
    }
    
    if (parenthesesCount > 0) {
        errors.push_back("Unmatched opening parenthesis");
    }
    
    return errors;
}

std::vector<TXCoordinate> TXFormulaManager::parseFormulaReferences(const std::string& formula) const {
    std::vector<TXCoordinate> references;

    // 简化的公式引用解析
    // 匹配A1格式的单元格引用
    std::regex cellRefRegex(R"([A-Z]+[0-9]+)");
    std::sregex_iterator iter(formula.begin(), formula.end(), cellRefRegex);
    std::sregex_iterator end;

    for (; iter != end; ++iter) {
        std::string ref = iter->str();
        try {
            TXCoordinate coord = TXCoordinate::fromAddress(ref);
            if (coord.isValid()) {
                references.push_back(coord);
            }
        } catch (...) {
            // 忽略无效引用
        }
    }

    return references;
}

std::vector<TXRange> TXFormulaManager::parseFormulaRangeReferences(const std::string& formula) const {
    std::vector<TXRange> ranges;

    // 匹配A1:B2格式的范围引用
    std::regex rangeRefRegex(R"([A-Z]+[0-9]+:[A-Z]+[0-9]+)");
    std::sregex_iterator iter(formula.begin(), formula.end(), rangeRefRegex);
    std::sregex_iterator end;

    for (; iter != end; ++iter) {
        std::string ref = iter->str();
        try {
            // 解析范围字符串
            size_t colonPos = ref.find(':');
            if (colonPos != std::string::npos) {
                std::string startRef = ref.substr(0, colonPos);
                std::string endRef = ref.substr(colonPos + 1);

                TXCoordinate start = TXCoordinate::fromAddress(startRef);
                TXCoordinate end = TXCoordinate::fromAddress(endRef);

                if (start.isValid() && end.isValid()) {
                    ranges.emplace_back(start, end);
                }
            }
        } catch (...) {
            // 忽略无效引用
        }
    }

    return ranges;
}

TXFormulaManager::FormulaStats TXFormulaManager::getFormulaStats(const TXCellManager& cellManager) const {
    FormulaStats stats;
    stats.namedRanges = namedRanges_.size();

    for (auto it = cellManager.begin(); it != cellManager.end(); ++it) {
        const auto& cell = it->second;
        if (cell.hasFormula()) {
            ++stats.totalFormulas;

            std::string formula = cell.getFormula();
            if (validateFormula(formula)) {
                ++stats.validFormulas;
            } else {
                ++stats.invalidFormulas;
            }
        }
    }

    // 检测循环引用
    if (detectCircularReferences(cellManager)) {
        auto circularRefs = getCircularReferences(cellManager);
        stats.circularReferences = circularRefs.size();
    }

    return stats;
}

void TXFormulaManager::clear() {
    namedRanges_.clear();
    options_ = FormulaCalculationOptions::createDefault();
}

// ==================== 私有辅助方法 ====================

bool TXFormulaManager::detectCircularReferencesHelper(const TXCoordinate& coord,
                                                     std::unordered_set<TXCoordinate, CoordinateHash>& visiting,
                                                     std::unordered_set<TXCoordinate, CoordinateHash>& visited,
                                                     const TXCellManager& cellManager) const {
    if (visiting.find(coord) != visiting.end()) {
        return true; // 发现循环
    }

    if (visited.find(coord) != visited.end()) {
        return false; // 已经访问过
    }

    visiting.insert(coord);

    const auto* cell = cellManager.getCell(coord);
    if (cell && cell->hasFormula()) {
        std::vector<TXCoordinate> refs = parseFormulaReferences(cell->getFormula());
        for (const TXCoordinate& ref : refs) {
            if (detectCircularReferencesHelper(ref, visiting, visited, cellManager)) {
                return true;
            }
        }
    }

    visiting.erase(coord);
    visited.insert(coord);

    return false;
}

bool TXFormulaManager::findCircularReferencePath(const TXCoordinate& coord,
                                                std::unordered_set<TXCoordinate, CoordinateHash>& visiting,
                                                std::unordered_set<TXCoordinate, CoordinateHash>& visited,
                                                std::vector<TXCoordinate>& path,
                                                const TXCellManager& cellManager) const {
    if (visiting.find(coord) != visiting.end()) {
        // 找到循环，记录路径
        auto it = std::find(path.begin(), path.end(), coord);
        if (it != path.end()) {
            path.erase(path.begin(), it); // 只保留循环部分
        }
        return true;
    }

    if (visited.find(coord) != visited.end()) {
        return false;
    }

    visiting.insert(coord);
    path.push_back(coord);

    const auto* cell = cellManager.getCell(coord);
    if (cell && cell->hasFormula()) {
        std::vector<TXCoordinate> refs = parseFormulaReferences(cell->getFormula());
        for (const TXCoordinate& ref : refs) {
            if (findCircularReferencePath(ref, visiting, visited, path, cellManager)) {
                return true;
            }
        }
    }

    visiting.erase(coord);
    visited.insert(coord);
    path.pop_back();

    return false;
}

std::vector<TXCoordinate> TXFormulaManager::getCalculationOrder(const DependencyGraph& dependencies) const {
    std::vector<TXCoordinate> order;
    std::unordered_set<TXCoordinate, CoordinateHash> visited;
    std::unordered_set<TXCoordinate, CoordinateHash> visiting;

    // 拓扑排序
    for (const auto& pair : dependencies) {
        const TXCoordinate& coord = pair.first;
        if (visited.find(coord) == visited.end()) {
            topologicalSort(coord, dependencies, visited, visiting, order);
        }
    }

    return order;
}

void TXFormulaManager::topologicalSort(const TXCoordinate& coord,
                                      const DependencyGraph& dependencies,
                                      std::unordered_set<TXCoordinate, CoordinateHash>& visited,
                                      std::unordered_set<TXCoordinate, CoordinateHash>& visiting,
                                      std::vector<TXCoordinate>& order) const {
    if (visiting.find(coord) != visiting.end()) {
        return; // 循环引用，跳过
    }

    if (visited.find(coord) != visited.end()) {
        return; // 已访问
    }

    visiting.insert(coord);

    auto it = dependencies.find(coord);
    if (it != dependencies.end()) {
        for (const TXCoordinate& dep : it->second) {
            topologicalSort(dep, dependencies, visited, visiting, order);
        }
    }

    visiting.erase(coord);
    visited.insert(coord);
    order.push_back(coord);
}

bool TXFormulaManager::isValidNamedRangeName(const std::string& name) const {
    if (name.empty() || name.length() > 255) {
        return false;
    }

    // 名称不能以数字开头
    if (std::isdigit(name[0])) {
        return false;
    }

    // 名称只能包含字母、数字、下划线和点
    for (char c : name) {
        if (!std::isalnum(c) && c != '_' && c != '.') {
            return false;
        }
    }

    return true;
}

cell_value_t TXFormulaManager::evaluateFormula(const std::string& formula, const TXCellManager& cellManager) const {
    // 简化的公式计算实现
    // 实际实现需要完整的表达式解析器

    if (formula.length() < 2 || formula[0] != '=') {
        return std::string("#ERROR!");
    }

    std::string expr = formula.substr(1); // 去掉等号

    // 简单的数字计算
    try {
        double result = std::stod(expr);
        return result;
    } catch (...) {
        // 不是简单数字，返回错误
        return std::string("#VALUE!");
    }
}

} // namespace TinaXlsx
