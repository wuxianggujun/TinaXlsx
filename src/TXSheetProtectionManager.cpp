#include "TinaXlsx/TXSheetProtectionManager.hpp"
#include "TinaXlsx/TXCellManager.hpp"
#include <functional>
#include <algorithm>

namespace TinaXlsx {

// ==================== 工作表保护 ====================

bool TXSheetProtectionManager::protectSheet(const std::string& password, const SheetProtection& protection) {
    protection_ = protection;
    protection_.isProtected = true;
    
    if (!password.empty()) {
        protection_.passwordHash = generatePasswordHash(password);
    }
    
    return true;
}

bool TXSheetProtectionManager::unprotectSheet(const std::string& password) {
    if (!protection_.isProtected) {
        return true; // 已经未保护
    }
    
    if (!protection_.passwordHash.empty()) {
        if (!verifyPassword(password)) {
            return false; // 密码错误
        }
    }
    
    protection_.isProtected = false;
    protection_.passwordHash.clear();
    
    return true;
}

bool TXSheetProtectionManager::verifyPassword(const std::string& password) const {
    if (protection_.passwordHash.empty()) {
        return true; // 没有密码保护
    }
    
    return generatePasswordHash(password) == protection_.passwordHash;
}

// ==================== 单元格锁定 ====================

bool TXSheetProtectionManager::setCellLocked(const TXCoordinate& coord, bool locked, TXCellManager& cellManager) {
    auto* cell = cellManager.getCell(coord);
    if (!cell) {
        // 创建新单元格
        cellManager.setCellValue(coord, std::string(""));
        cell = cellManager.getCell(coord);
    }
    
    if (cell) {
        cell->setLocked(locked);
        return true;
    }
    
    return false;
}

bool TXSheetProtectionManager::isCellLocked(const TXCoordinate& coord, const TXCellManager& cellManager) const {
    const auto* cell = cellManager.getCell(coord);
    return cell ? cell->isLocked() : true; // 默认锁定
}

std::size_t TXSheetProtectionManager::setRangeLocked(const TXRange& range, bool locked, TXCellManager& cellManager) {
    if (!range.isValid()) {
        return 0;
    }
    
    std::size_t count = 0;
    auto start = range.getStart();
    auto end = range.getEnd();
    
    for (row_t r = start.getRow(); r <= end.getRow(); ++r) {
        for (column_t c = start.getCol(); c <= end.getCol(); ++c) {
            if (setCellLocked(TXCoordinate(r, c), locked, cellManager)) {
                ++count;
            }
        }
    }
    
    return count;
}

std::size_t TXSheetProtectionManager::setCellsLocked(const std::vector<TXCoordinate>& coords, bool locked, TXCellManager& cellManager) {
    std::size_t count = 0;
    for (const auto& coord : coords) {
        if (setCellLocked(coord, locked, cellManager)) {
            ++count;
        }
    }
    return count;
}

// ==================== 权限验证 ====================

bool TXSheetProtectionManager::isOperationAllowed(OperationType operation) const {
    if (!protection_.isProtected) {
        return true; // 未保护时允许所有操作
    }
    
    switch (operation) {
        case OperationType::SelectLockedCells:
            return protection_.selectLockedCells;
        case OperationType::SelectUnlockedCells:
            return protection_.selectUnlockedCells;
        case OperationType::FormatCells:
            return protection_.formatCells;
        case OperationType::FormatColumns:
            return protection_.formatColumns;
        case OperationType::FormatRows:
            return protection_.formatRows;
        case OperationType::InsertColumns:
            return protection_.insertColumns;
        case OperationType::InsertRows:
            return protection_.insertRows;
        case OperationType::DeleteColumns:
            return protection_.deleteColumns;
        case OperationType::DeleteRows:
            return protection_.deleteRows;
        case OperationType::InsertHyperlinks:
            return protection_.insertHyperlinks;
        case OperationType::Sort:
            return protection_.sort;
        case OperationType::AutoFilter:
            return protection_.autoFilter;
        case OperationType::PivotTables:
            return protection_.pivotTables;
        case OperationType::Objects:
            return protection_.objects;
        case OperationType::Scenarios:
            return protection_.scenarios;
        default:
            return false;
    }
}

bool TXSheetProtectionManager::isOperationAllowed(const std::string& operationName) const {
    OperationType operation = stringToOperationType(operationName);
    return isOperationAllowed(operation);
}

bool TXSheetProtectionManager::isCellEditable(const TXCoordinate& coord, const TXCellManager& cellManager) const {
    if (!protection_.isProtected) {
        return true; // 未保护时可以编辑
    }
    
    // 检查单元格是否锁定
    bool locked = isCellLocked(coord, cellManager);
    
    // 如果单元格锁定，则不可编辑
    // 如果单元格未锁定，则可编辑
    return !locked;
}

bool TXSheetProtectionManager::isRangeEditable(const TXRange& range, const TXCellManager& cellManager) const {
    if (!protection_.isProtected) {
        return true; // 未保护时可以编辑
    }
    
    auto start = range.getStart();
    auto end = range.getEnd();
    
    // 检查范围内所有单元格是否都可编辑
    for (row_t r = start.getRow(); r <= end.getRow(); ++r) {
        for (column_t c = start.getCol(); c <= end.getCol(); ++c) {
            if (!isCellEditable(TXCoordinate(r, c), cellManager)) {
                return false;
            }
        }
    }
    
    return true;
}

// ==================== 保护状态查询 ====================

std::vector<TXCoordinate> TXSheetProtectionManager::getLockedCells(const TXCellManager& cellManager) const {
    std::vector<TXCoordinate> lockedCells;
    
    for (auto it = cellManager.begin(); it != cellManager.end(); ++it) {
        const auto& coord = it->first;
        if (isCellLocked(coord, cellManager)) {
            lockedCells.push_back(coord);
        }
    }
    
    return lockedCells;
}

std::vector<TXCoordinate> TXSheetProtectionManager::getUnlockedCells(const TXCellManager& cellManager) const {
    std::vector<TXCoordinate> unlockedCells;
    
    for (auto it = cellManager.begin(); it != cellManager.end(); ++it) {
        const auto& coord = it->first;
        if (!isCellLocked(coord, cellManager)) {
            unlockedCells.push_back(coord);
        }
    }
    
    return unlockedCells;
}

TXSheetProtectionManager::ProtectionStats TXSheetProtectionManager::getProtectionStats(const TXCellManager& cellManager) const {
    ProtectionStats stats;
    stats.isProtected = protection_.isProtected;
    stats.hasPassword = !protection_.passwordHash.empty();
    
    // 统计锁定和未锁定的单元格数量
    for (auto it = cellManager.begin(); it != cellManager.end(); ++it) {
        if (isCellLocked(it->first, cellManager)) {
            ++stats.lockedCellCount;
        } else {
            ++stats.unlockedCellCount;
        }
    }
    
    // 统计允许的操作数量
    std::vector<OperationType> allOperations = {
        OperationType::SelectLockedCells, OperationType::SelectUnlockedCells,
        OperationType::FormatCells, OperationType::FormatColumns, OperationType::FormatRows,
        OperationType::InsertColumns, OperationType::InsertRows,
        OperationType::DeleteColumns, OperationType::DeleteRows,
        OperationType::InsertHyperlinks, OperationType::Sort, OperationType::AutoFilter,
        OperationType::PivotTables, OperationType::Objects, OperationType::Scenarios
    };
    
    for (OperationType op : allOperations) {
        if (isOperationAllowed(op)) {
            ++stats.allowedOperationCount;
        }
    }
    
    return stats;
}

// ==================== 工具方法 ====================

void TXSheetProtectionManager::clear() {
    protection_ = SheetProtection{};
}

void TXSheetProtectionManager::reset() {
    protection_ = SheetProtection{};
}

// ==================== 私有辅助方法 ====================

std::string TXSheetProtectionManager::generatePasswordHash(const std::string& password) const {
    // 简化的哈希实现（实际应使用MD5或更安全的算法）
    std::hash<std::string> hasher;
    return std::to_string(hasher(password));
}

std::string TXSheetProtectionManager::operationTypeToString(OperationType operation) const {
    switch (operation) {
        case OperationType::SelectLockedCells: return "selectLockedCells";
        case OperationType::SelectUnlockedCells: return "selectUnlockedCells";
        case OperationType::FormatCells: return "formatCells";
        case OperationType::FormatColumns: return "formatColumns";
        case OperationType::FormatRows: return "formatRows";
        case OperationType::InsertColumns: return "insertColumns";
        case OperationType::InsertRows: return "insertRows";
        case OperationType::DeleteColumns: return "deleteColumns";
        case OperationType::DeleteRows: return "deleteRows";
        case OperationType::InsertHyperlinks: return "insertHyperlinks";
        case OperationType::Sort: return "sort";
        case OperationType::AutoFilter: return "autoFilter";
        case OperationType::PivotTables: return "pivotTables";
        case OperationType::Objects: return "objects";
        case OperationType::Scenarios: return "scenarios";
        default: return "unknown";
    }
}

TXSheetProtectionManager::OperationType TXSheetProtectionManager::stringToOperationType(const std::string& operationName) const {
    if (operationName == "selectLockedCells") return OperationType::SelectLockedCells;
    if (operationName == "selectUnlockedCells") return OperationType::SelectUnlockedCells;
    if (operationName == "formatCells") return OperationType::FormatCells;
    if (operationName == "formatColumns") return OperationType::FormatColumns;
    if (operationName == "formatRows") return OperationType::FormatRows;
    if (operationName == "insertColumns") return OperationType::InsertColumns;
    if (operationName == "insertRows") return OperationType::InsertRows;
    if (operationName == "deleteColumns") return OperationType::DeleteColumns;
    if (operationName == "deleteRows") return OperationType::DeleteRows;
    if (operationName == "insertHyperlinks") return OperationType::InsertHyperlinks;
    if (operationName == "sort") return OperationType::Sort;
    if (operationName == "autoFilter") return OperationType::AutoFilter;
    if (operationName == "pivotTables") return OperationType::PivotTables;
    if (operationName == "objects") return OperationType::Objects;
    if (operationName == "scenarios") return OperationType::Scenarios;
    
    return OperationType::SelectLockedCells; // 默认值
}

} // namespace TinaXlsx
