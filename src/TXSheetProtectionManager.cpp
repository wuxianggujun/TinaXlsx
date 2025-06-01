#include "TinaXlsx/TXSheetProtectionManager.hpp"
#include "TinaXlsx/TXCellManager.hpp"
#include "TinaXlsx/TXSha512.hpp"
#include <functional>
#include <algorithm>
#include <cstdint>
#include <random>
#include <sstream>
#include <iomanip>

namespace TinaXlsx {

// ==================== 工作表保护 ====================

bool TXSheetProtectionManager::protectSheet(const std::string& password, const SheetProtection& protection) {
    protection_ = protection;
    protection_.isProtected = true;

    if (!password.empty()) {
        // 生成随机盐值
        protection_.saltValue = TXExcelPasswordHash::generateSalt(16);

        // 使用SHA-512算法计算密码哈希
        protection_.passwordHash = TXExcelPasswordHash::calculateHash(
            password,
            protection_.saltValue,
            protection_.spinCount
        );

        // 设置算法信息
        protection_.algorithmName = "SHA-512";
    } else {
        // 无密码保护
        protection_.passwordHash.clear();
        protection_.saltValue.clear();
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

    // 使用SHA-512算法验证密码
    return TXExcelPasswordHash::verifyPassword(
        password,
        protection_.saltValue,
        protection_.passwordHash,
        protection_.spinCount
    );
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

        // 重要：需要通过样式系统来实际应用锁定状态
        // 这里需要获取工作簿的样式管理器来更新单元格样式
        // 但是TXSheetProtectionManager没有直接访问TXWorkbook的权限
        // 所以这个功能需要在TXSheet层面来实现

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
    // 现代Excel使用SHA-512算法，但由于我们没有加密库，
    // 我们暂时使用Legacy Password Hash Algorithm作为后备
    // TODO: 实现完整的SHA-512算法以完全兼容现代Excel

    if (password.empty()) {
        return "";
    }

    // 为了兼容性，我们先使用Legacy算法
    uint16_t passwordLength = static_cast<uint16_t>(password.length());
    uint16_t passwordHash = 0;

    // 从密码的最后一个字符开始，向前遍历
    const char* pch = &password[passwordLength];
    while (pch-- != password.c_str()) {
        // 循环左移1位，保持16位
        passwordHash = ((passwordHash >> 14) & 0x01) |
                      ((passwordHash << 1) & 0x7fff);
        // 与当前字符异或
        passwordHash ^= static_cast<uint8_t>(*pch);
    }

    // 最后一次循环左移
    passwordHash = ((passwordHash >> 14) & 0x01) |
                  ((passwordHash << 1) & 0x7fff);

    // 与特殊常量异或
    passwordHash ^= (0x8000 | ('N' << 8) | 'K');

    // 与密码长度异或
    passwordHash ^= passwordLength;

    // 转换为字符串（十进制格式）
    return std::to_string(passwordHash);
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
