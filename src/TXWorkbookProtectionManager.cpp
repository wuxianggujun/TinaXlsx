#include "TinaXlsx/TXWorkbookProtectionManager.hpp"
#include "TinaXlsx/TXSha512.hpp"

namespace TinaXlsx {

bool TXWorkbookProtectionManager::protectWorkbook(const std::string& password, const WorkbookProtection& protection) {
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

bool TXWorkbookProtectionManager::unprotectWorkbook(const std::string& password) {
    if (!protection_.isProtected) {
        return true; // 没有保护，直接成功
    }
    
    if (verifyPassword(password)) {
        protection_.isProtected = false;
        protection_.passwordHash.clear();
        protection_.saltValue.clear();
        return true;
    }
    
    return false; // 密码错误
}

bool TXWorkbookProtectionManager::verifyPassword(const std::string& password) const {
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

bool TXWorkbookProtectionManager::isWorkbookProtected() const {
    return protection_.isProtected;
}

const TXWorkbookProtectionManager::WorkbookProtection& TXWorkbookProtectionManager::getWorkbookProtection() const {
    return protection_;
}

void TXWorkbookProtectionManager::setWorkbookProtection(const WorkbookProtection& protection) {
    protection_ = protection;
}

std::string TXWorkbookProtectionManager::generatePasswordHash(const std::string& password) const {
    if (password.empty()) {
        return "";
    }
    
    // 生成临时盐值用于哈希计算
    std::string tempSalt = TXExcelPasswordHash::generateSalt(16);
    return TXExcelPasswordHash::calculateHash(password, tempSalt, 100000);
}

// 便捷方法实现

bool TXWorkbookProtectionManager::protectStructure(const std::string& password) {
    WorkbookProtection protection;
    protection.lockStructure = true;
    protection.lockWindows = false;
    protection.lockRevision = false;
    return protectWorkbook(password, protection);
}

bool TXWorkbookProtectionManager::protectWindows(const std::string& password) {
    WorkbookProtection protection;
    protection.lockStructure = false;
    protection.lockWindows = true;
    protection.lockRevision = false;
    return protectWorkbook(password, protection);
}

bool TXWorkbookProtectionManager::protectRevision(const std::string& password) {
    WorkbookProtection protection;
    protection.lockStructure = false;
    protection.lockWindows = false;
    protection.lockRevision = true;
    return protectWorkbook(password, protection);
}

bool TXWorkbookProtectionManager::isStructureProtected() const {
    return protection_.isProtected && protection_.lockStructure;
}

bool TXWorkbookProtectionManager::isWindowsProtected() const {
    return protection_.isProtected && protection_.lockWindows;
}

bool TXWorkbookProtectionManager::isRevisionProtected() const {
    return protection_.isProtected && protection_.lockRevision;
}

} // namespace TinaXlsx
