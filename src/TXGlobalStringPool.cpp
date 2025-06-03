//
// @file TXGlobalStringPool.cpp
// @brief 全局字符串池实现
//

#include "TinaXlsx/TXGlobalStringPool.hpp"
#include <memory>

namespace TinaXlsx {

TXGlobalStringPool& TXGlobalStringPool::instance() {
    static TXGlobalStringPool instance;
    static std::once_flag initialized;
    std::call_once(initialized, [&]() {
        instance.initializeCommonStrings();
    });
    return instance;
}

const std::string& TXGlobalStringPool::intern(const std::string& str) {
    std::lock_guard<std::mutex> lock(mutex_);
    
    // 🚀 使用insert的返回值避免重复查找
    auto [it, inserted] = pool_.insert(str);
    return *it;
}

const std::string& TXGlobalStringPool::intern(std::string_view str) {
    std::lock_guard<std::mutex> lock(mutex_);
    
    // 先查找是否已存在
    auto it = pool_.find(std::string(str));
    if (it != pool_.end()) {
        return *it;
    }
    
    // 不存在则插入
    auto [inserted_it, inserted] = pool_.insert(std::string(str));
    return *inserted_it;
}

bool TXGlobalStringPool::isInterned(const std::string& str) const {
    std::lock_guard<std::mutex> lock(mutex_);
    return pool_.find(str) != pool_.end();
}

size_t TXGlobalStringPool::size() const {
    std::lock_guard<std::mutex> lock(mutex_);
    return pool_.size();
}

void TXGlobalStringPool::clear() {
    std::lock_guard<std::mutex> lock(mutex_);
    pool_.clear();
    // 重新初始化常用字符串
    initializeCommonStrings();
}

// 🚀 常用字符串常量实现
const std::string& TXGlobalStringPool::EMPTY_STRING() {
    static const std::string& str = instance().intern(std::string(""));
    return str;
}

const std::string& TXGlobalStringPool::DEFAULT_SHEET_NAME() {
    static const std::string& str = instance().intern(std::string("Sheet1"));
    return str;
}

const std::string& TXGlobalStringPool::TRUE_STRING() {
    static const std::string& str = instance().intern(std::string("true"));
    return str;
}

const std::string& TXGlobalStringPool::FALSE_STRING() {
    static const std::string& str = instance().intern(std::string("false"));
    return str;
}

const std::string& TXGlobalStringPool::ZERO_STRING() {
    static const std::string& str = instance().intern(std::string("0"));
    return str;
}

const std::string& TXGlobalStringPool::ONE_STRING() {
    static const std::string& str = instance().intern(std::string("1"));
    return str;
}

const std::string& TXGlobalStringPool::XML_DECLARATION() {
    static const std::string& str = instance().intern(std::string("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"));
    return str;
}

const std::string& TXGlobalStringPool::WORKSHEET_NAMESPACE() {
    static const std::string& str = instance().intern(std::string("http://schemas.openxmlformats.org/spreadsheetml/2006/main"));
    return str;
}

const std::string& TXGlobalStringPool::RELATIONSHIPS_NAMESPACE() {
    static const std::string& str = instance().intern(std::string("http://schemas.openxmlformats.org/package/2006/relationships"));
    return str;
}

void TXGlobalStringPool::initializeCommonStrings() {
    // 预内化常用字符串，避免运行时查找
    static const std::vector<std::string> commonStrings = {
        "",
        "Sheet1", "Sheet2", "Sheet3",
        "true", "false",
        "0", "1", "2", "3", "4", "5", "6", "7", "8", "9",
        "A1", "B1", "C1", "D1", "E1",
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>",
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "http://schemas.openxmlformats.org/package/2006/relationships",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        "General",
        "Calibri",
        "Arial"
    };
    
    for (const auto& str : commonStrings) {
        pool_.insert(str);
    }
}

} // namespace TinaXlsx
