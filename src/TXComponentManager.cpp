#include "TinaXlsx/TXComponentManager.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCompactCell.hpp"
#include <algorithm>

namespace TinaXlsx {

// ==================== ComponentManager 实现 ====================

void ComponentManager::registerComponent(ExcelComponent component) {
    registered_components_.insert(component);
}

bool ComponentManager::hasComponent(ExcelComponent component) const {
    return registered_components_.find(component) != registered_components_.end();
}

const std::unordered_set<ExcelComponent>& ComponentManager::getComponents() const {
    return registered_components_;
}
    
void ComponentManager::reset() {
    registered_components_.clear();
    // 基础工作簿组件始终存在
    registerComponent(ExcelComponent::BasicWorkbook);
}

} // namespace TinaXlsx 
