#include "TinaXlsx/TXStyleTemplate.hpp"
#include <unordered_map>
#include <sstream>
#include <fstream>

namespace TinaXlsx {

// ==================== TXStyleTemplate 实现 ====================

class TXStyleTemplate::Impl {
public:
    StyleTemplateInfo info_;
    TXCellStyle baseStyle_;
    std::unordered_map<std::string, TXCellStyle> namedStyles_;
    
    Impl() {}
    
    Impl(const Impl& other) 
        : info_(other.info_)
        , baseStyle_(other.baseStyle_)
        , namedStyles_(other.namedStyles_) {}
};

TXStyleTemplate::TXStyleTemplate() : pImpl(std::make_unique<Impl>()) {}

TXStyleTemplate::TXStyleTemplate(const StyleTemplateInfo& info) 
    : pImpl(std::make_unique<Impl>()) {
    pImpl->info_ = info;
}

TXStyleTemplate::~TXStyleTemplate() = default;

TXStyleTemplate::TXStyleTemplate(const TXStyleTemplate& other) 
    : pImpl(std::make_unique<Impl>(*other.pImpl)) {}

TXStyleTemplate& TXStyleTemplate::operator=(const TXStyleTemplate& other) {
    if (this != &other) {
        pImpl = std::make_unique<Impl>(*other.pImpl);
    }
    return *this;
}

TXStyleTemplate::TXStyleTemplate(TXStyleTemplate&& other) noexcept 
    : pImpl(std::move(other.pImpl)) {}

TXStyleTemplate& TXStyleTemplate::operator=(TXStyleTemplate&& other) noexcept {
    if (this != &other) {
        pImpl = std::move(other.pImpl);
    }
    return *this;
}

// ==================== TXStyleTemplate 基本访问器实现 ====================

std::string TXStyleTemplate::getId() const {
    return pImpl->info_.id;
}

std::string TXStyleTemplate::getName() const {
    return pImpl->info_.name;
}

// Getters
const StyleTemplateInfo& TXStyleTemplate::getInfo() const {
    return pImpl->info_;
}

const TXCellStyle& TXStyleTemplate::getBaseStyle() const {
    return pImpl->baseStyle_;
}

std::vector<std::string> TXStyleTemplate::getNamedStyleNames() const {
    std::vector<std::string> names;
    names.reserve(pImpl->namedStyles_.size());
    for (const auto& pair : pImpl->namedStyles_) {
        names.push_back(pair.first);
    }
    return names;
}

// Setters
void TXStyleTemplate::setInfo(const StyleTemplateInfo& info) {
    pImpl->info_ = info;
}

void TXStyleTemplate::setBaseStyle(const TXCellStyle& style) {
    pImpl->baseStyle_ = style;
}

// Named styles management
void TXStyleTemplate::addNamedStyle(const std::string& name, const TXCellStyle& style) {
    pImpl->namedStyles_[name] = style;
}

void TXStyleTemplate::removeNamedStyle(const std::string& name) {
    pImpl->namedStyles_.erase(name);
}

const TXCellStyle* TXStyleTemplate::getNamedStyle(const std::string& name) const {
    auto it = pImpl->namedStyles_.find(name);
    return (it != pImpl->namedStyles_.end()) ? &it->second : nullptr;
}



void TXStyleTemplate::clearNamedStyles() {
    pImpl->namedStyles_.clear();
}













// ==================== TXStyleTemplateManager 实现 ====================

class TXStyleTemplateManager::Impl {
public:
    std::unordered_map<std::string, TXStyleTemplate> templates_;
    std::string defaultTemplateName_;
    
    Impl() {
        createBuiltinTemplates();
    }
    
    void createBuiltinTemplates();
};

TXStyleTemplateManager& TXStyleTemplateManager::getInstance() {
    static TXStyleTemplateManager instance;
    return instance;
}

TXStyleTemplateManager::TXStyleTemplateManager() : pImpl(std::make_unique<Impl>()) {}
TXStyleTemplateManager::~TXStyleTemplateManager() = default;

// Template management
bool TXStyleTemplateManager::registerTemplate(const TXStyleTemplate& styleTemplate) {
    // 从模板中获取实际的ID作为key
    std::string templateId = styleTemplate.getId();
    if (templateId.empty()) {
        return false; // 无效的模板ID
    }
    auto result = pImpl->templates_.emplace(templateId, styleTemplate);
    return result.second;
}

bool TXStyleTemplateManager::unregisterTemplate(const std::string& name) {
    return pImpl->templates_.erase(name) > 0;
}



const TXStyleTemplate* TXStyleTemplateManager::getTemplate(const std::string& name) const {
    auto it = pImpl->templates_.find(name);
    return (it != pImpl->templates_.end()) ? &it->second : nullptr;
}

TXStyleTemplate* TXStyleTemplateManager::getTemplate(const std::string& name) {
    auto it = pImpl->templates_.find(name);
    return (it != pImpl->templates_.end()) ? &it->second : nullptr;
}

bool TXStyleTemplateManager::hasTemplate(const std::string& name) const {
    return pImpl->templates_.find(name) != pImpl->templates_.end();
}









// Built-in templates creation
void TXStyleTemplateManager::Impl::createBuiltinTemplates() {
    // Professional template
    {
        StyleTemplateInfo info("professional", "Professional", StyleTemplateCategory::Custom);
        info.description = "Professional business style";
        info.version = "1.0";
        
        TXStyleTemplate tmpl(info);
        TXCellStyle baseStyle;
        
        // 设置专业样式
        TXFont font;
        font.setName("Arial");
        font.setSize(11);
        font.setColor(TXColor(0, 0, 0));
        baseStyle.setFont(font);
        
        TXAlignment alignment;
        alignment.setHorizontal(HorizontalAlignment::Left);
        alignment.setVertical(VerticalAlignment::Middle);
        baseStyle.setAlignment(alignment);
        
        tmpl.setBaseStyle(baseStyle);
        
        // 添加标题样式
        TXCellStyle headerStyle = baseStyle;
        TXFont headerFont = font;
        headerFont.setBold(true);
        headerFont.setSize(14);
        headerStyle.setFont(headerFont);
        
        TXFill headerFill;
        headerFill.setPattern(FillPattern::Solid);
        headerFill.setForegroundColor(TXColor(220, 220, 220));
        headerStyle.setFill(headerFill);
        
        tmpl.addNamedStyle("Header", headerStyle);
        
        templates_["Professional"] = tmpl;
    }
    
    // Modern template
    {
        StyleTemplateInfo info("modern", "Modern", StyleTemplateCategory::Custom);
        info.description = "Modern clean style";
        info.version = "1.0";
        
        TXStyleTemplate tmpl(info);
        TXCellStyle baseStyle;
        
        TXFont font;
        font.setName("Segoe UI");
        font.setSize(10);
        font.setColor(TXColor(68, 68, 68));
        baseStyle.setFont(font);
        
        tmpl.setBaseStyle(baseStyle);
        templates_["Modern"] = tmpl;
    }
    
    // Classic template
    {
        StyleTemplateInfo info("classic", "Classic", StyleTemplateCategory::Custom);
        info.description = "Classic traditional style";
        info.version = "1.0";
        
        TXStyleTemplate tmpl(info);
        TXCellStyle baseStyle;
        
        TXFont font;
        font.setName("Times New Roman");
        font.setSize(12);
        font.setColor(TXColor(0, 0, 0));
        baseStyle.setFont(font);
        
        tmpl.setBaseStyle(baseStyle);
        templates_["Classic"] = tmpl;
    }
    
    defaultTemplateName_ = "Professional";
}

} // namespace TinaXlsx