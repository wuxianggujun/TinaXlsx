//
// @file TXWorkbook.cpp - 移除impl模式的新实现
//

#include <algorithm>
#include <regex>

#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXWorkbookContext.hpp"

#include "TinaXlsx/TXStylesXmlHandler.hpp"
#include "TinaXlsx/TXWorkbookXmlHandler.hpp"
#include "TinaXlsx/TXWorksheetXmlHandler.hpp"
#include "TinaXlsx/TXDocumentPropertiesXmlHandler.hpp"
#include "TinaXlsx/TXContentTypesXmlHandler.hpp"
#include "TinaXlsx/TXMainRelsXmlHandler.hpp"
#include "TinaXlsx/TXWorkbookRelsXmlHandler.hpp"
#include "TinaXlsx/TXSharedStringsXmlHandler.hpp"

namespace TinaXlsx
{
    // ==================== TXWorkbook 实现 ====================

    TXWorkbook::TXWorkbook() 
        : sheets_()
        , active_sheet_index_(0)
        , last_error_()
        , component_manager_()
        , auto_component_detection_(true)
        , style_manager_()
        , shared_strings_pool_()
        , context_(std::make_unique<TXWorkbookContext>(sheets_, style_manager_, component_manager_, shared_strings_pool_)) {
        component_manager_.registerComponent(ExcelComponent::BasicWorkbook);
    }

    TXWorkbook::~TXWorkbook() = default;

    TXWorkbook::TXWorkbook(TXWorkbook&& other) noexcept
        : sheets_(std::move(other.sheets_))
        , active_sheet_index_(other.active_sheet_index_)
        , last_error_(std::move(other.last_error_))
        , component_manager_(std::move(other.component_manager_))
        , auto_component_detection_(other.auto_component_detection_)
        , style_manager_(std::move(other.style_manager_))
        , shared_strings_pool_(std::move(other.shared_strings_pool_))
        , context_(std::move(other.context_)) {
    }

    TXWorkbook& TXWorkbook::operator=(TXWorkbook&& other) noexcept {
        if (this != &other) {
            sheets_ = std::move(other.sheets_);
            active_sheet_index_ = other.active_sheet_index_;
            last_error_ = std::move(other.last_error_);
            component_manager_ = std::move(other.component_manager_);
            auto_component_detection_ = other.auto_component_detection_;
            style_manager_ = std::move(other.style_manager_);
            shared_strings_pool_ = std::move(other.shared_strings_pool_);
            context_ = std::move(other.context_);
        }
        return *this;
    }

    bool TXWorkbook::loadFromFile(const std::string& filename) {
        // 清空现有数据
        clear();

        TXZipArchiveReader zipReader;
        if (!zipReader.open(filename)) {
            last_error_ = "Failed to open XLSX file: " + zipReader.lastError();
            return false;
        }

        // 加载 workbook.xml（必须首先加载以获取工作表信息）
        TXWorkbookXmlHandler workbookHandler;
        if (!workbookHandler.load(zipReader, *context_)) {
            last_error_ = workbookHandler.lastError();
            return false;
        }

        // 加载 styles.xml（如果存在）
        if (component_manager_.hasComponent(ExcelComponent::Styles)) {
            StylesXmlHandler stylesHandler;
            if (!stylesHandler.load(zipReader, *context_)) {
                last_error_ = stylesHandler.lastError();
                return false;
            }
        }

        // 加载 sharedStrings.xml（如果存在）
        if (component_manager_.hasComponent(ExcelComponent::SharedStrings)) {
            TXSharedStringsXmlHandler sharedStringsHandler;
            if (!sharedStringsHandler.load(zipReader, *context_)) {
                last_error_ = sharedStringsHandler.lastError();
                return false;
            }
        }

        // 加载每个工作表
        for (size_t i = 0; i < sheets_.size(); ++i) {
            TXWorksheetXmlHandler worksheetHandler(i);
            if (!worksheetHandler.load(zipReader, *context_)) {
                last_error_ = worksheetHandler.lastError();
                return false;
            }
        }

        // 加载其他组件（如文档属性）
        if (component_manager_.hasComponent(ExcelComponent::DocumentProperties)) {
            TXDocumentPropertiesXmlHandler docPropsHandler;
            if (!docPropsHandler.load(zipReader, *context_)) {
                last_error_ = docPropsHandler.lastError();
                return false;
            }
        }

        return true;
    }

    bool TXWorkbook::saveToFile(const std::string& filename) {
        TXZipArchiveWriter zipWriter;
        if (!zipWriter.open(filename, false)) {
            last_error_ = "无法创建文件: " + filename;
            return false;
        }
        
        // 保存 [Content_Types].xml
        TXContentTypesXmlHandler contentTypesHandler;
        if (!contentTypesHandler.save(zipWriter, *context_)) {
            last_error_ = contentTypesHandler.lastError();
            return false;
        }

        TXMainRelsXmlHandler mainRelsHandler;
        if (!mainRelsHandler.save(zipWriter, *context_)) {
            last_error_ = mainRelsHandler.lastError();
            return false;
        }

        // 保存 workbook.xml
        TXWorkbookXmlHandler workbookHandler;
        if (!workbookHandler.save(zipWriter, *context_)) {
            last_error_ = workbookHandler.lastError();
            return false;
        }

        // 保存 workbook.xml.rels
        TXWorkbookRelsXmlHandler workbookRelsHandler;
        if (!workbookRelsHandler.save(zipWriter, *context_)) {
            last_error_ = workbookRelsHandler.lastError();
            return false;
        }

        // 保存 styles.xml（如果启用了样式组件）
        if (component_manager_.hasComponent(ExcelComponent::Styles)) {
            StylesXmlHandler stylesHandler;
            if (!stylesHandler.save(zipWriter, *context_)) {
                last_error_ = stylesHandler.lastError();
                return false;
            }
        }

        // 保存 sharedStrings.xml（如果启用了共享字符串组件）
        if (component_manager_.hasComponent(ExcelComponent::SharedStrings)) {
            TXSharedStringsXmlHandler sharedStringsHandler;
            if (!sharedStringsHandler.save(zipWriter, *context_)) {
                last_error_ = sharedStringsHandler.lastError();
                return false;
            }
        }

        // 保存每个工作表
        for (size_t i = 0; i < sheets_.size(); ++i) {
            TXWorksheetXmlHandler worksheetHandler(i);
            if (!worksheetHandler.save(zipWriter, *context_)) {
                last_error_ = worksheetHandler.lastError();
                return false;
            }
        }

        // 保存文档属性
        if (component_manager_.hasComponent(ExcelComponent::DocumentProperties)) {
            TXDocumentPropertiesXmlHandler docPropsHandler;
            if (!docPropsHandler.save(zipWriter, *context_)) {
                last_error_ = docPropsHandler.lastError();
                return false;
            }
        }

        return true;
    }

    TXSheet* TXWorkbook::storeSheet(std::unique_ptr<TXSheet> sheet_uptr) {
        if (!sheet_uptr) {
            last_error_ = "Attempted to store a null sheet.";
            return nullptr;
        }
        
        TXSheet* sheet_ptr = sheet_uptr.get();
        sheets_.push_back(std::move(sheet_uptr));

        // 更新组件使用记录
        if (auto_component_detection_) {
            component_manager_.registerComponent(ExcelComponent::BasicWorkbook);
        }

        // 如果这是第一个工作表，则将其设为活动工作表
        if (sheets_.size() == 1) {
            active_sheet_index_ = 0;
        }
        
        last_error_.clear();
        return sheet_ptr;
    }

    TXSheet* TXWorkbook::addSheet(const std::string& name) {
        auto sheet_obj = std::make_unique<TXSheet>(name, this);
        return storeSheet(std::move(sheet_obj));
    }

    TXSheet* TXWorkbook::addSheet(std::unique_ptr<TXSheet> sheet) {
        return storeSheet(std::move(sheet));
    }

    bool TXWorkbook::removeSheet(const std::string& name) {
        auto it = std::find_if(sheets_.begin(), sheets_.end(),
                               [&name](const std::unique_ptr<TXSheet>& sheet) {
                                   return sheet->getName() == name;
                               });

        if (it == sheets_.end()) {
            last_error_ = "Sheet not found: " + name;
            return false;
        }

        sheets_.erase(it);

        // 调整活动工作表索引
        if (active_sheet_index_ >= sheets_.size() && !sheets_.empty()) {
            active_sheet_index_ = sheets_.size() - 1;
        } else if (sheets_.empty()) {
            active_sheet_index_ = 0;
        }

        return true;
    }

    bool TXWorkbook::hasSheet(const std::string& name) const {
        return std::find_if(sheets_.begin(), sheets_.end(),
                            [&name](const std::unique_ptr<TXSheet>& sheet) {
                                return sheet->getName() == name;
                            }) != sheets_.end();
    }

    bool TXWorkbook::renameSheet(const std::string& oldName, const std::string& newName) {
        auto sheet = getSheet(oldName);
        if (!sheet) {
            last_error_ = "Sheet not found: " + oldName;
            return false;
        }

        // 检查新名称是否已存在
        if (hasSheet(newName)) {
            last_error_ = "Sheet name already exists: " + newName;
            return false;
        }

        sheet->setName(newName);
        return true;
    }

    TXSheet* TXWorkbook::getSheet(u64 index) const {
        if (index < sheets_.size()) {
            return sheets_[index].get();
        }
        return nullptr;
    }

    TXSheet* TXWorkbook::getSheet(const std::string& name) {
        auto it = std::find_if(sheets_.begin(), sheets_.end(),
                               [&name](const std::unique_ptr<TXSheet>& sheet) {
                                   return sheet->getName() == name;
                               });

        return (it != sheets_.end()) ? it->get() : nullptr;
    }

    u64 TXWorkbook::getSheetCount() const {
        return sheets_.size();
    }

    std::vector<std::string> TXWorkbook::getSheetNames() const {
        std::vector<std::string> names;
        names.reserve(sheets_.size());

        for (const auto& sheet : sheets_) {
            names.push_back(sheet->getName());
        }

        return names;
    }

    bool TXWorkbook::setActiveSheet(u64 index) {
        if (index < sheets_.size()) {
            active_sheet_index_ = index;
            return true;
        }
        last_error_ = "Sheet index out of range";
        return false;
    }

    bool TXWorkbook::setActiveSheet(const std::string& name) {
        auto it = std::find_if(sheets_.begin(), sheets_.end(),
                               [&name](const std::unique_ptr<TXSheet>& sheet) {
                                   return sheet->getName() == name;
                               });

        if (it != sheets_.end()) {
            active_sheet_index_ = std::distance(sheets_.begin(), it);
            return true;
        }

        last_error_ = "Sheet not found: " + name;
        return false;
    }

    u64 TXWorkbook::getActiveSheetIndex() const {
        return active_sheet_index_;
    }

    TXSheet* TXWorkbook::getActiveSheet() {
        return getSheet(active_sheet_index_);
    }

    const std::string& TXWorkbook::getLastError() const {
        return last_error_;
    }

    void TXWorkbook::clear() {
        sheets_.clear();
        active_sheet_index_ = 0;
        last_error_.clear();
        component_manager_.reset();
        style_manager_ = std::move(TXStyleManager());
        shared_strings_pool_ = TXSharedStringsPool();
        context_ = std::make_unique<TXWorkbookContext>(sheets_, style_manager_, component_manager_, shared_strings_pool_);
    }

    ComponentManager& TXWorkbook::getComponentManager() {
        return component_manager_;
    }

    const ComponentManager& TXWorkbook::getComponentManager() const {
        return component_manager_;
    }

    void TXWorkbook::setAutoComponentDetection(bool enable) {
        auto_component_detection_ = enable;
    }

    void TXWorkbook::registerComponent(ExcelComponent component) {
        component_manager_.registerComponent(component);
    }

    std::vector<std::unique_ptr<TXSheet>>& TXWorkbook::getSheets() {
        return sheets_;
    }

    const std::vector<std::unique_ptr<TXSheet>>& TXWorkbook::getSheets() const {
        return sheets_;
    }

    TXStyleManager& TXWorkbook::getStyleManager() {
        return style_manager_;
    }

    const TXStyleManager& TXWorkbook::getStyleManager() const {
        return style_manager_;
    }

    u32 TXWorkbook::registerOrGetStyleFId(const TXCellStyle& style) {
        return style_manager_.registerCellStyleXF(style);
    }

    TXWorkbookContext* TXWorkbook::getContext() {
        return context_.get();
    }

    bool TXWorkbook::isEmpty() const {
        return sheets_.empty();
    }

    void TXWorkbook::prepareForSaving() {
        // 这里可以添加保存前的准备工作
        // 例如更新共享字符串池等
    }

} // namespace TinaXlsx
