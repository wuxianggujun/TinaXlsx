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
#include "TinaXlsx/TXWorksheetRelsXmlHandler.hpp"
#include "TinaXlsx/TXSharedStringsXmlHandler.hpp"
#include "TinaXlsx/TXChartXmlHandler.hpp"

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
        , context_(std::make_unique<TXWorkbookContext>(sheets_, style_manager_, component_manager_, shared_strings_pool_, workbook_protection_manager_))
        , workbook_protection_manager_() {
        // 默认注册基础组件
        component_manager_.registerComponent(ExcelComponent::BasicWorkbook);
        // 默认注册SharedStrings组件，避免Content Types问题
        component_manager_.registerComponent(ExcelComponent::SharedStrings);
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
        , context_(std::move(other.context_))
        , workbook_protection_manager_(std::move(other.workbook_protection_manager_)) {
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
            workbook_protection_manager_ = std::move(other.workbook_protection_manager_);
        }
        return *this;
    }

    bool TXWorkbook::loadFromFile(const std::string& filename) {
        // 清空现有数据
        clear();

        TXZipArchiveReader zipReader;
        if (!zipReader.open(filename)) {
            last_error_ = "Failed to open XLSX file.";
            return false;
        }

        // 加载 workbook.xml（必须首先加载以获取工作表信息）
        TXWorkbookXmlHandler workbookHandler;
        auto workbookLoadResult = workbookHandler.load(zipReader, *context_);
        if (workbookLoadResult.isError()) {
            last_error_ = "Workbook load failed: " + workbookLoadResult.error().getMessage();
            return false;
        }

        // 加载 styles.xml（如果存在）
        if (component_manager_.hasComponent(ExcelComponent::Styles)) {
            StylesXmlHandler stylesHandler;
            auto stylesLoadResult = stylesHandler.load(zipReader, *context_);
            if (stylesLoadResult.isError()) {
                last_error_ = "Styles load failed: " + stylesLoadResult.error().getMessage();
                return false;
            }
        }

        // 加载 sharedStrings.xml（如果存在）
        if (component_manager_.hasComponent(ExcelComponent::SharedStrings)) {
            TXSharedStringsXmlHandler sharedStringsHandler;
            auto sharedStringsLoadResult = sharedStringsHandler.load(zipReader, *context_);
            if (sharedStringsLoadResult.isError()) {
                last_error_ = "Shared strings load failed: " + sharedStringsLoadResult.error().getMessage();
                return false;
            }
        }

        // 加载每个工作表
        for (size_t i = 0; i < sheets_.size(); ++i) {
            TXWorksheetXmlHandler worksheetHandler(i);
            auto worksheetLoadResult = worksheetHandler.load(zipReader, *context_);
            if (worksheetLoadResult.isError()) {
                last_error_ = "Worksheet " + std::to_string(i) + " load failed: " + worksheetLoadResult.error().getMessage();
                return false;
            }
        }

        // 加载其他组件（如文档属性）
        if (component_manager_.hasComponent(ExcelComponent::DocumentProperties)) {
            TXDocumentPropertiesXmlHandler docPropsHandler;
            auto docPropsLoadResult = docPropsHandler.load(zipReader, *context_);
            if (docPropsLoadResult.isError()) {
                last_error_ = "Document properties load failed: " + docPropsLoadResult.error().getMessage();
                return false;
            }
        }

        return true;
    }

    bool TXWorkbook::saveToFile(const std::string& filename) {
        // 在保存前准备组件检测
        prepareForSaving();

        TXZipArchiveWriter zipWriter;
        if (!zipWriter.open(filename, false)) {
            last_error_ = "无法创建文件: " + filename;
            return false;
        }

        // 保存 [Content_Types].xml
        TXContentTypesXmlHandler contentTypesHandler;
        auto contentTypesResult = contentTypesHandler.save(zipWriter, *context_);
        if (contentTypesResult.isError()) {
            last_error_ = "Content types save failed: " + contentTypesResult.error().getMessage();
            return false;
        }

        TXMainRelsXmlHandler mainRelsHandler;
        auto mainRelsResult = mainRelsHandler.save(zipWriter, *context_);
        if (mainRelsResult.isError()) {
            last_error_ = "Main rels save failed: " + mainRelsResult.error().getMessage();
            return false;
        }

        // 保存 workbook.xml
        TXWorkbookXmlHandler workbookHandler;
        auto workbookResult = workbookHandler.save(zipWriter, *context_);
        if (workbookResult.isError()) {
            last_error_ = "Workbook save failed: " + workbookResult.error().getMessage();
            return false;
        }

        // 保存 workbook.xml.rels
        TXWorkbookRelsXmlHandler workbookRelsHandler;
        auto workbookRelsResult = workbookRelsHandler.save(zipWriter, *context_);
        if (workbookRelsResult.isError()) {
            last_error_ = "Workbook rels save failed: " + workbookRelsResult.error().getMessage();
            return false;
        }

        // 保存 styles.xml（如果启用了样式组件）
        if (component_manager_.hasComponent(ExcelComponent::Styles)) {
            StylesXmlHandler stylesHandler;
            auto stylesResult = stylesHandler.save(zipWriter, *context_);
            if (stylesResult.isError()) {
                last_error_ = "Styles save failed: " + stylesResult.error().getMessage();
                return false;
            }
        }

        // 保存每个工作表（必须在sharedStrings之前，因为工作表保存时会填充共享字符串池）
        for (size_t i = 0; i < sheets_.size(); ++i) {
            TXWorksheetXmlHandler worksheetHandler(i);
            auto worksheetResult = worksheetHandler.save(zipWriter, *context_);
            if (worksheetResult.isError()) {
                last_error_ = "Worksheet " + std::to_string(i) + " save failed: " + worksheetResult.error().getMessage();
                return false;
            }

            // 保存工作表关系文件（如果有图表）
            const TXSheet* sheet = sheets_[i].get();
            if (sheet->getChartCount() > 0) {
                TXWorksheetRelsXmlHandler worksheetRelsHandler(static_cast<u32>(i));
                auto worksheetRelsResult = worksheetRelsHandler.save(zipWriter, *context_);
                if (worksheetRelsResult.isError()) {
                    last_error_ = "Worksheet rels " + std::to_string(i) + " save failed: " + worksheetRelsResult.error().getMessage();
                    return false;
                }

                // 保存绘图文件
                TXDrawingXmlHandler drawingHandler(static_cast<u32>(i));
                auto drawingResult = drawingHandler.save(zipWriter, *context_);
                if (drawingResult.isError()) {
                    last_error_ = "Drawing " + std::to_string(i) + " save failed: " + drawingResult.error().getMessage();
                    return false;
                }

                // 保存绘图关系文件
                TXDrawingRelsXmlHandler drawingRelsHandler(static_cast<u32>(i));
                auto drawingRelsResult = drawingRelsHandler.save(zipWriter, *context_);
                if (drawingRelsResult.isError()) {
                    last_error_ = "Drawing rels " + std::to_string(i) + " save failed: " + drawingRelsResult.error().getMessage();
                    return false;
                }

                // 保存每个图表文件
                auto charts = sheet->getAllCharts();
                for (size_t j = 0; j < charts.size(); ++j) {
                    TXChartXmlHandler chartHandler(charts[j], static_cast<u32>(j));
                    auto chartResult = chartHandler.save(zipWriter, *context_);
                    if (chartResult.isError()) {
                        last_error_ = "Chart " + std::to_string(j) + " save failed: " + chartResult.error().getMessage();
                        return false;
                    }

                    // 保存图表关系文件
                    TXChartRelsXmlHandler chartRelsHandler(static_cast<u32>(j));
                    auto chartRelsResult = chartRelsHandler.save(zipWriter, *context_);
                    if (chartRelsResult.isError()) {
                        last_error_ = "Chart rels " + std::to_string(j) + " save failed: " + chartRelsResult.error().getMessage();
                        return false;
                    }
                }
            }
        }

        // 保存 sharedStrings.xml（如果启用了共享字符串组件）
        if (component_manager_.hasComponent(ExcelComponent::SharedStrings)) {
            TXSharedStringsXmlHandler sharedStringsHandler;
            auto sharedStringsResult = sharedStringsHandler.save(zipWriter, *context_);
            if (sharedStringsResult.isError()) {
                last_error_ = "Shared strings save failed: " + sharedStringsResult.error().getMessage();
                return false;
            }
        }

        // 保存文档属性
        if (component_manager_.hasComponent(ExcelComponent::DocumentProperties)) {
            TXDocumentPropertiesXmlHandler docPropsHandler;
            auto docPropsResult = docPropsHandler.save(zipWriter, *context_);
            if (docPropsResult.isError()) {
                last_error_ = "Document properties save failed: " + docPropsResult.error().getMessage();
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
        context_ = std::make_unique<TXWorkbookContext>(sheets_, style_manager_, component_manager_, shared_strings_pool_, workbook_protection_manager_);
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
        // 扫描所有工作表，检测需要的组件
        bool hasStringCells = false;
        bool hasMergedCells = false;
        bool hasStyledCells = false;
        
        for (const auto& sheet : sheets_) {
            if (!sheet) continue;
            
            // 检查是否有合并单元格
            if (sheet->getMergeCount() > 0) {
                hasMergedCells = true;
            }
            
            // 获取已使用的范围
            auto usedRange = sheet->getUsedRange();
            if (usedRange.isValid()) {
                auto start = usedRange.getStart();
                auto end = usedRange.getEnd();
                
                for (row_t row = start.getRow(); row <= end.getRow(); ++row) {
                    for (column_t col = start.getCol(); col <= end.getCol(); ++col) {
                        const TXCell* cell = sheet->getCell(row, col);
                        if (!cell || cell->isEmpty()) continue;
                        
                        // 检查是否有字符串值
                        if (cell->getType() == TXCell::CellType::String) {
                            hasStringCells = true;
                        }
                        
                        // 检查是否有样式
                        if (cell->getStyleIndex() != 0) {
                            hasStyledCells = true;
                        }
                    }
                }
            }
        }
        
        // 根据检测结果注册组件
        if (hasStringCells) {
            component_manager_.registerComponent(ExcelComponent::SharedStrings);
        }
        
        if (hasMergedCells) {
            component_manager_.registerComponent(ExcelComponent::MergedCells);
        }
        
        if (hasStyledCells) {
            component_manager_.registerComponent(ExcelComponent::Styles);
        }
    }

    // ==================== 工作簿保护功能实现 ====================

    TXWorkbookProtectionManager& TXWorkbook::getWorkbookProtectionManager() {
        return workbook_protection_manager_;
    }

    const TXWorkbookProtectionManager& TXWorkbook::getWorkbookProtectionManager() const {
        return workbook_protection_manager_;
    }

    bool TXWorkbook::protectWorkbook(const std::string& password,
                                   const TXWorkbookProtectionManager::WorkbookProtection& protection) {
        return workbook_protection_manager_.protectWorkbook(password, protection);
    }

    bool TXWorkbook::unprotectWorkbook(const std::string& password) {
        return workbook_protection_manager_.unprotectWorkbook(password);
    }

    bool TXWorkbook::isWorkbookProtected() const {
        return workbook_protection_manager_.isWorkbookProtected();
    }

    bool TXWorkbook::protectStructure(const std::string& password) {
        return workbook_protection_manager_.protectStructure(password);
    }

    bool TXWorkbook::protectWindows(const std::string& password) {
        return workbook_protection_manager_.protectWindows(password);
    }

} // namespace TinaXlsx
