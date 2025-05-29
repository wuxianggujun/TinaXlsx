//
// @file TXWorkbook.cpp - 新架构实现
//

#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"

#include "TinaXlsx/TXStylesXmlHandler.hpp"
#include "TinaXlsx/TXWorkbookXmlHandler.hpp"
#include "TinaXlsx/TXWorksheetXmlHandler.hpp"
#include "TinaXlsx/TXDocumentPropertiesXmlHandler.hpp"

#include "TinaXlsx/TXContentTypesXmlHandler.hpp"
#include "TinaXlsx/TXMainRelsXmlHandler.hpp"
#include "TinaXlsx/TXWorkbookRelsXmlHandler.hpp"
#include "TinaXlsx/TXSharedStringsXmlHandler.hpp"

#include <algorithm>
#include <regex>

namespace TinaXlsx
{
    class TXWorkbook::Impl
    {
    public:
        Impl() : active_sheet_index_(0), auto_component_detection_(true),
                 style_manager_pimpl_(std::make_unique<TXStyleManager>())
        {
            component_manager_.registerComponent(ExcelComponent::BasicWorkbook);
        }

        ~Impl() = default;

        bool loadFromFile(const std::string& filename)
        {
            // 清空现有数据
            clear();

            TXZipArchiveReader zipReader;
            if (!zipReader.open(filename))
            {
                last_error_ = "Failed to open XLSX file: " + zipReader.lastError();
                return false;
            }

            TXWorkbookContext context{sheets_, *style_manager_pimpl_, component_manager_};

            // 加载 workbook.xml（必须首先加载以获取工作表信息）
            TXWorkbookXmlHandler workbookHandler;
            if (!workbookHandler.load(zipReader, context))
            {
                last_error_ = workbookHandler.lastError();
                return false;
            }

            // 加载 styles.xml（如果存在）
            if (component_manager_.hasComponent(ExcelComponent::Styles))
            {
                StylesXmlHandler stylesHandler;
                if (!stylesHandler.load(zipReader, context))
                {
                    last_error_ = stylesHandler.lastError();
                    return false;
                }
            }

            // 加载 sharedStrings.xml（如果存在）
            if (component_manager_.hasComponent(ExcelComponent::SharedStrings))
            {
                TXSharedStringsXmlHandler sharedStringsHandler;
                if (!sharedStringsHandler.load(zipReader, context))
                {
                    last_error_ = sharedStringsHandler.lastError();
                    return false;
                }
            }

            // 加载每个工作表
            for (size_t i = 0; i < sheets_.size(); ++i)
            {
                TXWorksheetXmlHandler worksheetHandler(i);
                if (!worksheetHandler.load(zipReader, context))
                {
                    last_error_ = worksheetHandler.lastError();
                    return false;
                }
            }

            // 加载其他组件（如文档属性）
            if (component_manager_.hasComponent(ExcelComponent::DocumentProperties))
            {
                TXDocumentPropertiesXmlHandler docPropsHandler;
                if (!docPropsHandler.load(zipReader, context))
                {
                    last_error_ = docPropsHandler.lastError();
                    return false;
                }
            }
            return true;
        }

        bool saveToFile(const std::string& filename)
        {
            TXZipArchiveWriter zipWriter;
            if (!zipWriter.open(filename, false))
            {
                last_error_ = "无法创建文件: " + filename;
                return false;
            }

            if (auto_component_detection_)
            {
                for (const auto& sheet : sheets_)
                {
                    component_manager_.autoDetectComponents(sheet.get());
                }
            }


            TXWorkbookContext context{sheets_, *style_manager_pimpl_, component_manager_};

            // 保存 [Content_Types].xml
            TXContentTypesXmlHandler contentTypesHandler;
            if (!contentTypesHandler.save(zipWriter, context))
            {
                last_error_ = contentTypesHandler.lastError();
                return false;
            }

            TXMainRelsXmlHandler mainRelsHandler;
            if (!mainRelsHandler.save(zipWriter, context))
            {
                last_error_ = mainRelsHandler.lastError();
                return false;
            }

            // 保存 workbook.xml
            TXWorkbookXmlHandler workbookHandler;
            if (!workbookHandler.save(zipWriter, context))
            {
                last_error_ = workbookHandler.lastError();
                return false;
            }

            // 保存 styles.xml（如果启用了样式组件）
            if (component_manager_.hasComponent(ExcelComponent::Styles))
            {
                StylesXmlHandler stylesHandler;
                if (!stylesHandler.save(zipWriter, context))
                {
                    last_error_ = stylesHandler.lastError();
                    return false;
                }
            }

            // 保存 sharedStrings.xml（如果启用了共享字符串组件）
            if (component_manager_.hasComponent(ExcelComponent::SharedStrings))
            {
                TXSharedStringsXmlHandler sharedStringsHandler;
                if (!sharedStringsHandler.save(zipWriter, context))
                {
                    last_error_ = sharedStringsHandler.lastError();
                    return false;
                }
            }

            // 保存每个工作表
            for (size_t i = 0; i < sheets_.size(); ++i)
            {
                TXWorksheetXmlHandler worksheetHandler(i);
                if (!worksheetHandler.save(zipWriter, context))
                {
                    last_error_ = worksheetHandler.lastError();
                    return false;
                }
            }

            // 保存文档属性
            if (component_manager_.hasComponent(ExcelComponent::DocumentProperties))
            {
                TXDocumentPropertiesXmlHandler docPropsHandler;
                if (!docPropsHandler.save(zipWriter, context))
                {
                    last_error_ = docPropsHandler.lastError();
                    return false;
                }
            }

            return true;
        }

        TXSheet* storeSheet(std::unique_ptr<TXSheet> sheet_uptr)
        {
            if (!sheet_uptr)
            {
                last_error_ = "Attempted to store a null sheet.";
                return nullptr;
            }
            TXSheet* sheet_ptr = sheet_uptr.get();
            sheets_.push_back(std::move(sheet_uptr));

            // 更新组件使用记录
            if (auto_component_detection_)
            {
                component_manager_.registerComponent(ExcelComponent::BasicWorkbook);
            }

            // 如果这是第一个工作表，则将其设为活动工作表
            if (sheets_.size() == 1)
            {
                active_sheet_index_ = 0;
            }
            last_error_.clear();
            return sheet_ptr;
        }


        bool removeSheet(const std::string& name)
        {
            auto it = std::find_if(sheets_.begin(), sheets_.end(),
                                   [&name](const std::unique_ptr<TXSheet>& sheet)
                                   {
                                       return sheet->getName() == name;
                                   });

            if (it == sheets_.end())
            {
                last_error_ = "Sheet not found: " + name;
                return false;
            }

            sheets_.erase(it);

            // 调整活动工作表索引
            if (active_sheet_index_ >= sheets_.size() && !sheets_.empty())
            {
                active_sheet_index_ = sheets_.size() - 1;
            }
            else if (sheets_.empty())
            {
                active_sheet_index_ = 0;
            }

            return true;
        }

        bool hasSheet(const std::string& name) const
        {
            return std::find_if(sheets_.begin(), sheets_.end(),
                                [&name](const std::unique_ptr<TXSheet>& sheet)
                                {
                                    return sheet->getName() == name;
                                }) != sheets_.end();
        }

        bool renameSheet(const std::string& oldName, const std::string& newName)
        {
            auto sheet = getSheet(oldName);
            if (!sheet)
            {
                last_error_ = "Sheet not found: " + oldName;
                return false;
            }

            // 检查新名称是否已存在
            if (hasSheet(newName))
            {
                last_error_ = "Sheet name already exists: " + newName;
                return false;
            }

            sheet->setName(newName);
            return true;
        }

        TXSheet* getSheet(std::size_t index)
        {
            if (index < sheets_.size())
            {
                return sheets_[index].get();
            }
            return nullptr;
        }

        TXSheet* getSheet(const std::string& name)
        {
            auto it = std::find_if(sheets_.begin(), sheets_.end(),
                                   [&name](const std::unique_ptr<TXSheet>& sheet)
                                   {
                                       return sheet->getName() == name;
                                   });

            return (it != sheets_.end()) ? it->get() : nullptr;
        }

        const TXSheet* getSheet(std::size_t index) const
        {
            if (index < sheets_.size())
            {
                return sheets_[index].get();
            }
            return nullptr;
        }

        const TXSheet* getSheet(const std::string& name) const
        {
            auto it = std::find_if(sheets_.begin(), sheets_.end(),
                                   [&name](const std::unique_ptr<TXSheet>& sheet)
                                   {
                                       return sheet->getName() == name;
                                   });

            return (it != sheets_.end()) ? it->get() : nullptr;
        }

        std::size_t getSheetCount() const
        {
            return sheets_.size();
        }

        std::vector<std::string> getSheetNames() const
        {
            std::vector<std::string> names;
            names.reserve(sheets_.size());

            for (const auto& sheet : sheets_)
            {
                names.push_back(sheet->getName());
            }

            return names;
        }

        bool setActiveSheet(std::size_t index)
        {
            if (index < sheets_.size())
            {
                active_sheet_index_ = index;
                return true;
            }
            last_error_ = "Sheet index out of range";
            return false;
        }

        bool setActiveSheet(const std::string& name)
        {
            auto it = std::find_if(sheets_.begin(), sheets_.end(),
                                   [&name](const std::unique_ptr<TXSheet>& sheet)
                                   {
                                       return sheet->getName() == name;
                                   });

            if (it != sheets_.end())
            {
                active_sheet_index_ = std::distance(sheets_.begin(), it);
                return true;
            }

            last_error_ = "Sheet not found: " + name;
            return false;
        }

        std::size_t getActiveSheetIndex() const
        {
            return active_sheet_index_;
        }

        TXSheet* getActiveSheet()
        {
            return getSheet(active_sheet_index_);
        }

        const TXSheet* getActiveSheet() const
        {
            return getSheet(active_sheet_index_);
        }

        TXStyleManager* getStyleManager() const
        {
            return style_manager_pimpl_.get();
        }

        const std::string& getLastError() const
        {
            return last_error_;
        }

        void clear()
        {
            sheets_.clear();
            active_sheet_index_ = 0;
            last_error_.clear();
            component_manager_.reset();
            style_manager_pimpl_ = std::make_unique<TXStyleManager>();
        }

        ComponentManager& getComponentManager()
        {
            return component_manager_;
        }

        const ComponentManager& getComponentManager() const
        {
            return component_manager_;
        }

        void setAutoComponentDetection(bool enable)
        {
            auto_component_detection_ = enable;
        }

        void registerComponent(ExcelComponent component)
        {
            component_manager_.registerComponent(component);
        }

        std::vector<std::unique_ptr<TXSheet>>& getSheets()
        {
            return sheets_;
        }

        const std::vector<std::unique_ptr<TXSheet>>& getSheets() const
        {
            return sheets_;
        }

        TXStyleManager& getStyleManagerRef()
        {
            return *style_manager_pimpl_;
        }

        const TXStyleManager& getStyleManagerRef() const
        {
            return *style_manager_pimpl_;
        }

    private:
        std::vector<std::unique_ptr<TXSheet>> sheets_;
        std::size_t active_sheet_index_;
        std::string last_error_;
        ComponentManager component_manager_;
        bool auto_component_detection_;
        std::unique_ptr<TXStyleManager> style_manager_pimpl_;
    };

    // TXWorkbook 实现

    TXWorkbook::TXWorkbook() : pImpl(std::make_unique<Impl>())
    {
    }

    TXWorkbook::~TXWorkbook() = default;

    TXWorkbook::TXWorkbook(TXWorkbook&& other) noexcept
        : pImpl(std::move(other.pImpl))
    {
    }

    TXWorkbook& TXWorkbook::operator=(TXWorkbook&& other) noexcept
    {
        if (this != &other)
        {
            pImpl = std::move(other.pImpl);
        }
        return *this;
    }

    bool TXWorkbook::loadFromFile(const std::string& filename)
    {
        return pImpl->loadFromFile(filename);
    }

    bool TXWorkbook::saveToFile(const std::string& filename)
    {
        return pImpl->saveToFile(filename);
    }

    TXSheet* TXWorkbook::addSheet(const std::string& name)
    {
        if (!pImpl)
        {
            // 通常 pImpl 不应该为 nullptr，除非构造失败
            return nullptr;
        }
        auto sheet_obj = std::make_unique<TXSheet>(name, this);

        return pImpl->storeSheet(std::move(sheet_obj));
    }

    TXSheet* TXWorkbook::addSheet(std::unique_ptr<TXSheet> sheet) const
    {
        if (!pImpl || !sheet) return nullptr;
        return pImpl->storeSheet(std::move(sheet));
    }

    
    TXSheet* TXWorkbook::getSheet(u64 index) const
    {
        return pImpl->getSheet(index);
    }

    TXSheet* TXWorkbook::getSheet(const std::string& name)
    {
        return pImpl->getSheet(name);
    }

    std::size_t TXWorkbook::getSheetCount() const
    {
        return pImpl->getSheetCount();
    }

    std::vector<std::string> TXWorkbook::getSheetNames() const
    {
        return pImpl->getSheetNames();
    }

    bool TXWorkbook::setActiveSheet(const std::string& name)
    {
        return pImpl->setActiveSheet(name);
    }

    bool TXWorkbook::removeSheet(const std::string& name)
    {
        return pImpl->removeSheet(name);
    }

    u32 TXWorkbook::registerOrGetStyleFId(const TXCellStyle& style)
    {
        return pImpl->getStyleManager()->registerCellStyleXF(style);
    }

    bool TXWorkbook::hasSheet(const std::string& name) const
    {
        return pImpl->hasSheet(name);
    }

    bool TXWorkbook::renameSheet(const std::string& oldName, const std::string& newName)
    {
        return pImpl->renameSheet(oldName, newName);
    }

    TXSheet* TXWorkbook::getActiveSheet()
    {
        return pImpl->getActiveSheet();
    }

    const std::string& TXWorkbook::getLastError() const
    {
        return pImpl->getLastError();
    }

    ComponentManager& TXWorkbook::getComponentManager()
    {
        return pImpl->getComponentManager();
    }

    const ComponentManager& TXWorkbook::getComponentManager() const
    {
        return pImpl->getComponentManager();
    }

    void TXWorkbook::setAutoComponentDetection(bool enable)
    {
        pImpl->setAutoComponentDetection(enable);
    }

    void TXWorkbook::registerComponent(ExcelComponent component)
    {
        pImpl->registerComponent(component);
    }

    void TXWorkbook::clear()
    {
        pImpl->clear();
    }

    bool TXWorkbook::isEmpty() const
    {
        return pImpl->getSheetCount() == 0;
    }

    std::vector<std::unique_ptr<TXSheet>>& TXWorkbook::getSheets()
    {
        return pImpl->getSheets();
    }

    const std::vector<std::unique_ptr<TXSheet>>& TXWorkbook::getSheets() const
    {
        return pImpl->getSheets();
    }

    TXStyleManager& TXWorkbook::getStyleManager()
    {
        return pImpl->getStyleManagerRef();
    }

    const TXStyleManager& TXWorkbook::getStyleManager() const
    {
        return pImpl->getStyleManagerRef();
    }
} // namespace TinaXlsx
