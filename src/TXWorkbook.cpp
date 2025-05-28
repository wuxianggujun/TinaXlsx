//
// @file TXWorkbook.cpp - 新架构实现
//

#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXWorksheetWriter.hpp"
#include "TinaXlsx/TXWorksheetReader.hpp"
#include "TinaXlsx/TXComponentManager.hpp"
#include <vector>
#include <unordered_map>
#include <unordered_set>
#include <algorithm>
#include <iostream>
#include <regex>

namespace TinaXlsx {

class TXWorkbook::Impl {
public:
    Impl() : active_sheet_index_(0), last_error_(""), auto_component_detection_(true) {
        component_manager_.registerComponent(ExcelComponent::BasicWorkbook);
    }

    ~Impl() = default;

    bool loadFromFile(const std::string& filename) {
        // 清空现有数据
        clear();
        
        // 使用工作表读取器读取每个工作表
        TXWorksheetReader reader;
        
        // 首先我们需要确定有多少个工作表
        // 这需要读取 workbook.xml，但为了简化，我们先假设只有一个工作表
        
        // 创建一个新工作表并读取
        auto sheet = addSheet("Sheet1");
        if (!sheet) {
            last_error_ = "Failed to create sheet";
            return false;
        }
        
        if (!reader.readWorksheetFromFile(filename, sheet, 1)) {
            last_error_ = "Failed to read worksheet: " + reader.getLastError();
            return false;
        }
        
        return true;
    }

    bool saveToFile(const std::string& filename) {
        if (sheets_.empty()) {
            last_error_ = "No sheets to save";
            return false;
        }
        
        // 使用工作表写入器保存每个工作表
        TXWorksheetWriter writer;
        
        for (std::size_t i = 0; i < sheets_.size(); ++i) {
            bool appendMode = (i > 0); // 第一个工作表创建文件，后续追加
            
            if (!writer.writeWorksheetToFile(filename, sheets_[i].get(), i + 1, appendMode)) {
                last_error_ = "Failed to write worksheet " + std::to_string(i + 1) + 
                            ": " + writer.getLastError();
                return false;
            }
        }
        
        return true;
    }    TXSheet* addSheet(const std::string& name) {
        auto sheet = std::make_unique<TXSheet>(name);
        TXSheet* sheetPtr = sheet.get();
        sheets_.push_back(std::move(sheet));
        
        // 更新组件使用记录
        if (auto_component_detection_) {
            component_manager_.registerComponent(ExcelComponent::BasicWorkbook);
        }
        
        return sheetPtr;
    }

    bool removeSheet(const std::string& name) {
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
        }
        
        return true;
    }

    bool hasSheet(const std::string& name) const {
        return std::find_if(sheets_.begin(), sheets_.end(),
            [&name](const std::unique_ptr<TXSheet>& sheet) {
                return sheet->getName() == name;
            }) != sheets_.end();
    }

    bool renameSheet(const std::string& oldName, const std::string& newName) {
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

    TXSheet* getSheet(std::size_t index) {
        if (index < sheets_.size()) {
            return sheets_[index].get();
        }
        return nullptr;
    }

    TXSheet* getSheet(const std::string& name) {
        auto it = std::find_if(sheets_.begin(), sheets_.end(),
            [&name](const std::unique_ptr<TXSheet>& sheet) {
                return sheet->getName() == name;
            });
        
        return (it != sheets_.end()) ? it->get() : nullptr;
    }

    const TXSheet* getSheet(std::size_t index) const {
        if (index < sheets_.size()) {
            return sheets_[index].get();
        }
        return nullptr;
    }

    const TXSheet* getSheet(const std::string& name) const {
        auto it = std::find_if(sheets_.begin(), sheets_.end(),
            [&name](const std::unique_ptr<TXSheet>& sheet) {
                return sheet->getName() == name;
            });
        
        return (it != sheets_.end()) ? it->get() : nullptr;
    }    std::size_t getSheetCount() const {
        return sheets_.size();
    }

    std::vector<std::string> getSheetNames() const {
        std::vector<std::string> names;
        names.reserve(sheets_.size());
        
        for (const auto& sheet : sheets_) {
            names.push_back(sheet->getName());
        }
        
        return names;
    }

    bool setActiveSheet(std::size_t index) {
        if (index < sheets_.size()) {
            active_sheet_index_ = index;
            return true;
        }
        last_error_ = "Sheet index out of range";
        return false;
    }

    bool setActiveSheet(const std::string& name) {
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

    std::size_t getActiveSheetIndex() const {
        return active_sheet_index_;
    }

    TXSheet* getActiveSheet() {
        return getSheet(active_sheet_index_);
    }

    const TXSheet* getActiveSheet() const {
        return getSheet(active_sheet_index_);
    }

    const std::string& getLastError() const {
        return last_error_;
    }

    void clear() {
        sheets_.clear();
        active_sheet_index_ = 0;
        last_error_.clear();
    }    ComponentManager& getComponentManager() {
        return component_manager_;
    }

    const ComponentManager& getComponentManager() const {
        return component_manager_;
    }

    void setAutoComponentDetection(bool enable) {
        auto_component_detection_ = enable;
    }

    void registerComponent(ExcelComponent component) {
        component_manager_.registerComponent(component);
    }

private:
    std::vector<std::unique_ptr<TXSheet>> sheets_;
    std::size_t active_sheet_index_;
    mutable std::string last_error_;
    ComponentManager component_manager_;
    bool auto_component_detection_;
};

// TXWorkbook 实现

TXWorkbook::TXWorkbook() : pImpl(std::make_unique<Impl>()) {}

TXWorkbook::~TXWorkbook() = default;

TXWorkbook::TXWorkbook(TXWorkbook&& other) noexcept
    : pImpl(std::move(other.pImpl)) {}

TXWorkbook& TXWorkbook::operator=(TXWorkbook&& other) noexcept {
    if (this != &other) {
        pImpl = std::move(other.pImpl);
    }
    return *this;
}

bool TXWorkbook::loadFromFile(const std::string& filename) {
    return pImpl->loadFromFile(filename);
}

bool TXWorkbook::saveToFile(const std::string& filename) {
    return pImpl->saveToFile(filename);
}

TXSheet* TXWorkbook::addSheet(const std::string& name) {
    return pImpl->addSheet(name);
}TXSheet* TXWorkbook::getSheet(std::size_t index) {
    return pImpl->getSheet(index);
}

TXSheet* TXWorkbook::getSheet(const std::string& name) {
    return pImpl->getSheet(name);
}



std::size_t TXWorkbook::getSheetCount() const {
    return pImpl->getSheetCount();
}

std::vector<std::string> TXWorkbook::getSheetNames() const {
    return pImpl->getSheetNames();
}

bool TXWorkbook::setActiveSheet(const std::string& name) {
    return pImpl->setActiveSheet(name);
}

bool TXWorkbook::removeSheet(const std::string& name) {
    return pImpl->removeSheet(name);
}

bool TXWorkbook::hasSheet(const std::string& name) const {
    return pImpl->hasSheet(name);
}

bool TXWorkbook::renameSheet(const std::string& oldName, const std::string& newName) {
    return pImpl->renameSheet(oldName, newName);
}

TXSheet* TXWorkbook::getActiveSheet() {
    return pImpl->getActiveSheet();
}



const std::string& TXWorkbook::getLastError() const {
    return pImpl->getLastError();
}

ComponentManager& TXWorkbook::getComponentManager() {
    return pImpl->getComponentManager();
}

const ComponentManager& TXWorkbook::getComponentManager() const {
    return pImpl->getComponentManager();
}

void TXWorkbook::setAutoComponentDetection(bool enable) {
    pImpl->setAutoComponentDetection(enable);
}

void TXWorkbook::registerComponent(ExcelComponent component) {
    pImpl->registerComponent(component);
}

void TXWorkbook::clear() {
    pImpl->clear();
}

bool TXWorkbook::isEmpty() const {
    return pImpl->getSheetCount() == 0;
}

} // namespace TinaXlsx