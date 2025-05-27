#include "TinaXlsx/TXWorkbook.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXZipHandler.hpp"
#include "TinaXlsx/TXXmlHandler.hpp"
#include <vector>
#include <unordered_map>
#include <algorithm>

namespace TinaXlsx {

class TXWorkbook::Impl {
public:
    Impl() : active_sheet_index_(0), last_error_("") {}

    ~Impl() = default;

    bool loadFromFile(const std::string& filename) {
        // 清空现有数据
        clear();
        
        TXZipHandler zip;
        if (!zip.open(filename, TXZipHandler::OpenMode::Read)) {
            last_error_ = "Failed to open XLSX file: " + zip.getLastError();
            return false;
        }

        // 检查是否是有效的XLSX文件
        if (!zip.hasFile("[Content_Types].xml")) {
            last_error_ = "Invalid XLSX file: missing [Content_Types].xml";
            return false;
        }

        // 读取工作簿结构（简化版）
        if (zip.hasFile("xl/workbook.xml")) {
            std::string workbook_xml = zip.readFileToString("xl/workbook.xml");
            if (workbook_xml.empty()) {
                last_error_ = "Failed to read workbook.xml";
                return false;
            }

            // 解析工作簿XML获取工作表信息
            if (!parseWorkbook(workbook_xml)) {
                return false;
            }
        }

        // 如果没有工作表，创建一个默认的
        if (sheets_.empty()) {
            addSheet("Sheet1");
        }

        last_error_.clear();
        return true;
    }

    bool saveToFile(const std::string& filename) {
        if (sheets_.empty()) {
            last_error_ = "No sheets to save";
            return false;
        }

        TXZipHandler zip;
        if (!zip.open(filename, TXZipHandler::OpenMode::Write)) {
            last_error_ = "Failed to create XLSX file: " + zip.getLastError();
            return false;
        }

        // 写入必要的XLSX结构文件
        if (!writeXlsxStructure(zip)) {
            return false;
        }

        zip.close();
        last_error_.clear();
        return true;
    }

    TXSheet* addSheet(const std::string& name) {
        // 检查名称是否已存在
        if (hasSheet(name)) {
            last_error_ = "Sheet name already exists: " + name;
            return nullptr;
        }

        sheets_.emplace_back(std::make_unique<TXSheet>(name));
        last_error_.clear();
        return sheets_.back().get();
    }

    TXSheet* getSheet(const std::string& name) {
        for (auto& sheet : sheets_) {
            if (sheet->getName() == name) {
                return sheet.get();
            }
        }
        return nullptr;
    }

    TXSheet* getSheet(std::size_t index) {
        if (index < sheets_.size()) {
            return sheets_[index].get();
        }
        return nullptr;
    }

    bool removeSheet(const std::string& name) {
        auto it = std::find_if(sheets_.begin(), sheets_.end(),
            [&name](const std::unique_ptr<TXSheet>& sheet) {
                return sheet->getName() == name;
            });

        if (it != sheets_.end()) {
            sheets_.erase(it);
            // 调整活动工作表索引
            if (active_sheet_index_ >= sheets_.size() && !sheets_.empty()) {
                active_sheet_index_ = sheets_.size() - 1;
            }
            last_error_.clear();
            return true;
        }

        last_error_ = "Sheet not found: " + name;
        return false;
    }

    std::size_t getSheetCount() const {
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

    bool hasSheet(const std::string& name) const {
        return getSheet(name) != nullptr;
    }

    bool renameSheet(const std::string& oldName, const std::string& newName) {
        if (hasSheet(newName)) {
            last_error_ = "New sheet name already exists: " + newName;
            return false;
        }

        TXSheet* sheet = getSheet(oldName);
        if (!sheet) {
            last_error_ = "Sheet not found: " + oldName;
            return false;
        }

        sheet->setName(newName);
        last_error_.clear();
        return true;
    }

    TXSheet* getActiveSheet() {
        if (active_sheet_index_ < sheets_.size()) {
            return sheets_[active_sheet_index_].get();
        }
        return nullptr;
    }

    bool setActiveSheet(const std::string& name) {
        for (std::size_t i = 0; i < sheets_.size(); ++i) {
            if (sheets_[i]->getName() == name) {
                active_sheet_index_ = i;
                last_error_.clear();
                return true;
            }
        }

        last_error_ = "Sheet not found: " + name;
        return false;
    }

    void clear() {
        sheets_.clear();
        active_sheet_index_ = 0;
        last_error_.clear();
    }

    bool isEmpty() const {
        return sheets_.empty();
    }

    const std::string& getLastError() const {
        return last_error_;
    }

private:
    TXSheet* getSheet(const std::string& name) const {
        for (const auto& sheet : sheets_) {
            if (sheet->getName() == name) {
                return sheet.get();
            }
        }
        return nullptr;
    }

    bool parseWorkbook(const std::string& workbook_xml) {
        TXXmlHandler xml;
        if (!xml.parseFromString(workbook_xml)) {
            last_error_ = "Failed to parse workbook XML: " + xml.getLastError();
            return false;
        }

        // 查找工作表信息（简化版，只提取名称）
        auto sheet_nodes = xml.findNodes("//sheet");
        for (const auto& node : sheet_nodes) {
            auto name_it = node.attributes.find("name");
            if (name_it != node.attributes.end()) {
                addSheet(name_it->second);
            }
        }

        return true;
    }

    bool writeXlsxStructure(TXZipHandler& zip) {
        // 写入Content Types
        std::string content_types = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>)";

        for (std::size_t i = 0; i < sheets_.size(); ++i) {
            content_types += R"(<Override PartName="/xl/worksheets/sheet)" + std::to_string(i + 1) + 
                           R"(.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>)";
        }

        content_types += "</Types>";

        if (!zip.writeFile("[Content_Types].xml", content_types)) {
            last_error_ = "Failed to write Content Types: " + zip.getLastError();
            return false;
        }

        // 写入主关系文件
        std::string main_rels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)";

        if (!zip.writeFile("_rels/.rels", main_rels)) {
            last_error_ = "Failed to write main relationships: " + zip.getLastError();
            return false;
        }

        // 写入工作簿文件
        if (!writeWorkbook(zip)) {
            return false;
        }

        // 写入工作表文件
        for (std::size_t i = 0; i < sheets_.size(); ++i) {
            if (!writeWorksheet(zip, i)) {
                return false;
            }
        }

        return true;
    }

    bool writeWorkbook(TXZipHandler& zip) {
        std::string workbook_xml = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>)";

        for (std::size_t i = 0; i < sheets_.size(); ++i) {
            workbook_xml += R"(<sheet name=")" + sheets_[i]->getName() + 
                          R"(" sheetId=")" + std::to_string(i + 1) + 
                          R"(" r:id="rId)" + std::to_string(i + 1) + R"("/>)";
        }

        workbook_xml += "</sheets></workbook>";

        if (!zip.writeFile("xl/workbook.xml", workbook_xml)) {
            last_error_ = "Failed to write workbook: " + zip.getLastError();
            return false;
        }

        // 写入工作簿关系文件
        std::string workbook_rels = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)";

        for (std::size_t i = 0; i < sheets_.size(); ++i) {
            workbook_rels += R"(<Relationship Id="rId)" + std::to_string(i + 1) + 
                           R"(" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet)" + 
                           std::to_string(i + 1) + R"(.xml"/>)";
        }

        workbook_rels += "</Relationships>";

        if (!zip.writeFile("xl/_rels/workbook.xml.rels", workbook_rels)) {
            last_error_ = "Failed to write workbook relationships: " + zip.getLastError();
            return false;
        }

        return true;
    }

    bool writeWorksheet(TXZipHandler& zip, std::size_t sheet_index) {
        if (sheet_index >= sheets_.size()) {
            last_error_ = "Invalid sheet index";
            return false;
        }

        const auto& sheet = sheets_[sheet_index];
        std::string worksheet_xml = R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>)";

        // 获取使用的范围
        auto used_range = sheet->getUsedRange();
        
        // 写入单元格数据（简化版）
        for (row_t row = used_range.getStart().getRow(); row <= used_range.getEnd().getRow(); ++row) {
            bool has_data = false;
            std::string row_xml = R"(<row r=")" + std::to_string(row.index()) + R"(">)";
            
            for (column_t col = used_range.getStart().getCol(); col <= used_range.getEnd().getCol(); ++col) {
                const TXCell* cell = sheet->getCell(row, col);
                if (cell && !cell->isEmpty()) {
                    has_data = true;
                    TXCoordinate coord(row, col);
                    std::string cell_ref = coord.toAddress();
                    std::string cell_value = cell->getStringValue();
                    
                    row_xml += R"(<c r=")" + cell_ref + R"(">)";
                    row_xml += "<v>" + cell_value + "</v>";
                    row_xml += "</c>";
                }
            }
            
            row_xml += "</row>";
            
            if (has_data) {
                worksheet_xml += row_xml;
            }
        }

        worksheet_xml += "</sheetData></worksheet>";

        std::string filename = "xl/worksheets/sheet" + std::to_string(sheet_index + 1) + ".xml";
        if (!zip.writeFile(filename, worksheet_xml)) {
            last_error_ = "Failed to write worksheet: " + zip.getLastError();
            return false;
        }

        return true;
    }

private:
    std::vector<std::unique_ptr<TXSheet>> sheets_;
    std::size_t active_sheet_index_;
    mutable std::string last_error_;
};

// TXWorkbook 实现
TXWorkbook::TXWorkbook() : pImpl(std::make_unique<Impl>()) {}

TXWorkbook::~TXWorkbook() = default;

TXWorkbook::TXWorkbook(TXWorkbook&& other) noexcept : pImpl(std::move(other.pImpl)) {}

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
}

TXSheet* TXWorkbook::getSheet(const std::string& name) {
    return pImpl->getSheet(name);
}

TXSheet* TXWorkbook::getSheet(std::size_t index) {
    return pImpl->getSheet(index);
}

bool TXWorkbook::removeSheet(const std::string& name) {
    return pImpl->removeSheet(name);
}

std::size_t TXWorkbook::getSheetCount() const {
    return pImpl->getSheetCount();
}

std::vector<std::string> TXWorkbook::getSheetNames() const {
    return pImpl->getSheetNames();
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

bool TXWorkbook::setActiveSheet(const std::string& name) {
    return pImpl->setActiveSheet(name);
}

const std::string& TXWorkbook::getLastError() const {
    return pImpl->getLastError();
}

void TXWorkbook::clear() {
    pImpl->clear();
}

bool TXWorkbook::isEmpty() const {
    return pImpl->isEmpty();
}

} // namespace TinaXlsx 