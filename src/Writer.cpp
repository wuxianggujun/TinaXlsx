/**
 * @file Writer.cpp
 * @brief 高性能Excel写入器实现
 */

#include "TinaXlsx/Writer.hpp"
#include "TinaXlsx/Worksheet.hpp"
#include <xlsxwriter.h>
#include <memory>
#include <unordered_map>
#include <algorithm>
#include <filesystem>
#include <locale>
#include <codecvt>

namespace TinaXlsx {

struct Writer::Impl {
    lxw_workbook* workbook = nullptr;
    std::string filePath;
    std::string tempFilePath; // 用于处理非ASCII路径
    std::unordered_map<std::string, std::shared_ptr<Worksheet>> worksheets;
    std::vector<std::string> worksheetNames;
    bool closed = false;
    
    ~Impl() {
        close();
    }
    
    void close() {
        if (workbook && !closed) {
            workbook_close(workbook);
            closed = true;
        }
        workbook = nullptr;
        worksheets.clear();
        worksheetNames.clear();
    }
};

Writer::Writer(const std::string& filePath) 
    : pImpl_(std::make_unique<Impl>()) {
    pImpl_->filePath = filePath;
    
    // 直接使用UTF-8编码创建workbook，与原始ExcelTemplate保持一致
    pImpl_->workbook = workbook_new(filePath.c_str());
    
    if (!pImpl_->workbook) {
        throw FileException("无法创建Excel文件: " + filePath);
    }
}

Writer::Writer(Writer&& other) noexcept 
    : pImpl_(std::move(other.pImpl_)) {
}

Writer& Writer::operator=(Writer&& other) noexcept {
    if (this != &other) {
        pImpl_ = std::move(other.pImpl_);
    }
    return *this;
}

Writer::~Writer() = default;

std::shared_ptr<Worksheet> Writer::createWorksheet(const std::string& sheetName) {
    if (!pImpl_->workbook || pImpl_->closed) {
        throw WorksheetException("工作簿已关闭");
    }
    
    // 生成工作表名称
    std::string finalName = sheetName;
    if (finalName.empty()) {
        finalName = "Sheet" + std::to_string(pImpl_->worksheetNames.size() + 1);
    }
    
    // 检查名称是否已存在
    if (pImpl_->worksheets.find(finalName) != pImpl_->worksheets.end()) {
        throw WorksheetException("工作表名称已存在: " + finalName);
    }
    
    // 创建libxlsxwriter工作表
    lxw_worksheet* ws = workbook_add_worksheet(pImpl_->workbook, finalName.c_str());
    if (!ws) {
        throw WorksheetException("无法创建工作表: " + finalName);
    }
    
    // 创建Worksheet对象
    auto worksheet = std::make_shared<Worksheet>(ws, pImpl_->workbook, finalName);
    pImpl_->worksheets[finalName] = worksheet;
    pImpl_->worksheetNames.push_back(finalName);
    
    return worksheet;
}

std::shared_ptr<Worksheet> Writer::getWorksheet(const std::string& sheetName) {
    auto it = pImpl_->worksheets.find(sheetName);
    if (it != pImpl_->worksheets.end()) {
        return it->second;
    }
    return nullptr;
}

std::shared_ptr<Worksheet> Writer::getWorksheet(SheetIndex index) {
    if (index >= pImpl_->worksheetNames.size()) {
        return nullptr;
    }
    
    const std::string& name = pImpl_->worksheetNames[index];
    return getWorksheet(name);
}

std::vector<std::string> Writer::getWorksheetNames() const {
    return pImpl_->worksheetNames;
}

size_t Writer::getWorksheetCount() const {
    return pImpl_->worksheetNames.size();
}

std::unique_ptr<Format> Writer::createFormat() {
    if (!pImpl_->workbook || pImpl_->closed) {
        throw FormatException("工作簿已关闭");
    }
    
    return std::make_unique<Format>(pImpl_->workbook);
}

FormatBuilder Writer::getFormatBuilder() {
    if (!pImpl_->workbook || pImpl_->closed) {
        throw FormatException("工作簿已关闭");
    }
    
    return FormatBuilder(pImpl_->workbook);
}

void Writer::setDocumentProperties(
    const std::string& title,
    const std::string& subject,
    const std::string& author,
    const std::string& manager,
    const std::string& company,
    const std::string& category,
    const std::string& keywords,
    const std::string& comments) {
    
    if (!pImpl_->workbook || pImpl_->closed) {
        throw Exception("工作簿已关闭");
    }
    
    lxw_doc_properties properties = {};
    
    if (!title.empty()) properties.title = title.c_str();
    if (!subject.empty()) properties.subject = subject.c_str();
    if (!author.empty()) properties.author = author.c_str();
    if (!manager.empty()) properties.manager = manager.c_str();
    if (!company.empty()) properties.company = company.c_str();
    if (!category.empty()) properties.category = category.c_str();
    if (!keywords.empty()) properties.keywords = keywords.c_str();
    if (!comments.empty()) properties.comments = comments.c_str();
    
    workbook_set_properties(pImpl_->workbook, &properties);
}

void Writer::setCustomProperty(const std::string& name, const std::string& value) {
    if (!pImpl_->workbook || pImpl_->closed) {
        throw Exception("工作簿已关闭");
    }
    
    workbook_set_custom_property_string(pImpl_->workbook, name.c_str(), value.c_str());
}

void Writer::setCustomProperty(const std::string& name, double value) {
    if (!pImpl_->workbook || pImpl_->closed) {
        throw Exception("工作簿已关闭");
    }
    
    workbook_set_custom_property_number(pImpl_->workbook, name.c_str(), value);
}

void Writer::setCustomProperty(const std::string& name, bool value) {
    if (!pImpl_->workbook || pImpl_->closed) {
        throw Exception("工作簿已关闭");
    }
    
    workbook_set_custom_property_boolean(pImpl_->workbook, name.c_str(), value);
}

bool Writer::save() {
    if (!pImpl_->workbook) {
        return false;
    }
    
    if (pImpl_->closed) {
        return true; // 已经保存过了
    }
    
    // 直接关闭workbook，与原始ExcelTemplate保持一致
    lxw_error error = workbook_close(pImpl_->workbook);
    pImpl_->closed = true;
    
    return error == LXW_NO_ERROR;
}

bool Writer::close() {
    return save();
}

lxw_workbook* Writer::getInternalWorkbook() const {
    return pImpl_->workbook;
}

bool Writer::isClosed() const {
    return pImpl_->closed;
}

void Writer::setDefaultDateFormat(const std::string& format) {
    if (!pImpl_->workbook || pImpl_->closed) {
        throw Exception("工作簿已关闭");
    }
    
    // libxlsxwriter没有workbook_set_default_date_format函数
    // 可以通过创建一个默认的日期格式来实现类似功能
    // 这里暂时注释掉，用户需要手动为日期单元格设置格式
    // workbook_set_default_date_format(pImpl_->workbook, format.c_str());
}

void Writer::defineNamedRange(const std::string& name, const std::string& sheetName, const CellRange& range) {
    if (!pImpl_->workbook || pImpl_->closed) {
        throw Exception("工作簿已关闭");
    }
    
    std::string rangeStr = sheetName + "!" + 
                          cellPositionToString(range.start) + ":" + 
                          cellPositionToString(range.end);
    
    workbook_define_name(pImpl_->workbook, name.c_str(), rangeStr.c_str());
}

} // namespace TinaXlsx 