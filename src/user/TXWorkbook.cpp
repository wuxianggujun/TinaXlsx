//
// @file TXWorkbook.cpp
// @brief 🚀 用户层工作簿类实现
//

#include <TinaXlsx/user/TXWorkbook.hpp>
#include <TinaXlsx/io/TXExcelIO.hpp>
#include <TinaXlsx/TXHighPerformanceLogger.hpp>
#include <sstream>
#include <algorithm>

namespace TinaXlsx {

// ==================== 静态工厂方法 ====================

std::unique_ptr<TXWorkbook> TXWorkbook::create(const std::string& name) {
    auto workbook = std::make_unique<TXWorkbook>(name);

    // 创建默认工作表
    workbook->addSheet("Sheet1");

    // 新创建的工作簿标记为已保存状态
    workbook->markAsSaved();

    TX_LOG_INFO("创建新工作簿: {}", name);
    return workbook;
}

TXResult<std::unique_ptr<TXWorkbook>> TXWorkbook::load(const std::string& file_path) {
    TX_LOG_INFO("加载工作簿文件: {}", file_path);
    return TXExcelIO::loadFromFile(file_path);
}

TXResult<std::unique_ptr<TXWorkbook>> TXWorkbook::loadFromMemory(const void* data, size_t size) {
    TX_LOG_INFO("从内存加载工作簿: {} 字节", size);
    return TXExcelIO::loadFromMemory(data, size);
}

// ==================== 构造和析构 ====================

TXWorkbook::TXWorkbook(const std::string& name)
    : name_(name)
    , sheets_(GlobalUnifiedMemoryManager::getInstance())  // 初始化TXVector
    , active_sheet_index_(-1)
    , has_unsaved_changes_(false)
    , memory_manager_(GlobalUnifiedMemoryManager::getInstance())
    , string_pool_(TXGlobalStringPool::instance()) {

    TX_LOG_DEBUG("TXWorkbook '{}' 构造完成", name);
}

TXWorkbook::~TXWorkbook() {
    TX_LOG_DEBUG("TXWorkbook '{}' 析构", name_);
}

TXWorkbook::TXWorkbook(TXWorkbook&& other) noexcept
    : name_(std::move(other.name_))
    , sheets_(std::move(other.sheets_))
    , sheet_map_(std::move(other.sheet_map_))
    , active_sheet_index_(other.active_sheet_index_)
    , file_path_(std::move(other.file_path_))
    , has_unsaved_changes_(other.has_unsaved_changes_)
    , memory_manager_(other.memory_manager_)
    , string_pool_(other.string_pool_) {

    other.active_sheet_index_ = -1;
    other.has_unsaved_changes_ = false;
}

TXWorkbook& TXWorkbook::operator=(TXWorkbook&& other) noexcept {
    if (this != &other) {
        name_ = std::move(other.name_);
        sheets_ = std::move(other.sheets_);
        sheet_map_ = std::move(other.sheet_map_);
        active_sheet_index_ = other.active_sheet_index_;
        file_path_ = std::move(other.file_path_);
        has_unsaved_changes_ = other.has_unsaved_changes_;
        
        other.active_sheet_index_ = -1;
        other.has_unsaved_changes_ = false;
    }
    return *this;
}

// ==================== 工作表管理 ====================

TXSheet* TXWorkbook::addSheet(const std::string& name) {
    // 生成唯一名称
    std::string unique_name = generateUniqueSheetName(name);
    
    // 创建新工作表
    auto sheet = std::make_unique<TXSheet>(unique_name, memory_manager_, string_pool_);
    TXSheet* sheet_ptr = sheet.get();
    
    // 添加到列表
    size_t index = sheets_.size();
    sheets_.push_back(std::move(sheet));
    sheet_map_[unique_name] = index;
    
    // 如果是第一个工作表，设为活动工作表
    if (active_sheet_index_ == -1) {
        active_sheet_index_ = 0;
    }
    
    markAsModified();
    TX_LOG_DEBUG("添加工作表: {} (索引: {})", unique_name, index);
    
    return sheet_ptr;
}

TXResult<TXSheet*> TXWorkbook::insertSheet(size_t index, const std::string& name) {
    if (index > sheets_.size()) {
        return TXResult<TXSheet*>(TXError(TXErrorCode::InvalidArgument, "插入位置超出范围"));
    }
    
    // 生成唯一名称
    std::string unique_name = generateUniqueSheetName(name);
    
    // 创建新工作表
    auto sheet = std::make_unique<TXSheet>(unique_name, memory_manager_, string_pool_);
    TXSheet* sheet_ptr = sheet.get();
    
    // 插入到指定位置
    sheets_.insert(sheets_.begin() + index, std::move(sheet));
    
    // 更新映射和活动索引
    updateSheetMap();
    if (active_sheet_index_ >= static_cast<int>(index)) {
        active_sheet_index_++;
    }
    
    markAsModified();
    TX_LOG_DEBUG("插入工作表: {} (索引: {})", unique_name, index);
    
    return TXResult<TXSheet*>(sheet_ptr);
}

TXResult<void> TXWorkbook::removeSheet(size_t index) {
    if (!isValidIndex(index)) {
        return TXResult<void>(TXError(TXErrorCode::InvalidArgument, "工作表索引无效"));
    }
    
    if (sheets_.size() <= 1) {
        return TXResult<void>(TXError(TXErrorCode::InvalidOperation, "不能删除最后一个工作表"));
    }
    
    std::string sheet_name = sheets_[index]->getName();
    
    // 删除工作表
    sheets_.erase(sheets_.begin() + index);
    
    // 更新映射和活动索引
    updateSheetMap();
    adjustActiveSheetIndex();
    
    markAsModified();
    TX_LOG_DEBUG("删除工作表: {} (索引: {})", sheet_name, index);
    
    return TXResult<void>();
}

TXResult<void> TXWorkbook::removeSheet(const std::string& name) {
    int index = findSheetIndex(name);
    if (index == -1) {
        return TXResult<void>(TXError(TXErrorCode::SheetNotFound, "工作表不存在: " + name));
    }
    
    return removeSheet(static_cast<size_t>(index));
}

TXResult<void> TXWorkbook::renameSheet(size_t index, const std::string& new_name) {
    if (!isValidIndex(index)) {
        return TXResult<void>(TXError(TXErrorCode::InvalidArgument, "工作表索引无效"));
    }
    
    // 生成唯一名称
    std::string unique_name = generateUniqueSheetName(new_name);
    
    // 重命名工作表
    std::string old_name = sheets_[index]->getName();
    sheets_[index]->setName(unique_name);
    
    // 更新映射
    updateSheetMap();
    
    markAsModified();
    TX_LOG_DEBUG("重命名工作表: {} -> {}", old_name, unique_name);
    
    return TXResult<void>();
}

TXResult<void> TXWorkbook::renameSheet(const std::string& old_name, const std::string& new_name) {
    int index = findSheetIndex(old_name);
    if (index == -1) {
        return TXResult<void>(TXError(TXErrorCode::SheetNotFound, "工作表不存在: " + old_name));
    }
    
    return renameSheet(static_cast<size_t>(index), new_name);
}

TXResult<void> TXWorkbook::moveSheet(size_t from_index, size_t to_index) {
    if (!isValidIndex(from_index) || to_index >= sheets_.size()) {
        return TXResult<void>(TXError(TXErrorCode::InvalidArgument, "工作表索引无效"));
    }
    
    if (from_index == to_index) {
        return TXResult<void>(); // 无需移动
    }
    
    // 移动工作表
    auto sheet = std::move(sheets_[from_index]);
    sheets_.erase(sheets_.begin() + from_index);
    sheets_.insert(sheets_.begin() + to_index, std::move(sheet));
    
    // 更新映射和活动索引
    updateSheetMap();
    if (active_sheet_index_ == static_cast<int>(from_index)) {
        active_sheet_index_ = static_cast<int>(to_index);
    } else if (active_sheet_index_ > static_cast<int>(from_index) && 
               active_sheet_index_ <= static_cast<int>(to_index)) {
        active_sheet_index_--;
    } else if (active_sheet_index_ < static_cast<int>(from_index) && 
               active_sheet_index_ >= static_cast<int>(to_index)) {
        active_sheet_index_++;
    }
    
    markAsModified();
    TX_LOG_DEBUG("移动工作表: {} -> {}", from_index, to_index);
    
    return TXResult<void>();
}

// ==================== 工作表访问 ====================

TXSheet* TXWorkbook::getSheet(size_t index) {
    if (!isValidIndex(index)) {
        return nullptr;
    }
    return sheets_[index].get();
}

const TXSheet* TXWorkbook::getSheet(size_t index) const {
    if (!isValidIndex(index)) {
        return nullptr;
    }
    return sheets_[index].get();
}

TXSheet* TXWorkbook::getSheet(const std::string& name) {
    auto it = sheet_map_.find(name);
    if (it == sheet_map_.end()) {
        return nullptr;
    }
    return sheets_[it->second].get();
}

const TXSheet* TXWorkbook::getSheet(const std::string& name) const {
    auto it = sheet_map_.find(name);
    if (it == sheet_map_.end()) {
        return nullptr;
    }
    return sheets_[it->second].get();
}

TXSheet* TXWorkbook::getActiveSheet() {
    if (active_sheet_index_ == -1 || !isValidIndex(static_cast<size_t>(active_sheet_index_))) {
        return nullptr;
    }
    return sheets_[active_sheet_index_].get();
}

const TXSheet* TXWorkbook::getActiveSheet() const {
    if (active_sheet_index_ == -1 || !isValidIndex(static_cast<size_t>(active_sheet_index_))) {
        return nullptr;
    }
    return sheets_[active_sheet_index_].get();
}

TXResult<void> TXWorkbook::setActiveSheet(size_t index) {
    if (!isValidIndex(index)) {
        return TXResult<void>(TXError(TXErrorCode::InvalidArgument, "工作表索引无效"));
    }
    
    active_sheet_index_ = static_cast<int>(index);
    TX_LOG_DEBUG("设置活动工作表: {} ({})", index, sheets_[index]->getName());
    
    return TXResult<void>();
}

TXResult<void> TXWorkbook::setActiveSheet(const std::string& name) {
    int index = findSheetIndex(name);
    if (index == -1) {
        return TXResult<void>(TXError(TXErrorCode::SheetNotFound, "工作表不存在: " + name));
    }
    
    return setActiveSheet(static_cast<size_t>(index));
}

TXVector<std::string> TXWorkbook::getSheetNames() const {
    TXVector<std::string> names(memory_manager_);
    names.reserve(sheets_.size());

    for (const auto& sheet : sheets_) {
        names.push_back(sheet->getName());
    }

    return names;
}

bool TXWorkbook::hasSheet(const std::string& name) const {
    return sheet_map_.find(name) != sheet_map_.end();
}

int TXWorkbook::findSheetIndex(const std::string& name) const {
    auto it = sheet_map_.find(name);
    if (it == sheet_map_.end()) {
        return -1;
    }
    return static_cast<int>(it->second);
}

// ==================== 文件操作 ====================

TXResult<void> TXWorkbook::saveAs(const std::string& file_path) {
    TX_LOG_INFO("保存工作簿到: {}", file_path);

    auto result = TXExcelIO::saveToFile(this, file_path);
    if (result.isOk()) {
        file_path_ = file_path;
        markAsSaved();
        TX_LOG_INFO("工作簿保存成功: {}", file_path);
    }

    return result;
}

TXResult<void> TXWorkbook::save() {
    if (file_path_.empty()) {
        return TXResult<void>(TXError(TXErrorCode::InvalidOperation, "未指定文件路径，请使用saveAs"));
    }

    return saveAs(file_path_);
}

TXResult<TXVector<uint8_t>> TXWorkbook::exportToMemory() {
    TX_LOG_INFO("导出工作簿到内存");
    return TXExcelIO::saveToMemory(this);
}

// ==================== 性能优化 ====================

void TXWorkbook::reserve(size_t estimated_sheets) {
    sheets_.reserve(estimated_sheets);
    sheet_map_.reserve(estimated_sheets);
    TX_LOG_DEBUG("预分配工作表容量: {}", estimated_sheets);
}

void TXWorkbook::optimize() {
    for (auto& sheet : sheets_) {
        sheet->optimize();
    }
    TX_LOG_DEBUG("优化所有工作表完成");
}

size_t TXWorkbook::compress() {
    size_t total_compressed = 0;
    for (auto& sheet : sheets_) {
        total_compressed += sheet->compress();
    }
    TX_LOG_DEBUG("压缩所有工作表: {} 个单元格", total_compressed);
    return total_compressed;
}

void TXWorkbook::shrinkToFit() {
    for (auto& sheet : sheets_) {
        sheet->shrinkToFit();
    }
    TX_LOG_DEBUG("收缩所有工作表内存完成");
}

// ==================== 调试和诊断 ====================

std::string TXWorkbook::toString() const {
    std::ostringstream oss;
    oss << "TXWorkbook{";
    oss << "名称=\"" << name_ << "\"";
    oss << ", 工作表数=" << sheets_.size();
    oss << ", 活动工作表=" << (active_sheet_index_ >= 0 ?
                              sheets_[active_sheet_index_]->getName() : "无");
    oss << ", 文件路径=\"" << file_path_ << "\"";
    oss << ", 未保存更改=" << (has_unsaved_changes_ ? "是" : "否");
    oss << "}";
    return oss.str();
}

bool TXWorkbook::isValid() const {
    if (sheets_.empty()) {
        return false;
    }

    if (active_sheet_index_ < 0 ||
        active_sheet_index_ >= static_cast<int>(sheets_.size())) {
        return false;
    }

    // 检查所有工作表是否有效
    for (const auto& sheet : sheets_) {
        if (!sheet || !sheet->isValid()) {
            return false;
        }
    }

    return true;
}

std::string TXWorkbook::getPerformanceStats() const {
    std::ostringstream oss;
    oss << "TXWorkbook性能统计:\n";
    oss << "  工作簿名称: " << name_ << "\n";
    oss << "  工作表数量: " << sheets_.size() << "\n";
    oss << "  内存使用量: " << (getMemoryUsage() / 1024.0 / 1024.0) << "MB\n";
    oss << "  未保存更改: " << (has_unsaved_changes_ ? "是" : "否") << "\n";

    size_t total_cells = 0;
    for (const auto& sheet : sheets_) {
        total_cells += sheet->getCellCount();
    }
    oss << "  总单元格数: " << total_cells << "\n";

    oss << "\n各工作表统计:\n";
    for (size_t i = 0; i < sheets_.size(); ++i) {
        oss << "  [" << i << "] " << sheets_[i]->getName()
            << ": " << sheets_[i]->getCellCount() << " 个单元格\n";
    }

    return oss.str();
}

size_t TXWorkbook::getMemoryUsage() const {
    size_t total = sizeof(*this);

    // 工作表内存
    for (const auto& sheet : sheets_) {
        // 这里需要实现TXSheet的内存使用统计
        total += 1024; // 临时估算
    }

    // 字符串内存
    total += name_.capacity();
    total += file_path_.capacity();

    return total;
}

// ==================== 内部辅助方法 ====================

void TXWorkbook::handleError(const std::string& operation, const TXError& error) const {
    TX_LOG_WARN("TXWorkbook操作失败: {} - 工作簿={}, 错误={}",
                operation, name_, error.getMessage());
}

bool TXWorkbook::isValidIndex(size_t index) const {
    return index < sheets_.size();
}

std::string TXWorkbook::generateUniqueSheetName(const std::string& base_name) const {
    std::string name = base_name;
    int counter = 1;

    while (hasSheet(name)) {
        name = base_name + std::to_string(counter);
        counter++;
    }

    return name;
}

void TXWorkbook::updateSheetMap() {
    sheet_map_.clear();
    for (size_t i = 0; i < sheets_.size(); ++i) {
        sheet_map_[sheets_[i]->getName()] = i;
    }
}

void TXWorkbook::adjustActiveSheetIndex() {
    if (sheets_.empty()) {
        active_sheet_index_ = -1;
    } else if (active_sheet_index_ >= static_cast<int>(sheets_.size())) {
        active_sheet_index_ = static_cast<int>(sheets_.size()) - 1;
    }
}

} // namespace TinaXlsx
