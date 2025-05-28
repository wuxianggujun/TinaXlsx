#include "TinaXlsx/TXBatchFormat.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXRange.hpp"
#include <memory>

namespace TinaXlsx {

// ==================== TXBatchFormatApplicator实现 ====================

TXBatchFormatApplicator::TXBatchFormatApplicator() = default;
TXBatchFormatApplicator::~TXBatchFormatApplicator() = default;

size_t TXBatchFormatApplicator::applyStyleToRange(TXSheet& sheet, const TXRange& range, const TXCellStyle& style, const FormatApplyOptions& options) {
    size_t count = 0;
    
    for (auto row = range.getStart().getRow(); row <= range.getEnd().getRow(); ++row) {
        for (auto col = range.getStart().getCol(); col <= range.getEnd().getCol(); ++col) {
            try {
                auto* cell = sheet.getCell(row, col);
                if (cell) {
                    applySingleCellStyle(*cell, style, options);
                    count++;
                }
            } catch (...) {
                // 跳过错误的单元格
            }
        }
    }
    
    return count;
}

size_t TXBatchFormatApplicator::applyStyleToRanges(TXSheet& sheet, const std::vector<TXRange>& ranges, const TXCellStyle& style, const FormatApplyOptions& options) {
    size_t totalCount = 0;
    
    for (const auto& range : ranges) {
        totalCount += applyStyleToRange(sheet, range, style, options);
    }
    
    return totalCount;
}

size_t TXBatchFormatApplicator::applyStyleToRows(TXSheet& sheet, row_t startRow, row_t endRow, const TXCellStyle& style, const FormatApplyOptions& options) {
    size_t count = 0;
    
    // 简化实现：只处理已存在的单元格
    for (auto row = startRow; row <= endRow; ++row) {
        // 假设最大列数为1000，实际应该根据sheet的实际列数
        for (column_t col(1); col <= column_t(1000); ++col) {
            try {
                auto* cell = sheet.getCell(row, col);
                if (cell && !cell->isEmpty()) {
                    applySingleCellStyle(*cell, style, options);
                    count++;
                }
            } catch (...) {
                // 跳过不存在的单元格
            }
        }
    }
    
    return count;
}

size_t TXBatchFormatApplicator::applyStyleToColumns(TXSheet& sheet, column_t startCol, column_t endCol, const TXCellStyle& style, const FormatApplyOptions& options) {
    size_t count = 0;
    
    // 简化实现：只处理已存在的单元格
    for (auto col = startCol; col <= endCol; ++col) {
        // 假设最大行数为10000，实际应该根据sheet的实际行数
        for (row_t row(1); row <= row_t(10000); ++row) {
            try {
                auto* cell = sheet.getCell(row, col);
                if (cell && !cell->isEmpty()) {
                    applySingleCellStyle(*cell, style, options);
                    count++;
                }
            } catch (...) {
                // 跳过不存在的单元格
            }
        }
    }
    
    return count;
}

void TXBatchFormatApplicator::applySingleCellStyle(TXCell& cell, const TXCellStyle& /*style*/, const FormatApplyOptions& options) {
    // 简化实现：直接设置样式索引
    // 在实际实现中，应该从样式管理器获取或创建样式
    
    switch (options.mode) {
        case FormatApplyMode::Replace:
            // 完全替换样式
            cell.setStyleIndex(1); // 临时使用固定索引，实际应该管理样式
            break;
            
        case FormatApplyMode::Merge:
            // 合并样式：需要样式管理器支持
            cell.setStyleIndex(1); // 临时实现
            break;
            
        case FormatApplyMode::Overlay:
            // 覆盖应用
            cell.setStyleIndex(1); // 临时实现
            break;
    }
}

// ==================== TXFormatBatchTask实现 ====================

class TXFormatBatchTask::Impl {
public:
    TaskType type_;
    TXRange targetRange_;
    FormatApplyOptions options_;
    TXCellStyle style_;
    
    Impl(TaskType type) : type_(type) {}
};

TXFormatBatchTask::TXFormatBatchTask(TaskType type) : pImpl(std::make_unique<Impl>(type)) {}

TXFormatBatchTask::~TXFormatBatchTask() = default;

void TXFormatBatchTask::setTargetRange(const TXRange& range) {
    pImpl->targetRange_ = range;
}

void TXFormatBatchTask::setOptions(const FormatApplyOptions& options) {
    pImpl->options_ = options;
}

void TXFormatBatchTask::setStyle(const TXCellStyle& style) {
    pImpl->style_ = style;
}

void TXFormatBatchTask::setTemplate(const TXStyleTemplate& /*styleTemplate*/) {
    // 临时实现
}

bool TXFormatBatchTask::execute(TXSheet& sheet) {
    try {
        TXBatchFormatApplicator applicator;
        applicator.applyStyleToRange(sheet, pImpl->targetRange_, pImpl->style_, pImpl->options_);
        return true;
    } catch (...) {
        return false;
    }
}

TXFormatBatchTask::TaskType TXFormatBatchTask::getType() const {
    return pImpl->type_;
}

// ==================== TXBatchFormatTaskManager实现 ====================

class TXBatchFormatTaskManager::Impl {
public:
    std::vector<std::unique_ptr<TXFormatBatchTask>> tasks_;
    size_t currentTaskIndex_;
    
    Impl() : currentTaskIndex_(0) {}
};

TXBatchFormatTaskManager::TXBatchFormatTaskManager() : pImpl(std::make_unique<Impl>()) {}

TXBatchFormatTaskManager::~TXBatchFormatTaskManager() = default;

void TXBatchFormatTaskManager::addTask(std::unique_ptr<TXFormatBatchTask> task) {
    pImpl->tasks_.push_back(std::move(task));
}

size_t TXBatchFormatTaskManager::executeAllTasks(TXSheet& sheet) {
    size_t executed = 0;
    for (auto& task : pImpl->tasks_) {
        if (task && task->execute(sheet)) {
            executed++;
        }
    }
    return executed;
}

void TXBatchFormatTaskManager::clearTasks() {
    pImpl->tasks_.clear();
    pImpl->currentTaskIndex_ = 0;
}

size_t TXBatchFormatTaskManager::getTaskCount() const {
    return pImpl->tasks_.size();
}

} // namespace TinaXlsx