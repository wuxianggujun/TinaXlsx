#include "TinaXlsx/TXParallelProcessor.hpp"
#include <algorithm>
#include <cmath>

namespace TinaXlsx {

TXSmartParallelCellProcessor::TXSmartParallelCellProcessor(const ProcessorConfig& config)
    : config_(config) {
    threadPool_ = std::make_unique<TXLockFreeThreadPool>(
        TXLockFreeThreadPool::PoolConfig{config_.numThreads}
    );
}

TXSmartParallelCellProcessor::~TXSmartParallelCellProcessor() = default;

size_t TXSmartParallelCellProcessor::calculateOptimalBatchSize(size_t totalItems) const {
    if (!config_.enableAdaptiveBatching) {
        return config_.minBatchSize;
    }

    // 自适应批次大小计算
    size_t numThreads = config_.numThreads;
    size_t baseBatchSize = std::max(totalItems / (numThreads * 4), size_t(1));

    // 限制在合理范围内
    return std::clamp(baseBatchSize, config_.minBatchSize, config_.maxBatchSize);
}

std::vector<std::pair<TXCoordinate, cell_value_t>> 
TXSmartParallelCellProcessor::sortForCacheEfficiency(
    const std::vector<std::pair<TXCoordinate, cell_value_t>>& values) const {
    
    auto sorted = values;
    
    // 按行优先，列次要的顺序排序，提高缓存效率
    std::sort(sorted.begin(), sorted.end(), 
        [](const auto& a, const auto& b) {
            if (a.first.getRow().index() != b.first.getRow().index()) {
                return a.first.getRow().index() < b.first.getRow().index();
            }
            return a.first.getCol().index() < b.first.getCol().index();
        });
    
    return sorted;
}

std::vector<std::vector<std::pair<TXCoordinate, cell_value_t>>> 
TXSmartParallelCellProcessor::createBalancedBatches(
    const std::vector<std::pair<TXCoordinate, cell_value_t>>& values, 
    size_t batchSize) const {
    
    std::vector<std::vector<std::pair<TXCoordinate, cell_value_t>>> batches;
    
    for (size_t i = 0; i < values.size(); i += batchSize) {
        size_t endIdx = std::min(i + batchSize, values.size());
        batches.emplace_back(values.begin() + i, values.begin() + endIdx);
    }
    
    return batches;
}

void TXSmartParallelCellProcessor::updateAdaptiveParameters(size_t totalItems, size_t processedItems) {
    // 更新自适应参数
    if (config_.enableAdaptiveBatching && processedItems > 0) {
        double efficiency = static_cast<double>(processedItems) / totalItems;

        // 根据效率调整批次大小范围
        if (efficiency > 0.95) {
            // 效率很高，可以增加最大批次大小
            config_.maxBatchSize = std::min(static_cast<size_t>(config_.maxBatchSize * 1.1), size_t(20000));
        } else if (efficiency < 0.8) {
            // 效率较低，减少最大批次大小
            config_.maxBatchSize = std::max(static_cast<size_t>(config_.maxBatchSize * 0.9), config_.minBatchSize * 2);
        }
    }
}

} // namespace TinaXlsx
