//
// @file TXBatchPipeline.cpp
// @brief Batch processing pipeline implementation
//

#include "TinaXlsx/TXBatchPipeline.hpp"
#include "TinaXlsx/TXBatchPipelineStages.hpp"
#include "TinaXlsx/TXResult.hpp"
#include "TinaXlsx/TXError.hpp"
#include <algorithm>
#include <sstream>
#include <iomanip>

namespace TinaXlsx {

// ==================== TXBatchPipeline Implementation ====================

TXBatchPipeline::TXBatchPipeline(const PipelineConfig& config) : config_(config) {
    // Initialize memory manager
    TXUnifiedMemoryManager::Config memory_config;
    memory_config.memory_limit = config_.memory_limit_mb * 1024 * 1024;
    memory_config.enable_monitoring = config_.enable_performance_monitoring;
    
    memory_manager_ = std::make_unique<TXUnifiedMemoryManager>(memory_config);
    
    // Initialize default stages
    initializeDefaultStages();
}

TXBatchPipeline::~TXBatchPipeline() {
    if (state_.load() == PipelineState::RUNNING) {
        // Safe to ignore return value in destructor
        (void)stop();
    }
}

void TXBatchPipeline::initializeDefaultStages() {
    // Create default 4-stage pipeline stages
    auto stages = TXPipelineStageFactory::createDefaultStages(*memory_manager_);
    
    if (stages.size() >= 4) {
        stage1_ = std::move(stages[0]); // Data preprocessing
        stage2_ = std::move(stages[1]); // XML generation
        stage3_ = std::move(stages[2]); // Compression processing
        stage4_ = std::move(stages[3]); // Output writing
    }
}

TXResult<void> TXBatchPipeline::start() {
    PipelineState expected = PipelineState::STOPPED;
    if (!state_.compare_exchange_strong(expected, PipelineState::STARTING)) {
        return Err(TXErrorCode::OperationFailed, "Pipeline is not in STOPPED state");
    }
    
    try {
        // Start memory manager monitoring
        if (config_.enable_performance_monitoring && memory_manager_) {
            memory_manager_->startMonitoring();
        }
        
        // Start worker threads
        startWorkerThreads();
        
        // Record start time
        start_time_ = std::chrono::steady_clock::now();
        last_throughput_update_ = start_time_;
        
        state_.store(PipelineState::RUNNING);
        
        return Ok();
    } catch (const std::exception& e) {
        state_.store(PipelineState::ERROR);
        return Err(TXErrorCode::OperationFailed, "Failed to start pipeline: " + std::string(e.what()));
    }
}

TXResult<void> TXBatchPipeline::stop() {
    PipelineState expected = PipelineState::RUNNING;
    if (!state_.compare_exchange_strong(expected, PipelineState::STOPPING)) {
        expected = PipelineState::PAUSED;
        if (!state_.compare_exchange_strong(expected, PipelineState::STOPPING)) {
            return Err(TXErrorCode::OperationFailed, "Pipeline is not in RUNNING or PAUSED state");
        }
    }
    
    try {
        // Set stop flag
        should_stop_.store(true);
        
        // Notify all waiting threads
        stage1_cv_.notify_all();
        stage2_cv_.notify_all();
        stage3_cv_.notify_all();
        stage4_cv_.notify_all();
        
        // Stop worker threads
        stopWorkerThreads();
        
        // Stop memory manager monitoring
        if (memory_manager_) {
            memory_manager_->stopMonitoring();
        }
        
        state_.store(PipelineState::STOPPED);
        should_stop_.store(false);
        
        return Ok();
    } catch (const std::exception& e) {
        state_.store(PipelineState::ERROR);
        return Err(TXErrorCode::OperationFailed, "Failed to stop pipeline: " + std::string(e.what()));
    }
}

TXResult<void> TXBatchPipeline::pause() {
    PipelineState expected = PipelineState::RUNNING;
    if (!state_.compare_exchange_strong(expected, PipelineState::PAUSED)) {
        return Err(TXErrorCode::OperationFailed, "Pipeline is not in RUNNING state");
    }
    
    return Ok();
}

TXResult<void> TXBatchPipeline::resume() {
    PipelineState expected = PipelineState::PAUSED;
    if (!state_.compare_exchange_strong(expected, PipelineState::RUNNING)) {
        return Err(TXErrorCode::OperationFailed, "Pipeline is not in PAUSED state");
    }
    
    // Notify all waiting threads
    stage1_cv_.notify_all();
    stage2_cv_.notify_all();
    stage3_cv_.notify_all();
    stage4_cv_.notify_all();
    
    return Ok();
}

TXResult<size_t> TXBatchPipeline::submitBatch(std::unique_ptr<TXBatchData> batch) {
    if (state_.load() != PipelineState::RUNNING) {
        return Err<size_t>(TXErrorCode::Unknown, "Pipeline is not running");
    }
    
    if (!batch) {
        return Err<size_t>(TXErrorCode::Unknown, "Batch data is null");
    }
    
    // Check memory limit
    if (!checkMemoryLimit()) {
        return Err<size_t>(TXErrorCode::Unknown, "Memory limit exceeded");
    }
    
    size_t batch_id = batch->batch_id;
    
    // Add batch to first stage queue
    {
        std::lock_guard<std::mutex> lock(stage1_mutex_);
        stage1_queue_.push(std::move(batch));
    }
    
    // Notify first stage worker thread
    stage1_cv_.notify_one();
    
    // Update statistics
    {
        std::lock_guard<std::mutex> lock(stats_mutex_);
        stats_.batches_in_pipeline++;
    }
    
    return Ok(batch_id);
}

TXResult<std::vector<size_t>> TXBatchPipeline::submitBatches(std::vector<std::unique_ptr<TXBatchData>> batches) {
    std::vector<size_t> batch_ids;
    batch_ids.reserve(batches.size());
    
    for (auto& batch : batches) {
        auto result = submitBatch(std::move(batch));
        if (!result.isOk()) {
            return Err<std::vector<size_t>>(result.error());
        }
        batch_ids.push_back(result.value());
    }
    
    return Ok(std::move(batch_ids));
}

TXResult<void> TXBatchPipeline::waitForCompletion(std::chrono::milliseconds timeout) {
    auto start_time = std::chrono::steady_clock::now();
    
    while (true) {
        // Check if all queues are empty
        bool all_empty = false;
        {
            std::lock_guard<std::mutex> lock1(stage1_mutex_);
            std::lock_guard<std::mutex> lock2(stage2_mutex_);
            std::lock_guard<std::mutex> lock3(stage3_mutex_);
            std::lock_guard<std::mutex> lock4(stage4_mutex_);
            
            all_empty = stage1_queue_.empty() && stage2_queue_.empty() && 
                       stage3_queue_.empty() && stage4_queue_.empty();
        }
        
        if (all_empty) {
            std::lock_guard<std::mutex> lock(stats_mutex_);
            if (stats_.batches_in_pipeline == 0) {
                break;
            }
        }
        
        // Check timeout
        auto current_time = std::chrono::steady_clock::now();
        if (std::chrono::duration_cast<std::chrono::milliseconds>(current_time - start_time) >= timeout) {
            return Err(TXErrorCode::OperationFailed, "Timeout waiting for completion");
        }
        
        // Brief wait
        std::this_thread::sleep_for(std::chrono::milliseconds(10));
    }
    
    return Ok();
}

std::vector<std::unique_ptr<TXBatchData>> TXBatchPipeline::getCompletedBatches() {
    std::vector<std::unique_ptr<TXBatchData>> completed;

    std::lock_guard<std::mutex> lock(completed_mutex_);
    while (!completed_queue_.empty()) {
        completed.push_back(std::move(completed_queue_.front()));
        completed_queue_.pop();
    }

    return completed;
}

TXBatchPipeline::PipelineStats TXBatchPipeline::getStats() const {
    std::lock_guard<std::mutex> lock(stats_mutex_);

    PipelineStats current_stats = stats_;

    // Calculate real-time statistics
    if (current_stats.total_batches_processed > 0) {
        current_stats.avg_pipeline_time = std::chrono::microseconds(
            current_stats.total_pipeline_time.count() / current_stats.total_batches_processed);
    }

    // Calculate throughput with higher precision
    auto current_time = std::chrono::steady_clock::now();
    auto elapsed = std::chrono::duration_cast<std::chrono::milliseconds>(current_time - start_time_);
    if (elapsed.count() > 0) {
        // Convert to seconds with decimal precision
        double elapsed_seconds = elapsed.count() / 1000.0;
        current_stats.overall_throughput = static_cast<double>(current_stats.total_batches_processed) / elapsed_seconds;
    }

    // Get memory statistics
    if (memory_manager_) {
        auto memory_stats = memory_manager_->getUnifiedStats();
        current_stats.current_memory_usage = memory_stats.total_memory_usage;
        current_stats.memory_efficiency = memory_stats.overall_efficiency;

        if (memory_stats.total_memory_usage > current_stats.peak_memory_usage) {
            current_stats.peak_memory_usage = memory_stats.total_memory_usage;
        }
    }

    return current_stats;
}

void TXBatchPipeline::resetStats() {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    stats_ = PipelineStats{};
    start_time_ = std::chrono::steady_clock::now();
    last_throughput_update_ = start_time_;
}

std::string TXBatchPipeline::generatePerformanceReport() const {
    auto stats = getStats();

    std::ostringstream report;
    report << "=== TXBatchPipeline Performance Report ===\n";

    // Overall statistics
    report << "\nOverall Statistics:\n";
    report << "  Processed batches: " << stats.total_batches_processed << "\n";
    report << "  Failed batches: " << stats.total_batches_failed << "\n";
    report << "  Batches in pipeline: " << stats.batches_in_pipeline << "\n";
    report << "  Overall throughput: " << std::fixed << std::setprecision(2) << stats.overall_throughput << " batches/sec\n";
    report << "  Average processing time: " << stats.avg_pipeline_time.count() << " μs\n";

    // Memory statistics
    report << "\nMemory Statistics:\n";
    report << "  Current memory usage: " << (stats.current_memory_usage / 1024.0 / 1024.0) << " MB\n";
    report << "  Peak memory usage: " << (stats.peak_memory_usage / 1024.0 / 1024.0) << " MB\n";
    report << "  Memory efficiency: " << (stats.memory_efficiency * 100) << "%\n";

    // Queue statistics
    report << "\nQueue Statistics:\n";
    report << "  Max queue depth: " << stats.max_queue_depth << "\n";
    report << "  Average queue utilization: " << (stats.avg_queue_utilization * 100) << "%\n";

    return report.str();
}

void TXBatchPipeline::startWorkerThreads() {
    should_stop_.store(false);

    // Start worker threads for each stage
    for (size_t i = 0; i < config_.stage1_threads; ++i) {
        stage1_threads_.emplace_back(&TXBatchPipeline::stageWorker, this, 1, i);
    }

    for (size_t i = 0; i < config_.stage2_threads; ++i) {
        stage2_threads_.emplace_back(&TXBatchPipeline::stageWorker, this, 2, i);
    }

    for (size_t i = 0; i < config_.stage3_threads; ++i) {
        stage3_threads_.emplace_back(&TXBatchPipeline::stageWorker, this, 3, i);
    }

    for (size_t i = 0; i < config_.stage4_threads; ++i) {
        stage4_threads_.emplace_back(&TXBatchPipeline::stageWorker, this, 4, i);
    }
}

void TXBatchPipeline::stopWorkerThreads() {
    // Wait for all threads to finish
    for (auto& thread : stage1_threads_) {
        if (thread.joinable()) thread.join();
    }
    stage1_threads_.clear();

    for (auto& thread : stage2_threads_) {
        if (thread.joinable()) thread.join();
    }
    stage2_threads_.clear();

    for (auto& thread : stage3_threads_) {
        if (thread.joinable()) thread.join();
    }
    stage3_threads_.clear();

    for (auto& thread : stage4_threads_) {
        if (thread.joinable()) thread.join();
    }
    stage4_threads_.clear();
}

void TXBatchPipeline::stageWorker(size_t stage_index, size_t thread_id) {
    while (!should_stop_.load()) {
        // Process different queues based on stage index
        std::unique_ptr<TXBatchData> batch;

        // Get batch from corresponding queue
        switch (stage_index) {
            case 1: {
                std::unique_lock<std::mutex> lock(stage1_mutex_);
                stage1_cv_.wait(lock, [this] { return !stage1_queue_.empty() || should_stop_.load(); });

                if (should_stop_.load()) break;

                if (!stage1_queue_.empty()) {
                    batch = std::move(stage1_queue_.front());
                    stage1_queue_.pop();
                }
                break;
            }
            case 2: {
                std::unique_lock<std::mutex> lock(stage2_mutex_);
                stage2_cv_.wait(lock, [this] { return !stage2_queue_.empty() || should_stop_.load(); });

                if (should_stop_.load()) break;

                if (!stage2_queue_.empty()) {
                    batch = std::move(stage2_queue_.front());
                    stage2_queue_.pop();
                }
                break;
            }
            case 3: {
                std::unique_lock<std::mutex> lock(stage3_mutex_);
                stage3_cv_.wait(lock, [this] { return !stage3_queue_.empty() || should_stop_.load(); });

                if (should_stop_.load()) break;

                if (!stage3_queue_.empty()) {
                    batch = std::move(stage3_queue_.front());
                    stage3_queue_.pop();
                }
                break;
            }
            case 4: {
                std::unique_lock<std::mutex> lock(stage4_mutex_);
                stage4_cv_.wait(lock, [this] { return !stage4_queue_.empty() || should_stop_.load(); });

                if (should_stop_.load()) break;

                if (!stage4_queue_.empty()) {
                    batch = std::move(stage4_queue_.front());
                    stage4_queue_.pop();
                }
                break;
            }
        }

        if (!batch) continue;

        // Process batch
        auto start_time = std::chrono::steady_clock::now();
        TXResult<std::unique_ptr<TXBatchData>> result = Err<std::unique_ptr<TXBatchData>>(TXErrorCode::Unknown, "No stage processor");

        switch (stage_index) {
            case 1:
                if (stage1_) result = stage1_->process(std::move(batch));
                break;
            case 2:
                if (stage2_) result = stage2_->process(std::move(batch));
                break;
            case 3:
                if (stage3_) result = stage3_->process(std::move(batch));
                break;
            case 4:
                if (stage4_) result = stage4_->process(std::move(batch));
                break;
        }

        auto end_time = std::chrono::steady_clock::now();
        auto processing_time = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);

        if (result.isOk()) {
            auto processed_batch = std::move(result.value());

            // Pass processed batch to next stage or completion queue
            if (stage_index < 4) {
                // Pass to next stage
                switch (stage_index) {
                    case 1: {
                        std::lock_guard<std::mutex> lock(stage2_mutex_);
                        stage2_queue_.push(std::move(processed_batch));
                        stage2_cv_.notify_one();
                        break;
                    }
                    case 2: {
                        std::lock_guard<std::mutex> lock(stage3_mutex_);
                        stage3_queue_.push(std::move(processed_batch));
                        stage3_cv_.notify_one();
                        break;
                    }
                    case 3: {
                        std::lock_guard<std::mutex> lock(stage4_mutex_);
                        stage4_queue_.push(std::move(processed_batch));
                        stage4_cv_.notify_one();
                        break;
                    }
                }
            } else {
                // Final stage, add to completion queue
                {
                    std::lock_guard<std::mutex> lock(completed_mutex_);
                    completed_queue_.push(std::move(processed_batch));
                }

                // Update statistics
                {
                    std::lock_guard<std::mutex> lock(stats_mutex_);
                    stats_.total_batches_processed++;
                    stats_.total_pipeline_time += processing_time;
                    stats_.batches_in_pipeline--;
                }

                completed_cv_.notify_all();
            }
        } else {
            // Processing failed
            std::lock_guard<std::mutex> lock(stats_mutex_);
            stats_.total_batches_failed++;
            stats_.batches_in_pipeline--;
        }
    }
}

bool TXBatchPipeline::checkMemoryLimit() const {
    if (!memory_manager_) return true;

    auto stats = memory_manager_->getUnifiedStats();
    size_t limit_bytes = config_.memory_limit_mb * 1024 * 1024;

    return stats.total_memory_usage < limit_bytes;
}

void TXBatchPipeline::optimizeMemory() {
    if (memory_manager_) {
        memory_manager_->smartCleanup();
    }
}

TXResult<void> TXBatchPipeline::updateConfig(const PipelineConfig& config) {
    if (state_.load() == PipelineState::RUNNING) {
        return Err(TXErrorCode::OperationFailed, "Cannot update config while pipeline is running");
    }

    config_ = config;

    // Reinitialize memory manager
    if (memory_manager_) {
        TXUnifiedMemoryManager::Config memory_config;
        memory_config.memory_limit = config_.memory_limit_mb * 1024 * 1024;
        memory_config.enable_monitoring = config_.enable_performance_monitoring;
        memory_manager_->updateConfig(memory_config);
    }

    return Ok();
}

double TXBatchPipeline::getCurrentThroughput() const {
    auto stats = getStats();
    return stats.overall_throughput;
}

size_t TXBatchPipeline::getCurrentMemoryUsage() const {
    if (!memory_manager_) return 0;

    auto stats = memory_manager_->getUnifiedStats();
    return stats.total_memory_usage;
}

TXResult<void> TXBatchPipeline::setCustomStage(size_t stage_index, std::unique_ptr<TXPipelineStage> stage) {
    if (state_.load() == PipelineState::RUNNING) {
        return Err(TXErrorCode::OperationFailed, "Cannot set custom stage while pipeline is running");
    }

    switch (stage_index) {
        case 1: stage1_ = std::move(stage); break;
        case 2: stage2_ = std::move(stage); break;
        case 3: stage3_ = std::move(stage); break;
        case 4: stage4_ = std::move(stage); break;
        default:
            return Err(TXErrorCode::OperationFailed, "Invalid stage index: " + std::to_string(stage_index));
    }

    return Ok();
}

} // namespace TinaXlsx
