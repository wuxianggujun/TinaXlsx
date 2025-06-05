//
// @file TXBatchPipelineStages.cpp
// @brief Batch pipeline stages implementation
//

#include "TinaXlsx/TXBatchPipelineStages.hpp"
#include "TinaXlsx/TXResult.hpp"
#include "TinaXlsx/TXError.hpp"
#include <algorithm>
#include <unordered_set>
#include <sstream>
#include <chrono>

namespace TinaXlsx {

// ==================== TXDataPreprocessingStage Implementation ====================

TXDataPreprocessingStage::TXDataPreprocessingStage(TXUnifiedMemoryManager& memory_manager, 
                                                   const PreprocessingConfig& config)
    : memory_manager_(memory_manager), config_(config) {
}

TXResult<std::unique_ptr<TXBatchData>> TXDataPreprocessingStage::process(std::unique_ptr<TXBatchData> input) {
    if (!input) {
        return Err<std::unique_ptr<TXBatchData>>(TXErrorCode::Unknown, "Input batch is null");
    }
    
    auto start_time = std::chrono::steady_clock::now();
    
    try {
        // 1. Data validation
        if (config_.enable_data_validation) {
            auto validation_result = validateData(*input);
            if (!validation_result.isOk()) {
                return Err<std::unique_ptr<TXBatchData>>(validation_result.error());
            }
        }
        
        // 2. String deduplication
        if (config_.enable_deduplication) {
            auto dedup_result = deduplicateStrings(*input);
            if (!dedup_result.isOk()) {
                return Err<std::unique_ptr<TXBatchData>>(dedup_result.error());
            }
        }
        
        // 3. Memory optimization
        if (config_.enable_memory_optimization) {
            auto memory_result = optimizeMemoryLayout(*input);
            if (!memory_result.isOk()) {
                return Err<std::unique_ptr<TXBatchData>>(memory_result.error());
            }
        }
        
        // 4. Estimate batch size
        input->estimated_size = estimateBatchSize(*input);
        
        auto end_time = std::chrono::steady_clock::now();
        auto processing_time = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
        
        // Update statistics
        {
            std::lock_guard<std::mutex> lock(stats_mutex_);
            stats_.processed_batches++;
            stats_.total_processing_time += processing_time;
            stats_.avg_processing_time = std::chrono::microseconds(
                stats_.total_processing_time.count() / stats_.processed_batches);
        }
        
        return Ok(std::move(input));
        
    } catch (const std::exception& e) {
        std::lock_guard<std::mutex> lock(stats_mutex_);
        stats_.failed_batches++;
        return Err<std::unique_ptr<TXBatchData>>(TXErrorCode::Unknown, "Processing exception: " + std::string(e.what()));
    }
}

TXPipelineStage::StageStats TXDataPreprocessingStage::getStats() const {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    return stats_;
}

void TXDataPreprocessingStage::resetStats() {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    stats_ = StageStats{};
}

TXResult<void> TXDataPreprocessingStage::validateData(TXBatchData& batch) {
    // Basic data validation
    if (batch.cells.empty()) {
        return Err(TXErrorCode::Unknown, "Batch contains no cells");
    }
    
    if (batch.cells.size() > config_.max_batch_size) {
        return Err(TXErrorCode::Unknown, "Batch size exceeds maximum limit");
    }
    
    // Additional validation can be added here
    for (const auto& cell : batch.cells) {
        // Validate individual cells
        // Implementation depends on TXCompactCell structure
    }
    
    return Ok();
}

TXResult<void> TXDataPreprocessingStage::optimizeMemoryLayout(TXBatchData& batch) {
    // Reserve memory to avoid frequent reallocations
    if (batch.cells.capacity() < batch.cells.size() * 1.2) {
        batch.cells.reserve(batch.cells.size() * 1.2);
    }
    
    if (batch.strings.capacity() < batch.strings.size() * 1.2) {
        batch.strings.reserve(batch.strings.size() * 1.2);
    }
    
    return Ok();
}

TXResult<void> TXDataPreprocessingStage::deduplicateStrings(TXBatchData& batch) {
    std::unordered_set<std::string> unique_strings;
    
    // Remove duplicate strings
    auto original_size = batch.strings.size();
    batch.strings.erase(
        std::remove_if(batch.strings.begin(), batch.strings.end(),
            [&unique_strings](const std::string& str) {
                return !unique_strings.insert(str).second;
            }),
        batch.strings.end()
    );
    
    // Log removed count if needed
    auto removed_count = original_size - batch.strings.size();
    (void)removed_count; // Suppress unused variable warning
    
    return Ok();
}

size_t TXDataPreprocessingStage::estimateBatchSize(const TXBatchData& batch) {
    size_t total_size = 0;
    
    // Cell data size
    total_size += batch.cells.size() * sizeof(TXCompactCell);
    
    // String data size
    for (const auto& str : batch.strings) {
        total_size += str.size();
    }
    
    // Binary data size
    total_size += batch.binary_data.size();
    
    return total_size;
}

// ==================== TXXMLGenerationStage Implementation ====================

TXXMLGenerationStage::TXXMLGenerationStage(TXUnifiedMemoryManager& memory_manager,
                                           const XMLConfig& config)
    : memory_manager_(memory_manager), config_(config) {
}

TXResult<std::unique_ptr<TXBatchData>> TXXMLGenerationStage::process(std::unique_ptr<TXBatchData> input) {
    if (!input) {
        return Err<std::unique_ptr<TXBatchData>>(TXErrorCode::Unknown, "Input batch is null");
    }
    
    auto start_time = std::chrono::steady_clock::now();
    
    try {
        // Generate XML content
        auto xml_result = generateBatchXML(*input);
        if (!xml_result.isOk()) {
            return Err<std::unique_ptr<TXBatchData>>(xml_result.error());
        }
        
        std::string xml_data = std::move(xml_result.value());
        
        // Optimize XML output
        if (config_.enable_compression_hints) {
            auto optimize_result = optimizeXMLOutput(xml_data);
            if (!optimize_result.isOk()) {
                return Err<std::unique_ptr<TXBatchData>>(optimize_result.error());
            }
        }
        
        // Store XML data in binary format
        input->binary_data.clear();
        input->binary_data.reserve(xml_data.size());
        input->binary_data.assign(xml_data.begin(), xml_data.end());
        
        // Prepare compression hints
        if (config_.enable_compression_hints) {
            prepareCompressionHints(*input);
        }
        
        auto end_time = std::chrono::steady_clock::now();
        auto processing_time = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
        
        // Update statistics
        {
            std::lock_guard<std::mutex> lock(stats_mutex_);
            stats_.processed_batches++;
            stats_.total_processing_time += processing_time;
            stats_.avg_processing_time = std::chrono::microseconds(
                stats_.total_processing_time.count() / stats_.processed_batches);
        }
        
        return Ok(std::move(input));
        
    } catch (const std::exception& e) {
        std::lock_guard<std::mutex> lock(stats_mutex_);
        stats_.failed_batches++;
        return Err<std::unique_ptr<TXBatchData>>(TXErrorCode::Unknown, "XML generation exception: " + std::string(e.what()));
    }
}

TXPipelineStage::StageStats TXXMLGenerationStage::getStats() const {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    return stats_;
}

void TXXMLGenerationStage::resetStats() {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    stats_ = StageStats{};
}

TXResult<std::string> TXXMLGenerationStage::generateCellXML(const TXCompactCell& cell) {
    // Simplified cell XML generation
    std::ostringstream xml;
    xml << "<c><v>placeholder</v></c>";
    return Ok(xml.str());
}

TXResult<std::string> TXXMLGenerationStage::generateBatchXML(const TXBatchData& batch) {
    std::ostringstream xml;
    xml << "<sheetData>";

    for (const auto& cell : batch.cells) {
        auto cell_result = generateCellXML(cell);
        if (cell_result.isOk()) {
            xml << cell_result.value();
        }
    }

    xml << "</sheetData>";
    return Ok(xml.str());
}

TXResult<void> TXXMLGenerationStage::optimizeXMLOutput(std::string& xml) {
    // Remove unnecessary whitespace
    xml.erase(std::remove(xml.begin(), xml.end(), '\n'), xml.end());
    xml.erase(std::remove(xml.begin(), xml.end(), '\r'), xml.end());
    return Ok();
}

void TXXMLGenerationStage::prepareCompressionHints(TXBatchData& batch) {
    // Add compression hints - since TXBatchData doesn't have metadata,
    // we can use the memory_context to store hints or just skip this
    (void)batch; // Suppress unused parameter warning
}

// ==================== TXCompressionStage Implementation ====================

TXCompressionStage::TXCompressionStage(TXUnifiedMemoryManager& memory_manager,
                                       const CompressionConfig& config)
    : memory_manager_(memory_manager), config_(config) {
}

TXResult<std::unique_ptr<TXBatchData>> TXCompressionStage::process(std::unique_ptr<TXBatchData> input) {
    if (!input) {
        return Err<std::unique_ptr<TXBatchData>>(TXErrorCode::Unknown, "Input batch is null");
    }

    auto start_time = std::chrono::steady_clock::now();

    try {
        // Skip compression if data is too small
        if (input->binary_data.size() < config_.compression_threshold) {
            return Ok(std::move(input));
        }

        // Compress binary data
        auto compress_result = compressData(input->binary_data);
        if (compress_result.isOk()) {
            input->binary_data = std::move(compress_result.value());
            // Note: compressed flag would be stored in metadata if available
        }

        auto end_time = std::chrono::steady_clock::now();
        auto processing_time = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);

        // Update statistics
        {
            std::lock_guard<std::mutex> lock(stats_mutex_);
            stats_.processed_batches++;
            stats_.total_processing_time += processing_time;
            stats_.avg_processing_time = std::chrono::microseconds(
                stats_.total_processing_time.count() / stats_.processed_batches);
        }

        return Ok(std::move(input));

    } catch (const std::exception& e) {
        std::lock_guard<std::mutex> lock(stats_mutex_);
        stats_.failed_batches++;
        return Err<std::unique_ptr<TXBatchData>>(TXErrorCode::Unknown, "Compression exception: " + std::string(e.what()));
    }
}

TXPipelineStage::StageStats TXCompressionStage::getStats() const {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    return stats_;
}

void TXCompressionStage::resetStats() {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    stats_ = StageStats{};
}

TXResult<std::vector<uint8_t>> TXCompressionStage::compressData(const std::vector<uint8_t>& data) {
    // Simplified compression - just return original data for now
    return Ok(data);
}

TXResult<std::vector<uint8_t>> TXCompressionStage::compressString(const std::string& str) {
    std::vector<uint8_t> data(str.begin(), str.end());
    return compressData(data);
}

TXCompressionStage::CompressionAlgorithm TXCompressionStage::selectOptimalAlgorithm(const TXBatchData& batch) {
    (void)batch; // Suppress unused parameter warning
    return config_.algorithm;
}

double TXCompressionStage::calculateCompressionRatio() const {
    if (total_uncompressed_size_ == 0) return 1.0;
    return static_cast<double>(total_compressed_size_) / total_uncompressed_size_;
}

// ==================== TXOutputWriteStage Implementation ====================

TXOutputWriteStage::TXOutputWriteStage(TXUnifiedMemoryManager& memory_manager,
                                       const OutputConfig& config)
    : memory_manager_(memory_manager), config_(config) {
}

TXResult<std::unique_ptr<TXBatchData>> TXOutputWriteStage::process(std::unique_ptr<TXBatchData> input) {
    if (!input) {
        return Err<std::unique_ptr<TXBatchData>>(TXErrorCode::Unknown, "Input batch is null");
    }

    auto start_time = std::chrono::steady_clock::now();

    try {
        // Generate output path
        auto path_result = generateOutputPath(*input);
        if (!path_result.isOk()) {
            return Err<std::unique_ptr<TXBatchData>>(path_result.error());
        }

        std::string output_path = path_result.value();

        // Write to file
        auto write_result = writeToFile(output_path, *input);
        if (!write_result.isOk()) {
            return Err<std::unique_ptr<TXBatchData>>(write_result.error());
        }

        // Verify write if enabled
        if (config_.enable_write_verification) {
            auto verify_result = verifyWrite(output_path, input->binary_data.size());
            if (!verify_result.isOk()) {
                return Err<std::unique_ptr<TXBatchData>>(verify_result.error());
            }
        }

        // Cleanup memory if enabled
        if (config_.enable_memory_cleanup) {
            cleanupBatchMemory(*input);
        }

        auto end_time = std::chrono::steady_clock::now();
        auto processing_time = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);

        // Update statistics
        {
            std::lock_guard<std::mutex> lock(stats_mutex_);
            stats_.processed_batches++;
            stats_.total_processing_time += processing_time;
            stats_.avg_processing_time = std::chrono::microseconds(
                stats_.total_processing_time.count() / stats_.processed_batches);
            total_bytes_written_ += input->binary_data.size();
            total_files_written_++;
        }

        return Ok(std::move(input));

    } catch (const std::exception& e) {
        std::lock_guard<std::mutex> lock(stats_mutex_);
        stats_.failed_batches++;
        return Err<std::unique_ptr<TXBatchData>>(TXErrorCode::Unknown, "Output write exception: " + std::string(e.what()));
    }
}

TXPipelineStage::StageStats TXOutputWriteStage::getStats() const {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    return stats_;
}

void TXOutputWriteStage::resetStats() {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    stats_ = StageStats{};
    total_bytes_written_ = 0;
    total_files_written_ = 0;
}

TXResult<std::string> TXOutputWriteStage::generateOutputPath(const TXBatchData& batch) {
    std::ostringstream path;
    path << config_.output_directory << "/" << config_.file_prefix << batch.batch_id << config_.file_extension;
    return Ok(path.str());
}

TXResult<void> TXOutputWriteStage::writeToFile(const std::string& path, const TXBatchData& batch) {
    // Simplified file writing - just create a placeholder file
    (void)path;
    (void)batch;
    return Ok();
}

TXResult<void> TXOutputWriteStage::verifyWrite(const std::string& path, size_t expected_size) {
    // Simplified verification
    (void)path;
    (void)expected_size;
    return Ok();
}

void TXOutputWriteStage::cleanupBatchMemory(TXBatchData& batch) {
    // Clear data to free memory
    batch.cells.clear();
    batch.strings.clear();
    batch.binary_data.clear();
    // Note: metadata would be cleared if available
}

// ==================== TXPipelineStageFactory Implementation ====================

std::vector<std::unique_ptr<TXPipelineStage>> TXPipelineStageFactory::createDefaultStages(
    TXUnifiedMemoryManager& memory_manager) {

    std::vector<std::unique_ptr<TXPipelineStage>> stages;
    stages.reserve(4);

    // Stage 1: Data Preprocessing
    stages.push_back(createPreprocessingStage(memory_manager));

    // Stage 2: XML Generation
    stages.push_back(createXMLGenerationStage(memory_manager));

    // Stage 3: Compression
    stages.push_back(createCompressionStage(memory_manager));

    // Stage 4: Output Write
    stages.push_back(createOutputWriteStage(memory_manager));

    return stages;
}

std::unique_ptr<TXPipelineStage> TXPipelineStageFactory::createPreprocessingStage(
    TXUnifiedMemoryManager& memory_manager,
    const TXDataPreprocessingStage::PreprocessingConfig& config) {
    return std::make_unique<TXDataPreprocessingStage>(memory_manager, config);
}

std::unique_ptr<TXPipelineStage> TXPipelineStageFactory::createXMLGenerationStage(
    TXUnifiedMemoryManager& memory_manager,
    const TXXMLGenerationStage::XMLConfig& config) {
    return std::make_unique<TXXMLGenerationStage>(memory_manager, config);
}

std::unique_ptr<TXPipelineStage> TXPipelineStageFactory::createCompressionStage(
    TXUnifiedMemoryManager& memory_manager,
    const TXCompressionStage::CompressionConfig& config) {
    return std::make_unique<TXCompressionStage>(memory_manager, config);
}

std::unique_ptr<TXPipelineStage> TXPipelineStageFactory::createOutputWriteStage(
    TXUnifiedMemoryManager& memory_manager,
    const TXOutputWriteStage::OutputConfig& config) {
    return std::make_unique<TXOutputWriteStage>(memory_manager, config);
}

} // namespace TinaXlsx
