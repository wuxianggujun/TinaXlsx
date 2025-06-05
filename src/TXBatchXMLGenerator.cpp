//
// @file TXBatchXMLGenerator.cpp
// @brief Batch XML generator implementation
//

#include "TinaXlsx/TXBatchXMLGenerator.hpp"
#include "TinaXlsx/TXResult.hpp"
#include "TinaXlsx/TXError.hpp"
#include <algorithm>
#include <sstream>
#include <iomanip>
#include <thread>

namespace TinaXlsx {

// Thread local storage
thread_local std::unique_ptr<std::ostringstream> TXBatchXMLGenerator::xml_buffer_;
thread_local std::unique_ptr<std::string> TXBatchXMLGenerator::temp_string_;

// ==================== TXBatchXMLGenerator Implementation ====================

TXBatchXMLGenerator::TXBatchXMLGenerator(TXUnifiedMemoryManager& memory_manager,
                                         const XMLGeneratorConfig& config)
    : memory_manager_(memory_manager), config_(config) {
    
    if (config_.enable_template_caching) {
        loadDefaultTemplates();
    }
}

TXBatchXMLGenerator::~TXBatchXMLGenerator() {
    // Cleanup resources
}

void TXBatchXMLGenerator::loadDefaultTemplates() {
    // Load default XML templates
    TXXMLTemplate cell_template;
    cell_template.header = "<c";
    cell_template.footer = "</c>";
    cell_template.cell_template = "<c{attributes}><v>{value}</v></c>";
    
    TXXMLTemplate row_template;
    row_template.header = "<row r=\"{row_index}\">";
    row_template.footer = "</row>";
    row_template.row_template = "<row r=\"{row_index}\">{cells}</row>";
    
    TXXMLTemplate worksheet_template;
    worksheet_template.header = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<worksheet>";
    worksheet_template.footer = "</worksheet>";
    
    setTemplate("cell", cell_template);
    setTemplate("row", row_template);
    setTemplate("worksheet", worksheet_template);
}

TXResult<std::string> TXBatchXMLGenerator::generateCellXML(const TXCompactCell& cell, const std::string& cellRef) {
    auto start_time = std::chrono::steady_clock::now();

    try {
        std::ostringstream& buffer = getXMLBuffer();
        buffer.str("");
        buffer.clear();

        // 🚀 生成Excel兼容的单元格XML
        buffer << "<c r=\"" << cellRef << "\"";

        // 处理样式索引
        if (u32 styleIndex = cell.getStyleIndex(); styleIndex != 0) {
            buffer << " s=\"" << styleIndex << "\"";
        }

        // 获取单元格值和类型
        const cell_value_t& value = cell.getValue();
        const TXCompactCell::CellType cellType = cell.getType();

        // 处理不同类型的单元格
        if (cellType == TXCompactCell::CellType::String) {
            const std::string& str = cell.getStringValue();

            // 🚀 参考DOM方式的字符串处理逻辑
            if (shouldUseInlineString(str)) {
                // 内联字符串
                buffer << " t=\"inlineStr\">";
                buffer << "<is><t>" << escapeXML(str) << "</t></is>";
            } else {
                // 共享字符串 - 这里需要传入共享字符串池
                // 暂时使用内联方式，后续需要支持共享字符串池
                buffer << " t=\"inlineStr\">";
                buffer << "<is><t>" << escapeXML(str) << "</t></is>";
            }
        }
        else if (std::holds_alternative<double>(value)) {
            // 浮点数类型
            buffer << ">";
            buffer << "<v>" << formatDoubleForExcel(std::get<double>(value)) << "</v>";
        }
        else if (std::holds_alternative<int64_t>(value)) {
            // 整数类型
            buffer << ">";
            buffer << "<v>" << std::get<int64_t>(value) << "</v>";
        }
        else if (std::holds_alternative<bool>(value)) {
            // 布尔类型
            buffer << " t=\"b\">";
            buffer << "<v>" << (std::get<bool>(value) ? "1" : "0") << "</v>";
        }
        else {
            // 空单元格
            buffer << ">";
        }

        buffer << "</c>";

        std::string result = buffer.str();

        auto end_time = std::chrono::steady_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);

        // Update statistics
        updateStats(1, result.size(), duration);

        return Ok(result);

    } catch (const std::exception& e) {
        return Err<std::string>(TXErrorCode::Unknown, "Cell XML generation failed: " + std::string(e.what()));
    }
}

TXResult<std::string> TXBatchXMLGenerator::generateCellsXML(const std::vector<TXCompactCell>& cells) {
    if (config_.enable_parallel_generation && cells.size() >= config_.parallel_threshold) {
        return generateXMLParallel(cells);
    } else {
        return generateXMLSerial(cells);
    }
}

TXResult<std::string> TXBatchXMLGenerator::generateRowXML(size_t row_index, const std::vector<std::pair<std::string, TXCompactCell>>& cells) {
    auto start_time = std::chrono::steady_clock::now();

    try {
        std::ostringstream& buffer = getXMLBuffer();
        buffer.str("");
        buffer.clear();

        buffer << "<row r=\"" << row_index << "\">";

        // 🚀 使用新的generateCellXML方法，传入单元格引用
        for (const auto& [cellRef, cell] : cells) {
            auto cell_result = generateCellXML(cell, cellRef);
            if (cell_result.isOk()) {
                buffer << cell_result.value();
            }
        }

        buffer << "</row>";

        std::string result = buffer.str();

        auto end_time = std::chrono::steady_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);

        updateStats(cells.size(), result.size(), duration);

        return Ok(result);

    } catch (const std::exception& e) {
        return Err<std::string>(TXErrorCode::Unknown, "Row XML generation failed: " + std::string(e.what()));
    }
}

TXResult<std::string> TXBatchXMLGenerator::generateRowsXML(const std::vector<std::pair<size_t, std::vector<std::pair<std::string, TXCompactCell>>>>& rows) {
    auto start_time = std::chrono::steady_clock::now();

    try {
        std::ostringstream& buffer = getXMLBuffer();
        buffer.str("");
        buffer.clear();

        size_t total_cells = 0;

        for (const auto& row_pair : rows) {
            size_t row_index = row_pair.first;
            const auto& cells = row_pair.second;

            auto row_result = generateRowXML(row_index, cells);
            if (row_result.isOk()) {
                buffer << row_result.value();
                total_cells += cells.size();
            }
        }

        std::string result = buffer.str();

        auto end_time = std::chrono::steady_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);

        updateStats(total_cells, result.size(), duration);

        return Ok(result);

    } catch (const std::exception& e) {
        return Err<std::string>(TXErrorCode::Unknown, "Rows XML generation failed: " + std::string(e.what()));
    }
}

TXResult<std::string> TXBatchXMLGenerator::generateWorksheetXML(const std::string& sheet_name,
                                                               const std::vector<std::pair<size_t, std::vector<std::pair<std::string, TXCompactCell>>>>& rows) {
    auto start_time = std::chrono::steady_clock::now();
    
    try {
        std::ostringstream& buffer = getXMLBuffer();
        buffer.str("");
        buffer.clear();
        
        // XML declaration
        if (config_.include_xml_declaration) {
            buffer << "<?xml version=\"1.0\" encoding=\"" << config_.encoding << "\"?>\n";
        }
        
        // Worksheet start
        buffer << "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n";
        buffer << "<sheetData>\n";
        
        // Generate rows data
        auto rows_result = generateRowsXML(rows);
        if (rows_result.isOk()) {
            buffer << rows_result.value();
        }
        
        // Worksheet end
        buffer << "\n</sheetData>\n";
        buffer << "</worksheet>";
        
        std::string result = buffer.str();
        
        // Optimize output
        if (config_.enable_compression_hints) {
            optimizeXMLOutput(result);
        }
        
        auto end_time = std::chrono::steady_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
        
        // Calculate total cells
        size_t total_cells = 0;
        for (const auto& row_pair : rows) {
            total_cells += row_pair.second.size();
        }
        
        updateStats(total_cells, result.size(), duration);
        
        return Ok(result);
        
    } catch (const std::exception& e) {
        return Err<std::string>(TXErrorCode::Unknown, "Worksheet XML generation failed: " + std::string(e.what()));
    }
}

std::unique_ptr<TXBatchXMLGenerator::TXXMLStream> TXBatchXMLGenerator::createXMLStream(const std::string& root_element) {
    return std::make_unique<TXXMLStream>(*this, root_element);
}

void TXBatchXMLGenerator::setTemplate(const std::string& template_name, const TXXMLTemplate& xml_template) {
    std::lock_guard<std::mutex> lock(templates_mutex_);
    templates_[template_name] = xml_template;
}

const TXXMLTemplate* TXBatchXMLGenerator::getTemplate(const std::string& template_name) const {
    std::lock_guard<std::mutex> lock(templates_mutex_);
    auto it = templates_.find(template_name);
    return (it != templates_.end()) ? &it->second : nullptr;
}

void TXBatchXMLGenerator::clearTemplateCache() {
    std::lock_guard<std::mutex> lock(templates_mutex_);
    templates_.clear();
}

void TXBatchXMLGenerator::warmup(size_t warmup_iterations) {
    // Warmup generator
    std::vector<TXCompactCell> warmup_cells(100);
    
    for (size_t i = 0; i < warmup_iterations; ++i) {
        // Create test data
        for (size_t j = 0; j < warmup_cells.size(); ++j) {
            TXCompactCell cell;
            cell.setValue(static_cast<double>(i * j));
            warmup_cells[j] = cell;
        }
        
        // Generate XML
        (void)generateCellsXML(warmup_cells);
    }
}

void TXBatchXMLGenerator::optimizeMemory() {
    // Optimize memory usage
    {
        std::lock_guard<std::mutex> lock(strings_mutex_);
        if (interned_strings_.size() > 10000) {
            // Clear some unused strings
            interned_strings_.clear();
        }
    }
    
    // Trigger memory manager optimization
    memory_manager_.smartCleanup();
}

size_t TXBatchXMLGenerator::compactCache() {
    size_t freed = 0;

    {
        std::lock_guard<std::mutex> lock(templates_mutex_);
        // Keep common templates, clear others
        if (templates_.size() > 10) {
            auto old_size = templates_.size();
            // More intelligent cleanup strategy can be implemented here
            templates_.clear();
            loadDefaultTemplates();
            freed += (old_size - templates_.size()) * 100; // Estimate
        }
    }

    {
        std::lock_guard<std::mutex> lock(strings_mutex_);
        auto old_size = interned_strings_.size();
        interned_strings_.clear();
        freed += old_size * 50; // Estimate
    }

    return freed;
}

TXBatchXMLGenerator::XMLGeneratorStats TXBatchXMLGenerator::getStats() const {
    std::lock_guard<std::mutex> lock(stats_mutex_);

    XMLGeneratorStats current_stats = stats_;

    // Calculate derived statistics
    if (current_stats.total_cells_processed > 0) {
        current_stats.avg_generation_time = std::chrono::microseconds(
            current_stats.total_generation_time.count() / current_stats.total_cells_processed);

        auto total_seconds = current_stats.total_generation_time.count() / 1000000.0;
        if (total_seconds > 0) {
            current_stats.generation_rate = current_stats.total_cells_processed / total_seconds;
        }
    }

    // Calculate memory efficiency
    current_stats.memory_usage = memory_manager_.getUnifiedStats().total_memory_usage;
    if (current_stats.memory_usage > 0) {
        current_stats.memory_efficiency = static_cast<double>(current_stats.total_bytes_generated) / current_stats.memory_usage;
    }

    return current_stats;
}

void TXBatchXMLGenerator::resetStats() {
    std::lock_guard<std::mutex> lock(stats_mutex_);
    stats_ = XMLGeneratorStats{};
}

std::string TXBatchXMLGenerator::generatePerformanceReport() const {
    auto stats = getStats();

    std::ostringstream report;
    report << "=== TXBatchXMLGenerator Performance Report ===\n";

    report << "\nGeneration Statistics:\n";
    report << "  Total XML generated: " << stats.total_xml_generated << "\n";
    report << "  Total cells processed: " << stats.total_cells_processed << "\n";
    report << "  Total bytes generated: " << (stats.total_bytes_generated / 1024.0) << " KB\n";

    report << "\nPerformance Metrics:\n";
    report << "  Average generation time: " << stats.avg_generation_time.count() << " μs\n";
    report << "  Generation rate: " << std::fixed << std::setprecision(0) << stats.generation_rate << " cells/sec\n";

    report << "\nMemory Usage:\n";
    report << "  Current memory: " << (stats.memory_usage / 1024.0 / 1024.0) << " MB\n";
    report << "  Peak memory: " << (stats.peak_memory_usage / 1024.0 / 1024.0) << " MB\n";
    report << "  Memory efficiency: " << (stats.memory_efficiency * 100) << "%\n";

    report << "\nCache Statistics:\n";
    report << "  Template cache hits: " << stats.template_cache_hits << "\n";
    report << "  Template cache misses: " << stats.template_cache_misses << "\n";
    report << "  String intern hits: " << stats.string_intern_hits << "\n";

    return report.str();
}

size_t TXBatchXMLGenerator::getCurrentMemoryUsage() const {
    return memory_manager_.getUnifiedStats().total_memory_usage;
}

void TXBatchXMLGenerator::updateConfig(const XMLGeneratorConfig& config) {
    config_ = config;

    if (config_.enable_template_caching && templates_.empty()) {
        loadDefaultTemplates();
    }
}

std::ostringstream& TXBatchXMLGenerator::getXMLBuffer() {
    if (!xml_buffer_) {
        xml_buffer_ = std::make_unique<std::ostringstream>();
        xml_buffer_->str().reserve(config_.initial_buffer_size);
    }
    return *xml_buffer_;
}

std::string& TXBatchXMLGenerator::getTempString() {
    if (!temp_string_) {
        temp_string_ = std::make_unique<std::string>();
        temp_string_->reserve(1024);
    }
    return *temp_string_;
}

const std::string& TXBatchXMLGenerator::internString(const std::string& str) {
    if (!config_.enable_string_interning) {
        return str;
    }

    std::lock_guard<std::mutex> lock(strings_mutex_);
    auto it = interned_strings_.find(str);
    if (it != interned_strings_.end()) {
        std::lock_guard<std::mutex> stats_lock(stats_mutex_);
        stats_.string_intern_hits++;
        return it->second;
    }

    auto result = interned_strings_.emplace(str, str);
    return result.first->second;
}

std::string TXBatchXMLGenerator::escapeXML(const std::string& str) {
    std::string& temp = getTempString();
    temp.clear();
    temp.reserve(str.size() * 1.2);

    for (char c : str) {
        switch (c) {
            case '<': temp += "&lt;"; break;
            case '>': temp += "&gt;"; break;
            case '&': temp += "&amp;"; break;
            case '"': temp += "&quot;"; break;
            case '\'': temp += "&apos;"; break;
            default: temp += c; break;
        }
    }

    return temp;
}

bool TXBatchXMLGenerator::shouldUseInlineString(const std::string& str) const {
    // 🚀 参考DOM方式的字符串处理策略

    // 策略1: 空字符串或单字符使用内联（避免共享字符串池污染）
    if (str.empty() || str.length() == 1) return true;

    // 策略2: 包含XML特殊字符的字符串使用内联（避免转义复杂性）
    if (str.find_first_of("<>&\"'") != std::string::npos) return true;

    // 策略3: 包含控制字符的字符串使用内联（避免XML解析问题）
    if (str.find_first_of("\n\r\t") != std::string::npos) return true;

    // 策略4: 非常长的字符串（>100字符）使用内联（避免共享字符串池膨胀）
    if (str.length() > 100) return true;

    // 策略5: 2-100字符的普通字符串使用共享字符串（最大化复用效果）
    return false;
}

std::string TXBatchXMLGenerator::formatDoubleForExcel(double value) const {
    // 🚀 使用与DOM方式相同的数字格式化
    // 这里需要包含TXNumberUtils，暂时使用简单的格式化

    // 处理特殊值
    if (std::isnan(value)) return "0";
    if (std::isinf(value)) return value > 0 ? "1E+308" : "-1E+308";

    // 使用高精度格式化
    std::ostringstream oss;
    oss << std::fixed << std::setprecision(15) << value;
    std::string result = oss.str();

    // 移除尾随零
    if (result.find('.') != std::string::npos) {
        result.erase(result.find_last_not_of('0') + 1, std::string::npos);
        if (result.back() == '.') {
            result.pop_back();
        }
    }

    return result;
}

std::string TXBatchXMLGenerator::formatCellValue(const TXCompactCell& cell) {
    if (cell.isEmpty()) {
        return "";
    }

    auto value = cell.getValue();
    return std::visit([](const auto& val) -> std::string {
        using T = std::decay_t<decltype(val)>;
        if constexpr (std::is_same_v<T, std::monostate>) {
            return "";
        } else if constexpr (std::is_same_v<T, std::string>) {
            return val;
        } else if constexpr (std::is_same_v<T, double>) {
            return std::to_string(val);
        } else if constexpr (std::is_same_v<T, int64_t>) {
            return std::to_string(val);
        } else if constexpr (std::is_same_v<T, bool>) {
            return val ? "1" : "0";
        } else {
            return "";
        }
    }, value);
}

std::string TXBatchXMLGenerator::generateCellAttributes(const TXCompactCell& cell) {
    std::string attributes;

    // 添加单元格类型属性
    if (!cell.isEmpty()) {
        auto value = cell.getValue();
        if (std::holds_alternative<std::string>(value)) {
            attributes += " t=\"inlineStr\"";
        } else if (std::holds_alternative<bool>(value)) {
            attributes += " t=\"b\"";
        }
        // 数字类型不需要特殊属性
    }

    // 添加样式索引（如果有）
    if (cell.hasStyle()) {
        attributes += " s=\"" + std::to_string(cell.getStyleIndex()) + "\"";
    }

    return attributes;
}

TXResult<std::string> TXBatchXMLGenerator::generateXMLParallel(const std::vector<TXCompactCell>& cells) {
    // Simplified parallel implementation
    return generateXMLSerial(cells);
}

TXResult<std::string> TXBatchXMLGenerator::generateXMLSerial(const std::vector<TXCompactCell>& cells) {
    auto start_time = std::chrono::steady_clock::now();

    try {
        std::ostringstream& buffer = getXMLBuffer();
        buffer.str("");
        buffer.clear();

        // ❌ 这个方法需要单元格引用，但这里没有提供
        // 暂时生成简单的引用，或者标记为不支持
        for (size_t i = 0; i < cells.size(); ++i) {
            // 生成简单的单元格引用 A1, B1, C1...
            std::string cellRef = "A" + std::to_string(i + 1);
            auto cell_result = generateCellXML(cells[i], cellRef);
            if (cell_result.isOk()) {
                buffer << cell_result.value();
            }
        }

        std::string result = buffer.str();

        auto end_time = std::chrono::steady_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);

        updateStats(cells.size(), result.size(), duration);

        return Ok(result);

    } catch (const std::exception& e) {
        return Err<std::string>(TXErrorCode::Unknown, "Serial XML generation failed: " + std::string(e.what()));
    }
}

void TXBatchXMLGenerator::updateStats(size_t cells_processed, size_t bytes_generated,
                                     std::chrono::microseconds generation_time) {
    std::lock_guard<std::mutex> lock(stats_mutex_);

    stats_.total_xml_generated++;
    stats_.total_cells_processed += cells_processed;
    stats_.total_bytes_generated += bytes_generated;
    stats_.total_generation_time += generation_time;

    size_t current_memory = memory_manager_.getUnifiedStats().total_memory_usage;
    stats_.memory_usage = current_memory;
    if (current_memory > stats_.peak_memory_usage) {
        stats_.peak_memory_usage = current_memory;
    }
}

double TXBatchXMLGenerator::estimateCompressionRatio(const std::string& xml) const {
    // Simplified compression ratio estimation
    size_t repeated_chars = 0;
    if (xml.size() > 1) {
        for (size_t i = 1; i < xml.size(); ++i) {
            if (xml[i] == xml[i-1]) {
                repeated_chars++;
            }
        }
    }

    return 1.0 - (static_cast<double>(repeated_chars) / xml.size());
}

void TXBatchXMLGenerator::optimizeXMLOutput(std::string& xml) {
    if (!config_.pretty_print) {
        // Remove unnecessary whitespace characters
        xml.erase(std::remove(xml.begin(), xml.end(), '\n'), xml.end());
        xml.erase(std::remove(xml.begin(), xml.end(), '\r'), xml.end());

        // Remove extra spaces between tags
        size_t pos = 0;
        while ((pos = xml.find(">  <", pos)) != std::string::npos) {
            xml.replace(pos, 4, "><");
            pos += 2;
        }
    }
}

// ==================== TXXMLStream Implementation ====================

TXBatchXMLGenerator::TXXMLStream::TXXMLStream(TXBatchXMLGenerator& generator, const std::string& root_element)
    : generator_(generator), root_element_(root_element), finalized_(false) {
    stream_ = std::make_unique<std::ostringstream>();
    // Write root element start
    *stream_ << "<" << root_element_ << ">";
}

TXBatchXMLGenerator::TXXMLStream::~TXXMLStream() {
    if (!finalized_) {
        // Finalize stream if not already done
        (void)finalize();
    }
}

TXResult<void> TXBatchXMLGenerator::TXXMLStream::writeElement(const std::string& element, const std::string& content) {
    if (finalized_) {
        return Err(TXErrorCode::Unknown, "Stream is finalized");
    }

    // Write element
    *stream_ << "<" << element << ">" << content << "</" << element << ">";
    return Ok();
}

TXResult<void> TXBatchXMLGenerator::TXXMLStream::writeCell(const TXCompactCell& cell, const std::string& cellRef) {
    if (finalized_) {
        return Err(TXErrorCode::Unknown, "Stream is finalized");
    }

    // Generate cell XML using the generator
    auto cell_result = generator_.generateCellXML(cell, cellRef);
    if (cell_result.isOk()) {
        *stream_ << cell_result.value();
        return Ok();
    } else {
        return Err(cell_result.error());
    }
}

TXResult<void> TXBatchXMLGenerator::TXXMLStream::writeRow(size_t row_index, const std::vector<std::pair<std::string, TXCompactCell>>& cells) {
    if (finalized_) {
        return Err(TXErrorCode::Unknown, "Stream is finalized");
    }

    // Generate row XML using the generator
    auto row_result = generator_.generateRowXML(row_index, cells);
    if (row_result.isOk()) {
        *stream_ << row_result.value();
        return Ok();
    } else {
        return Err(row_result.error());
    }
}

TXResult<std::string> TXBatchXMLGenerator::TXXMLStream::finalize() {
    if (finalized_) {
        return Err<std::string>(TXErrorCode::Unknown, "Stream is already finalized");
    }

    // Write root element end
    *stream_ << "</" << root_element_ << ">";

    finalized_ = true;
    return Ok(stream_->str());
}

} // namespace TinaXlsx
