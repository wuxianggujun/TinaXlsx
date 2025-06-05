//
// @file TXXMLTemplates.hpp
// @brief Excel XML模板系统 - 零拷贝高速XML生成
//

#pragma once

#include <string_view>
#include <string>
#include <fmt/format.h>

namespace TinaXlsx {

/**
 * @brief 编译时XML模板 - 零拷贝高性能
 */
class TXCompiledXMLTemplates {
public:
    // ==================== 工作表模板 ====================
    
    static constexpr std::string_view WORKSHEET_HEADER = 
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">)";
    
    static constexpr std::string_view WORKSHEET_FOOTER = "</worksheet>";
    
    static constexpr std::string_view SHEET_DATA_START = "<sheetData>";
    static constexpr std::string_view SHEET_DATA_END = "</sheetData>";
    
    // 行模板 - 使用 {} 为fmt::format占位符
    static constexpr std::string_view ROW_START = R"(<row r="{}">)";
    static constexpr std::string_view ROW_END = "</row>";
    
    // 单元格模板
    static constexpr std::string_view CELL_NUMBER = R"(<c r="{}" t="n"><v>{}</v></c>)";
    static constexpr std::string_view CELL_STRING = R"(<c r="{}" t="s"><v>{}</v></c>)";
    static constexpr std::string_view CELL_INLINE_STRING = R"(<c r="{}" t="inlineStr"><is><t>{}</t></is></c>)";
    static constexpr std::string_view CELL_BOOLEAN = R"(<c r="{}" t="b"><v>{}</v></c>)";
    static constexpr std::string_view CELL_FORMULA = R"(<c r="{}"><f>{}</f><v>{}</v></c>)";
    
    // ==================== 共享字符串模板 ====================
    
    static constexpr std::string_view SHARED_STRINGS_HEADER = 
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{}" uniqueCount="{}">)";
    
    static constexpr std::string_view SHARED_STRINGS_FOOTER = "</sst>";
    static constexpr std::string_view SHARED_STRING_ITEM = "<si><t>{}</t></si>";
    
    // ==================== 工作簿模板 ====================
    
    static constexpr std::string_view WORKBOOK_HEADER = 
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>)";
    
    static constexpr std::string_view WORKBOOK_FOOTER = "</sheets></workbook>";
    static constexpr std::string_view SHEET_ENTRY = R"(<sheet name="{}" sheetId="{}" r:id="rId{}"/>)";
    
    // ==================== 内容类型模板 ====================
    
    static constexpr std::string_view CONTENT_TYPES_HEADER = 
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>)";
    
    static constexpr std::string_view CONTENT_TYPES_FOOTER = "</Types>";
    static constexpr std::string_view WORKSHEET_CONTENT_TYPE = 
        R"(<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>)";
    
    static constexpr std::string_view SHARED_STRINGS_CONTENT_TYPE = 
        R"(<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>)";
    
    // ==================== 关系模板 ====================
    
    static constexpr std::string_view MAIN_RELS = 
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)";
    
    static constexpr std::string_view WORKBOOK_RELS_HEADER = 
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">)";
    
    static constexpr std::string_view WORKBOOK_RELS_FOOTER = "</Relationships>";
    static constexpr std::string_view WORKSHEET_REL = 
        R"(<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{}.xml"/>)";
    
    static constexpr std::string_view SHARED_STRINGS_REL = 
        R"(<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>)";
    
    // ==================== 应用属性模板 ====================
    
    static constexpr std::string_view APP_PROPERTIES = 
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>TinaXlsx</Application>
<DocSecurity>0</DocSecurity>
<ScaleCrop>false</ScaleCrop>
<Company></Company>
<LinksUpToDate>false</LinksUpToDate>
<SharedDoc>false</SharedDoc>
<HyperlinksChanged>false</HyperlinksChanged>
<AppVersion>1.0.0000</AppVersion>
</Properties>)";
    
    static constexpr std::string_view CORE_PROPERTIES = 
        R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>TinaXlsx</dc:creator>
<dcterms:created xsi:type="dcterms:W3CDTF">{}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">{}</dcterms:modified>
</cp:coreProperties>)";
    
    // ==================== 模板应用函数 ====================
    
    /**
     * @brief 应用带单个参数的模板
     */
    template<typename T>
    static std::string applyTemplate(std::string_view template_str, T&& arg) {
        return fmt::format(template_str, std::forward<T>(arg));
    }
    
    /**
     * @brief 应用带多个参数的模板
     */
    template<typename... Args>
    static std::string applyTemplate(std::string_view template_str, Args&&... args) {
        return fmt::format(template_str, std::forward<Args>(args)...);
    }
    
    /**
     * @brief 快速生成数值单元格XML
     */
    static std::string makeNumberCell(std::string_view coord, double value) {
        return fmt::format(CELL_NUMBER, coord, value);
    }
    
    /**
     * @brief 快速生成字符串单元格XML
     */
    static std::string makeStringCell(std::string_view coord, std::string_view text) {
        return fmt::format(CELL_INLINE_STRING, coord, escapeXML(text));
    }
    
    /**
     * @brief 快速生成行开始标签
     */
    static std::string makeRowStart(uint32_t row_number) {
        return fmt::format(ROW_START, row_number);
    }
    
    /**
     * @brief 转义XML特殊字符
     */
    static std::string escapeXML(std::string_view str) {
        std::string result;
        result.reserve(str.size() * 1.1); // 预留转义字符空间
        
        for (char c : str) {
            switch (c) {
                case '<': result += "&lt;"; break;
                case '>': result += "&gt;"; break;
                case '&': result += "&amp;"; break;
                case '"': result += "&quot;"; break;
                case '\'': result += "&apos;"; break;
                default: result += c; break;
            }
        }
        
        return result;
    }
    
    /**
     * @brief 批量转义XML字符串
     */
    static void escapeXMLBatch(const std::vector<std::string>& input, 
                              std::vector<std::string>& output) {
        output.clear();
        output.reserve(input.size());
        
        for (const auto& str : input) {
            output.push_back(escapeXML(str));
        }
    }
    
    /**
     * @brief 获取当前ISO时间戳
     */
    static std::string getCurrentTimestamp() {
        auto now = std::chrono::system_clock::now();
        auto time_t = std::chrono::system_clock::to_time_t(now);
        auto tm = *std::gmtime(&time_t);
        
        return fmt::format("{:04d}-{:02d}-{:02d}T{:02d}:{:02d}:{:02d}Z",
                          tm.tm_year + 1900, tm.tm_mon + 1, tm.tm_mday,
                          tm.tm_hour, tm.tm_min, tm.tm_sec);
    }
};

/**
 * @brief 高性能XML模板缓存
 */
class TXXMLTemplateCache {
public:
    struct CachedTemplate {
        std::string compiled_template;
        size_t estimated_size;
        bool is_dynamic;
        std::chrono::steady_clock::time_point created_time;
    };
    
    /**
     * @brief 缓存编译后的模板
     */
    void cacheTemplate(const std::string& key, const std::string& compiled_template, 
                      size_t estimated_size = 0) {
        std::lock_guard<std::mutex> lock(mutex_);
        
        CachedTemplate cached;
        cached.compiled_template = compiled_template;
        cached.estimated_size = estimated_size;
        cached.is_dynamic = false;
        cached.created_time = std::chrono::steady_clock::now();
        
        cache_[key] = std::move(cached);
    }
    
    /**
     * @brief 获取缓存的模板
     */
    std::optional<std::string> getCachedTemplate(const std::string& key) const {
        std::lock_guard<std::mutex> lock(mutex_);
        
        auto it = cache_.find(key);
        if (it != cache_.end()) {
            hits_++;
            return it->second.compiled_template;
        }
        
        misses_++;
        return std::nullopt;
    }
    
    /**
     * @brief 清理过期缓存
     */
    void cleanup(std::chrono::seconds max_age = std::chrono::seconds(300)) {
        std::lock_guard<std::mutex> lock(mutex_);
        
        auto now = std::chrono::steady_clock::now();
        auto it = cache_.begin();
        
        while (it != cache_.end()) {
            if (now - it->second.created_time > max_age) {
                it = cache_.erase(it);
            } else {
                ++it;
            }
        }
    }
    
    /**
     * @brief 获取缓存统计
     */
    struct CacheStats {
        size_t total_entries;
        size_t hits;
        size_t misses;
        double hit_ratio;
        size_t memory_usage_bytes;
    };
    
    CacheStats getStats() const {
        std::lock_guard<std::mutex> lock(mutex_);
        
        CacheStats stats;
        stats.total_entries = cache_.size();
        stats.hits = hits_;
        stats.misses = misses_;
        stats.hit_ratio = (hits_ + misses_ > 0) ? 
            static_cast<double>(hits_) / (hits_ + misses_) : 0.0;
        
        // 估算内存使用量
        stats.memory_usage_bytes = 0;
        for (const auto& [key, cached] : cache_) {
            stats.memory_usage_bytes += key.size() + cached.compiled_template.size() + 64;
        }
        
        return stats;
    }
    
    /**
     * @brief 清空缓存
     */
    void clear() {
        std::lock_guard<std::mutex> lock(mutex_);
        cache_.clear();
        hits_ = 0;
        misses_ = 0;
    }

private:
    mutable std::mutex mutex_;
    std::unordered_map<std::string, CachedTemplate> cache_;
    mutable std::atomic<size_t> hits_{0};
    mutable std::atomic<size_t> misses_{0};
};

/**
 * @brief Excel坐标转换工具
 */
class TXExcelCoordinates {
public:
    /**
     * @brief 将行列转换为Excel格式 (如: A1, B2, AA10)
     */
    static std::string rowColToString(uint32_t row, uint32_t col) {
        std::string result;
        
        // 列转换为字母 (A-Z, AA-ZZ, AAA-ZZZ...)
        uint32_t temp_col = col + 1; // Excel列从1开始
        do {
            temp_col--;
            result = char('A' + (temp_col % 26)) + result;
            temp_col /= 26;
        } while (temp_col > 0);
        
        // 行号 (1-based)
        result += std::to_string(row + 1);
        
        return result;
    }
    
    /**
     * @brief 将坐标对转换为Excel格式
     */
    static std::string coordToString(uint32_t packed_coord) {
        uint32_t row = packed_coord >> 16;
        uint32_t col = packed_coord & 0xFFFF;
        return rowColToString(row, col);
    }
    
    /**
     * @brief 批量转换坐标
     */
    static void coordsBatchToStrings(const uint32_t* coords, size_t count, 
                                    std::vector<std::string>& output) {
        output.clear();
        output.reserve(count);
        
        for (size_t i = 0; i < count; ++i) {
            output.push_back(coordToString(coords[i]));
        }
    }
    
    /**
     * @brief 解析Excel格式坐标 (如: A1 -> {0, 0})
     */
    static std::pair<uint32_t, uint32_t> parseCoordinate(std::string_view coord_str) {
        size_t i = 0;
        uint32_t col = 0;
        
        // 解析列部分
        while (i < coord_str.size() && std::isalpha(coord_str[i])) {
            col = col * 26 + (std::toupper(coord_str[i]) - 'A' + 1);
            ++i;
        }
        col--; // 转换为0-based
        
        // 解析行部分
        uint32_t row = 0;
        while (i < coord_str.size() && std::isdigit(coord_str[i])) {
            row = row * 10 + (coord_str[i] - '0');
            ++i;
        }
        row--; // 转换为0-based
        
        return {row, col};
    }
};

} // namespace TinaXlsx 