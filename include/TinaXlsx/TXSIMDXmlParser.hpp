//
// @file TXSIMDXmlParser.hpp
// @brief SIMD优化的XML解析器 - 专门针对Excel XML优化
//

#pragma once

#include <string>
#include <string_view>
#include <vector>
#include <memory>
#include <functional>
#include "TXTypes.hpp"
#include "TXCoordinate.hpp"

// SIMD支持检测
#ifdef _MSC_VER
    #include <intrin.h>
    #define TINAXLSX_HAS_SSE2 1
    #if defined(__AVX2__) || (defined(_MSC_VER) && _MSC_VER >= 1900)
        #define TINAXLSX_HAS_AVX2 1
    #endif
#elif defined(__GNUC__) || defined(__clang__)
    #ifdef __SSE2__
        #include <emmintrin.h>
        #define TINAXLSX_HAS_SSE2 1
    #endif
    #ifdef __AVX2__
        #include <immintrin.h>
        #define TINAXLSX_HAS_AVX2 1
    #endif
#endif

namespace TinaXlsx {

/**
 * @brief 轻量级XML节点（避免大量对象创建）
 */
struct FastXmlNode {
    std::string_view name;
    std::string_view value;
    std::string_view attributes;  // 原始属性字符串
    const char* start;            // 节点开始位置
    const char* end;              // 节点结束位置
    
    // 快速属性查找
    std::string_view getAttribute(std::string_view attrName) const;
    bool hasAttribute(std::string_view attrName) const;
};

/**
 * @brief SIMD优化的XML解析回调
 */
class ISIMDXmlCallback {
public:
    virtual ~ISIMDXmlCallback() = default;
    
    /**
     * @brief 处理单元格节点 (c)
     * @param node 单元格节点
     * @return true继续处理，false停止
     */
    virtual bool onCellNode(const FastXmlNode& node) = 0;
    
    /**
     * @brief 处理行节点 (row)
     * @param node 行节点
     * @return true继续处理，false停止
     */
    virtual bool onRowNode(const FastXmlNode& node) = 0;
    
    /**
     * @brief 处理共享字符串节点 (si)
     * @param index 字符串索引
     * @param text 字符串内容
     */
    virtual void onSharedString(u32 index, std::string_view text) = 0;
};

/**
 * @brief 🚀 SIMD优化的XML解析器
 * 
 * 专门针对Excel XML格式优化，使用SIMD指令加速：
 * - 快速标签查找
 * - 高效属性解析
 * - 批量字符处理
 */
class TXSIMDXmlParser {
public:
    /**
     * @brief SIMD支持级别
     */
    enum class SIMDLevel {
        None,
        SSE2,
        AVX2
    };
    
    /**
     * @brief 解析选项
     */
    struct ParseOptions {
        bool skipEmptyNodes = true;      // 跳过空节点
        bool validateXml = false;        // 是否验证XML格式
        size_t bufferSize = 64 * 1024;   // 缓冲区大小
        SIMDLevel forceSIMDLevel = SIMDLevel::None;  // 强制SIMD级别
    };
    
    TXSIMDXmlParser();
    ~TXSIMDXmlParser();
    
    /**
     * @brief 检测SIMD支持
     */
    static SIMDLevel detectSIMDSupport();
    
    /**
     * @brief 设置解析选项
     */
    void setOptions(const ParseOptions& options) { options_ = options; }
    
    /**
     * @brief 🚀 高性能解析工作表XML
     * @param xmlContent XML内容
     * @param callback 回调接口
     * @return 解析的节点数量
     */
    size_t parseWorksheet(std::string_view xmlContent, ISIMDXmlCallback& callback);
    
    /**
     * @brief 🚀 高性能解析共享字符串XML
     * @param xmlContent XML内容
     * @param callback 回调接口
     * @return 解析的字符串数量
     */
    size_t parseSharedStrings(std::string_view xmlContent, ISIMDXmlCallback& callback);
    
    /**
     * @brief 获取解析统计信息
     */
    struct ParseStats {
        size_t totalNodes = 0;
        size_t totalAttributes = 0;
        double parseTimeMs = 0.0;
        SIMDLevel usedSIMDLevel = SIMDLevel::None;
    };
    
    const ParseStats& getStats() const { return stats_; }

private:
    ParseOptions options_;
    ParseStats stats_;
    
    // SIMD实现
#ifdef TINAXLSX_HAS_SSE2
    size_t parseWorksheetSSE2(std::string_view xmlContent, ISIMDXmlCallback& callback);
    const char* findTagSSE2(const char* data, size_t length, const char* tag, size_t tagLen);
#endif

#ifdef TINAXLSX_HAS_AVX2
    size_t parseWorksheetAVX2(std::string_view xmlContent, ISIMDXmlCallback& callback);
    const char* findTagAVX2(const char* data, size_t length, const char* tag, size_t tagLen);
#endif

public:
    // 标准实现（回退）- 设为public供工作表解析器使用
    size_t parseWorksheetStandard(std::string_view xmlContent, ISIMDXmlCallback& callback);
    const char* findTagStandard(const char* data, size_t length, const char* tag, size_t tagLen);

private:
    
    // 辅助方法
    FastXmlNode parseNode(const char* start, const char* end);
    std::string_view extractNodeValue(const char* start, const char* end);
    std::string_view extractAttributes(const char* start, const char* end);
    
    // 快速字符查找
    const char* findChar(const char* data, size_t length, char target);
    const char* findAnyChar(const char* data, size_t length, const char* targets, size_t targetCount);
};

/**
 * @brief 🚀 高性能工作表解析器（使用SIMD）
 */
class TXSIMDWorksheetParser : public ISIMDXmlCallback {
public:
    explicit TXSIMDWorksheetParser(class TXSheet* sheet);
    
    /**
     * @brief 解析工作表
     * @param xmlContent XML内容
     * @return 解析的单元格数量
     */
    size_t parse(std::string_view xmlContent);
    
    // ISIMDXmlCallback 实现
    bool onCellNode(const FastXmlNode& node) override;
    bool onRowNode(const FastXmlNode& node) override;
    void onSharedString(u32 index, std::string_view text) override;
    
    /**
     * @brief 获取解析统计
     */
    struct WorksheetStats {
        size_t totalRows = 0;
        size_t totalCells = 0;
        size_t emptySkipped = 0;
        double parseTimeMs = 0.0;
    };
    
    const WorksheetStats& getStats() const { return stats_; }

private:
    class TXSheet* sheet_;
    TXSIMDXmlParser parser_;
    WorksheetStats stats_;
    
    // 批量单元格缓存
    struct CellBatch {
        std::vector<TXCoordinate> coordinates;
        std::vector<std::string> values;
        std::vector<u32> styleIndices;

        CellBatch() {
            coordinates.reserve(1000);
            values.reserve(1000);
            styleIndices.reserve(1000);
        }

        void reserve(size_t size) {
            coordinates.reserve(size);
            values.reserve(size);
            styleIndices.reserve(size);
        }

        void clear() {
            coordinates.clear();
            values.clear();
            styleIndices.clear();
        }

        size_t size() const { return coordinates.size(); }
    };
    
    CellBatch cellBatch_;
    static constexpr size_t BATCH_SIZE = 1000;
    
    void flushBatch();
    void addCellToBatch(const TXCoordinate& coord, const std::string& value, u32 styleIndex = 0);
    
    // 快速坐标解析
    TXCoordinate parseCoordinate(std::string_view ref);
    std::string parseValue(const FastXmlNode& cellNode);
    u32 parseStyleIndex(const FastXmlNode& cellNode);
};

} // namespace TinaXlsx
