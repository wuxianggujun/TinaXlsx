//
// @file TXFastXmlWriter.hpp
// @brief 超高性能XML写入器 - 专门为Excel文件优化
//

#pragma once

#include "TinaXlsx/TXUnifiedMemoryManager.hpp"
#include "TinaXlsx/TXHighPerformanceLogger.hpp"
#include <vector>
#include <string>
#include <string_view>
#include <cstring>
#include <cstdio>

namespace TinaXlsx {

/**
 * @brief 🚀 超高性能XML写入器
 * 
 * 专门为Excel XML文件优化的写入器：
 * - 零拷贝字符串操作
 * - 预编译的XML模板
 * - 批量写入优化
 * - SIMD优化的字符串转义
 */
class TXFastXmlWriter {
public:
    explicit TXFastXmlWriter(TXUnifiedMemoryManager& memory_manager, size_t initial_capacity = 1024 * 1024);
    ~TXFastXmlWriter() = default;

    // 禁用拷贝，允许移动
    TXFastXmlWriter(const TXFastXmlWriter&) = delete;
    TXFastXmlWriter& operator=(const TXFastXmlWriter&) = delete;
    TXFastXmlWriter(TXFastXmlWriter&&) = default;
    TXFastXmlWriter& operator=(TXFastXmlWriter&&) = default;

    /**
     * @brief 🚀 预编译的XML模板
     */
    struct XmlTemplates {
        static constexpr std::string_view XML_DECLARATION = 
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
        static constexpr std::string_view WORKSHEET_START = 
            "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">";
        static constexpr std::string_view WORKSHEET_END = "</worksheet>";
        static constexpr std::string_view SHEETDATA_START = "<sheetData>";
        static constexpr std::string_view SHEETDATA_END = "</sheetData>";
        static constexpr std::string_view ROW_END = "</row>";
    };

    /**
     * @brief 🚀 快速写入方法
     */
    void writeXmlDeclaration() { writeConstant(XmlTemplates::XML_DECLARATION); }
    void writeWorksheetStart() { writeConstant(XmlTemplates::WORKSHEET_START); }
    void writeWorksheetEnd() { writeConstant(XmlTemplates::WORKSHEET_END); }
    void writeSheetDataStart() { writeConstant(XmlTemplates::SHEETDATA_START); }
    void writeSheetDataEnd() { writeConstant(XmlTemplates::SHEETDATA_END); }

    /**
     * @brief 🚀 高性能行写入
     */
    void writeRowStart(uint32_t row_number);
    void writeRowEnd() { writeConstant(XmlTemplates::ROW_END); }

    /**
     * @brief 🚀 高性能单元格写入
     */
    void writeNumberCell(std::string_view coord, double value);
    void writeStringCell(std::string_view coord, std::string_view value);
    void writeInlineStringCell(std::string_view coord, std::string_view value);

    /**
     * @brief 🚀 批量单元格写入
     */
    void writeNumberCellsBatch(const std::vector<std::string>& coords, 
                               const double* values, size_t count);

    /**
     * @brief 🚀 原始数据写入
     */
    void writeRaw(const void* data, size_t size);
    void writeConstant(std::string_view str);
    void writeString(const std::string& str) { writeConstant(str); }

    /**
     * @brief 🚀 结果获取
     */
    std::vector<uint8_t> getResult() &&;
    const std::vector<uint8_t>& getBuffer() const { return buffer_; }
    size_t size() const { return current_pos_; }

    /**
     * @brief 🚀 性能优化
     */
    void reserve(size_t capacity);
    void clear();

private:
    TXUnifiedMemoryManager& memory_manager_;
    std::vector<uint8_t> buffer_;
    size_t current_pos_ = 0;

    // 🚀 内部优化方法
    void ensureCapacity(size_t additional_size);
    void writeEscapedString(std::string_view str);
    
    // 🚀 快速数值转换
    void writeDouble(double value);
    void writeUint32(uint32_t value);

    // 🚀 预编译的单元格模板
    static constexpr size_t MAX_COORD_LENGTH = 16;
    static constexpr size_t MAX_NUMBER_LENGTH = 32;
    
    // 缓存区用于避免临时分配
    thread_local static char coord_buffer_[MAX_COORD_LENGTH];
    thread_local static char number_buffer_[MAX_NUMBER_LENGTH];
};

/**
 * @brief 🚀 坐标转换工具
 */
class TXCoordConverter {
public:
    static void rowColToString(uint32_t row, uint32_t col, char* buffer, size_t buffer_size);
    static std::string rowColToString(uint32_t row, uint32_t col);
    
private:
    static void columnToLetters(uint32_t col, char* buffer, size_t& pos);
};

/**
 * @brief 🚀 快速数值转换工具
 */
class TXFastNumberConverter {
public:
    static size_t doubleToString(double value, char* buffer, size_t buffer_size);
    static size_t uint32ToString(uint32_t value, char* buffer, size_t buffer_size);
    
private:
    static constexpr double POWERS_OF_10[] = {
        1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9,
        1e10, 1e11, 1e12, 1e13, 1e14, 1e15
    };
};

} // namespace TinaXlsx
