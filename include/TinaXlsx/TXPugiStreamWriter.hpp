//
// @file TXPugiStreamWriter.hpp
// @brief 基于pugixml xml_writer的高性能流式XML写入器
//

#pragma once

#include <pugixml.hpp>
#include <vector>
#include <string>
#include <memory>
#include "TXTypes.hpp"
#include "TXResult.hpp"

namespace TinaXlsx {

// 前向声明
class TXZipArchiveWriter;

/**
 * @brief 基于pugixml xml_writer的缓冲写入器
 * 
 * 实现pugixml的xml_writer接口，将XML数据写入到内存缓冲区
 */
class TXBufferedXmlWriter : public pugi::xml_writer {
public:
    explicit TXBufferedXmlWriter(size_t bufferSize = 256 * 1024);
    
    // 实现xml_writer接口
    void write(const void* data, size_t size) override;
    
    // 获取缓冲区数据
    const std::vector<uint8_t>& getBuffer() const { return buffer_; }
    
    // 清空缓冲区
    void clear();
    
    // 获取写入的总字节数
    size_t getTotalBytesWritten() const { return totalBytes_; }

private:
    std::vector<uint8_t> buffer_;
    size_t totalBytes_;
};

/**
 * @brief 高性能Excel工作表流式写入器
 * 
 * 使用pugixml的流式API直接生成XML，避免DOM树构建的开销
 */
class TXPugiWorksheetWriter {
public:
    explicit TXPugiWorksheetWriter(size_t bufferSize = 256 * 1024);
    
    /**
     * @brief 开始写入工作表
     * @param usedRangeRef 使用范围引用（如A1:Z100）
     * @param hasCustomColumns 是否有自定义列宽
     */
    void startWorksheet(const std::string& usedRangeRef, bool hasCustomColumns = false);
    
    /**
     * @brief 写入列宽信息
     * @param columnIndex 列索引
     * @param width 列宽
     */
    void writeColumnWidth(u32 columnIndex, double width);
    
    /**
     * @brief 开始写入工作表数据部分
     */
    void startSheetData();
    
    /**
     * @brief 开始写入行
     * @param rowNumber 行号（1-based）
     */
    void startRow(u32 rowNumber);
    
    /**
     * @brief 写入字符串单元格（内联方式）
     * @param cellRef 单元格引用（如A1）
     * @param value 字符串值
     * @param styleIndex 样式索引（可选）
     */
    void writeCellInlineString(const std::string& cellRef, const std::string& value, u32 styleIndex = 0);
    
    /**
     * @brief 写入字符串单元格（共享字符串方式）
     * @param cellRef 单元格引用（如A1）
     * @param stringIndex 共享字符串索引
     * @param styleIndex 样式索引（可选）
     */
    void writeCellSharedString(const std::string& cellRef, u32 stringIndex, u32 styleIndex = 0);
    
    /**
     * @brief 写入数值单元格
     * @param cellRef 单元格引用（如A1）
     * @param value 数值
     * @param styleIndex 样式索引（可选）
     */
    void writeCellNumber(const std::string& cellRef, double value, u32 styleIndex = 0);
    
    /**
     * @brief 写入整数单元格
     * @param cellRef 单元格引用（如A1）
     * @param value 整数值
     * @param styleIndex 样式索引（可选）
     */
    void writeCellInteger(const std::string& cellRef, int64_t value, u32 styleIndex = 0);
    
    /**
     * @brief 写入布尔单元格
     * @param cellRef 单元格引用（如A1）
     * @param value 布尔值
     * @param styleIndex 样式索引（可选）
     */
    void writeCellBoolean(const std::string& cellRef, bool value, u32 styleIndex = 0);
    
    /**
     * @brief 结束当前行
     */
    void endRow();
    
    /**
     * @brief 结束工作表数据部分
     */
    void endSheetData();
    
    /**
     * @brief 结束工作表
     */
    void endWorksheet();
    
    /**
     * @brief 写入到ZIP文件
     * @param zipWriter ZIP写入器
     * @param partName 部件名称
     * @return 操作结果
     */
    TXResult<void> writeToZip(TXZipArchiveWriter& zipWriter, const std::string& partName);
    
    /**
     * @brief 获取生成的XML数据
     * @return XML数据缓冲区
     */
    const std::vector<uint8_t>& getXmlData() const;
    
    /**
     * @brief 重置写入器状态
     */
    void reset();
    
    /**
     * @brief 获取性能统计信息
     */
    struct PerformanceStats {
        size_t totalBytesWritten = 0;
        size_t cellsWritten = 0;
        size_t rowsWritten = 0;
    };
    
    PerformanceStats getStats() const { return stats_; }

private:
    std::unique_ptr<TXBufferedXmlWriter> writer_;
    std::unique_ptr<pugi::xml_document> doc_;
    PerformanceStats stats_;
    bool worksheetStarted_;
    bool sheetDataStarted_;
    bool rowStarted_;
    
    // 辅助方法
    void writeXmlDeclaration();
    void writeCellStart(const std::string& cellRef, const std::string& cellType, u32 styleIndex);
    void writeCellEnd();
    std::string escapeXmlText(const std::string& text);
};

/**
 * @brief 工作表流式写入器工厂
 * 
 * 根据数据量自动选择最优的写入策略
 */
class TXWorksheetWriterFactory {
public:
    /**
     * @brief 创建工作表写入器
     * @param estimatedCells 预估单元格数量
     * @return 写入器实例
     */
    static std::unique_ptr<TXPugiWorksheetWriter> createWriter(size_t estimatedCells);
    
    /**
     * @brief 判断是否应该使用流式写入器
     * @param estimatedCells 预估单元格数量
     * @return true表示应该使用流式写入器
     */
    static bool shouldUseStreamWriter(size_t estimatedCells);
    
private:
    static constexpr size_t STREAM_WRITER_THRESHOLD = 5000;
};

} // namespace TinaXlsx
