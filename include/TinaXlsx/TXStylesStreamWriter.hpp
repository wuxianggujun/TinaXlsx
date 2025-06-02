//
// @file TXStylesStreamWriter.hpp
// @brief 样式表流式写入器 - 高性能样式XML生成
//

#pragma once

#include "TXTypes.hpp"
#include "TXResult.hpp"
#include "TXPugiStreamWriter.hpp"
#include <memory>
#include <string>

namespace TinaXlsx {

// 前向声明
class TXZipArchiveWriter;
class TXStyleManager;

/**
 * @brief 样式表流式写入器
 * 
 * 专门用于高性能生成styles.xml文件
 */
class TXStylesStreamWriter {
public:
    explicit TXStylesStreamWriter(size_t bufferSize = 64 * 1024);
    ~TXStylesStreamWriter() = default;

    /**
     * @brief 开始写入样式表文档
     */
    void startDocument();
    
    /**
     * @brief 写入数字格式部分
     */
    void writeNumberFormats(const TXStyleManager& styleManager);
    
    /**
     * @brief 写入字体部分
     */
    void writeFonts(const TXStyleManager& styleManager);
    
    /**
     * @brief 写入填充部分
     */
    void writeFills(const TXStyleManager& styleManager);
    
    /**
     * @brief 写入边框部分
     */
    void writeBorders(const TXStyleManager& styleManager);
    
    /**
     * @brief 写入单元格样式XF部分
     */
    void writeCellXfs(const TXStyleManager& styleManager);
    
    /**
     * @brief 结束文档
     */
    void endDocument();
    
    /**
     * @brief 写入到ZIP文件
     */
    TXResult<void> writeToZip(TXZipArchiveWriter& zipWriter, const std::string& partName);
    
    /**
     * @brief 重置写入器
     */
    void reset();

private:
    std::unique_ptr<TXBufferedXmlWriter> writer_;
    bool documentStarted_;
    
    // 辅助方法
    void writeXmlDeclaration();
    void writeStyleSheetStart();
    void writeStyleSheetEnd();
    
    // 具体写入方法
    void writeNumFmt(uint32_t numFmtId, const std::string& formatCode);
    void writeFont(const class TXFont& font);
    void writeFill(const class TXFill& fill);
    void writeBorder(const class TXBorder& border);
    void writeCellXf(const struct CellXF& xf);
    
    // XML转义
    std::string escapeXmlAttribute(const std::string& text);
};

/**
 * @brief 样式表写入器工厂
 */
class TXStylesWriterFactory {
public:
    /**
     * @brief 判断是否应该使用流式写入器
     * @param styleCount 样式数量
     * @return true表示应该使用流式写入器
     */
    static bool shouldUseStreamWriter(size_t styleCount);
    
    /**
     * @brief 创建最优的样式写入器
     * @param styleCount 预估样式数量
     * @return 写入器实例
     */
    static std::unique_ptr<TXStylesStreamWriter> createWriter(size_t styleCount);

private:
    static constexpr size_t STREAM_WRITER_THRESHOLD = 100; // 100个样式以上使用流式
};

} // namespace TinaXlsx
