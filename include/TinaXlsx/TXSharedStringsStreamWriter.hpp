//
// @file TXSharedStringsStreamWriter.hpp
// @brief 共享字符串流式写入器 - 针对大量字符串优化
//

#pragma once

#include "TXPugiStreamWriter.hpp"
#include <unordered_set>
#include <string>

namespace TinaXlsx {

/**
 * @brief 共享字符串流式写入器
 * 
 * 专门针对大量共享字符串的高性能写入器
 */
class TXSharedStringsStreamWriter {
public:
    explicit TXSharedStringsStreamWriter(size_t bufferSize = 128 * 1024);
    
    /**
     * @brief 开始写入共享字符串文档
     * @param estimatedCount 预估字符串数量
     */
    void startDocument(size_t estimatedCount);
    
    /**
     * @brief 写入一个共享字符串
     * @param text 字符串内容
     * @param preserveSpace 是否保留空格
     */
    void writeString(const std::string& text, bool preserveSpace = false);
    
    /**
     * @brief 结束文档
     */
    void endDocument();
    
    /**
     * @brief 写入到ZIP文件
     * @param zipWriter ZIP写入器
     * @param partName 部件名称
     * @return 操作结果
     */
    TXResult<void> writeToZip(TXZipArchiveWriter& zipWriter, const std::string& partName);
    
    /**
     * @brief 获取写入的字符串数量
     */
    size_t getStringCount() const { return stringCount_; }
    
    /**
     * @brief 重置写入器
     */
    void reset();

private:
    std::unique_ptr<TXBufferedXmlWriter> writer_;
    size_t stringCount_;
    bool documentStarted_;

    /**
     * @brief 🚀 性能优化：直接写入转义文本，避免创建临时字符串
     */
    void writeEscapedXmlText(const std::string& text);

    /**
     * @brief XML文本转义（兼容性方法）
     */
    std::string escapeXmlText(const std::string& text);
};

/**
 * @brief 共享字符串写入器工厂
 */
class TXSharedStringsWriterFactory {
public:
    /**
     * @brief 创建写入器
     * @param stringCount 预估字符串数量
     * @return 写入器实例
     */
    static std::unique_ptr<TXSharedStringsStreamWriter> createWriter(size_t stringCount);
};

} // namespace TinaXlsx
