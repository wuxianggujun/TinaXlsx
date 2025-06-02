#pragma once

#include <string>
#include <vector>
#include <fstream>
#include <sstream>
#include <memory>
#include "TXTypes.hpp"

namespace TinaXlsx {

/**
 * @brief 高性能流式XML写入器
 * 
 * 专门为大量数据的XML生成优化，避免DOM树构建的内存开销
 * 特点：
 * - 流式写入，内存占用恒定
 * - 缓冲区优化，减少I/O次数
 * - 专门针对Excel XML格式优化
 */
class TXStreamXmlWriter {
public:
    /**
     * @brief 构造函数
     * @param bufferSize 内部缓冲区大小（字节）
     */
    explicit TXStreamXmlWriter(std::size_t bufferSize = 64 * 1024);
    
    /**
     * @brief 析构函数，自动刷新缓冲区
     */
    ~TXStreamXmlWriter();
    
    // 禁用拷贝，支持移动
    TXStreamXmlWriter(const TXStreamXmlWriter&) = delete;
    TXStreamXmlWriter& operator=(const TXStreamXmlWriter&) = delete;
    TXStreamXmlWriter(TXStreamXmlWriter&&) = default;
    TXStreamXmlWriter& operator=(TXStreamXmlWriter&&) = default;
    
    /**
     * @brief 开始写入到字符串
     * @return 成功返回true
     */
    bool startStringOutput();
    
    /**
     * @brief 开始写入到文件
     * @param filename 文件名
     * @return 成功返回true
     */
    bool startFileOutput(const std::string& filename);
    
    /**
     * @brief 开始写入到内存缓冲区
     * @param buffer 输出缓冲区
     * @return 成功返回true
     */
    bool startBufferOutput(std::vector<uint8_t>& buffer);
    
    /**
     * @brief 写入XML声明
     * @param encoding 编码格式
     */
    void writeXmlDeclaration(const std::string& encoding = "UTF-8");
    
    /**
     * @brief 开始元素
     * @param name 元素名
     */
    void startElement(const std::string& name);
    
    /**
     * @brief 添加属性（必须在startElement之后，endElement或writeText之前调用）
     * @param name 属性名
     * @param value 属性值
     */
    void addAttribute(const std::string& name, const std::string& value);
    
    /**
     * @brief 结束元素
     * @param name 元素名（用于验证）
     */
    void endElement(const std::string& name);
    
    /**
     * @brief 写入文本内容
     * @param text 文本内容
     * @param escapeXml 是否转义XML特殊字符
     */
    void writeText(const std::string& text, bool escapeXml = true);
    
    /**
     * @brief 写入完整的简单元素
     * @param name 元素名
     * @param text 文本内容
     * @param attributes 属性列表
     */
    void writeSimpleElement(const std::string& name, const std::string& text = "",
                           const std::vector<std::pair<std::string, std::string>>& attributes = {});
    
    /**
     * @brief 刷新缓冲区
     */
    void flush();
    
    /**
     * @brief 完成写入并获取结果
     * @return 如果是字符串输出，返回XML字符串；否则返回空字符串
     */
    std::string finish();
    
    /**
     * @brief 获取当前缓冲区大小
     * @return 缓冲区中的字节数
     */
    std::size_t getBufferSize() const { return buffer_.size(); }
    
    /**
     * @brief 获取写入的总字节数
     * @return 总字节数
     */
    std::size_t getTotalBytesWritten() const { return totalBytesWritten_; }

private:
    enum class OutputMode {
        String,
        File,
        Buffer
    };
    
    OutputMode outputMode_;
    std::string buffer_;
    std::size_t bufferCapacity_;
    std::size_t totalBytesWritten_;
    
    // 输出目标
    std::ostringstream stringStream_;
    std::ofstream fileStream_;
    std::vector<uint8_t>* outputBuffer_;
    
    // 状态管理
    std::vector<std::string> elementStack_;
    bool elementStarted_;
    bool attributesAllowed_;
    
    /**
     * @brief 内部写入方法
     * @param data 要写入的数据
     */
    void writeInternal(const std::string& data);
    
    /**
     * @brief 转义XML特殊字符
     * @param text 原始文本
     * @return 转义后的文本
     */
    std::string escapeXml(const std::string& text) const;
    
    /**
     * @brief 完成当前元素的开始标签
     */
    void finishElementStart();
};

/**
 * @brief Excel工作表XML流式写入器
 * 
 * 专门针对Excel工作表XML格式优化的写入器
 */
class TXWorksheetStreamWriter {
public:
    explicit TXWorksheetStreamWriter(std::size_t bufferSize = 128 * 1024);
    
    /**
     * @brief 开始写入工作表
     * @param usedRange 使用的范围
     */
    void startWorksheet(const std::string& usedRangeRef);
    
    /**
     * @brief 开始写入行
     * @param rowNumber 行号
     */
    void startRow(u32 rowNumber);
    
    /**
     * @brief 写入单元格
     * @param cellRef 单元格引用（如A1）
     * @param value 单元格值
     * @param cellType 单元格类型（可选）
     * @param styleIndex 样式索引（可选）
     */
    void writeCell(const std::string& cellRef, const std::string& value,
                   const std::string& cellType = "", u32 styleIndex = 0);
    
    /**
     * @brief 结束当前行
     */
    void endRow();
    
    /**
     * @brief 结束工作表
     */
    void endWorksheet();
    
    /**
     * @brief 获取生成的XML
     * @return XML字符串
     */
    std::string getXml();

private:
    TXStreamXmlWriter writer_;
    bool inRow_;
};

} // namespace TinaXlsx
