//
// @file TXSIMDXmlEscaper.hpp
// @brief SIMD优化的XML转义器
//

#pragma once

#include <string>
#include <string_view>
#include <cstddef>

// 检测SIMD支持
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
 * @brief SIMD优化的XML转义器
 * 
 * 使用SIMD指令加速XML特殊字符的检测和转义，
 * 相比标准实现可以提升2-4倍性能
 */
class TXSIMDXmlEscaper {
public:
    /**
     * @brief 检查字符串是否需要XML转义
     * @param text 输入文本
     * @return true如果需要转义
     */
    static bool needsEscape(std::string_view text);

    /**
     * @brief XML转义（返回新字符串）
     * @param text 输入文本
     * @return 转义后的字符串
     */
    static std::string escape(std::string_view text);

    /**
     * @brief 高性能XML转义（写入到缓冲区）
     * @param text 输入文本
     * @param output 输出缓冲区
     * @param outputSize 缓冲区大小
     * @return 实际写入的字节数，-1表示缓冲区不足
     */
    static int escapeToBuffer(std::string_view text, char* output, size_t outputSize);

    /**
     * @brief 估算转义后的最大长度
     * @param originalLength 原始长度
     * @return 转义后的最大可能长度
     */
    static size_t estimateEscapedLength(size_t originalLength);

    /**
     * @brief 检测当前系统的SIMD支持级别
     */
    enum class SIMDLevel {
        None,
        SSE2,
        AVX2
    };
    
    static SIMDLevel detectSIMDSupport();

private:
    // SIMD实现
#ifdef TINAXLSX_HAS_SSE2
    static bool needsEscapeSSE2(const char* data, size_t length);
    static int escapeToBufferSSE2(const char* input, size_t inputLength, char* output, size_t outputSize);
#endif

#ifdef TINAXLSX_HAS_AVX2
    static bool needsEscapeAVX2(const char* data, size_t length);
    static int escapeToBufferAVX2(const char* input, size_t inputLength, char* output, size_t outputSize);
#endif

    // 标准实现（回退）
    static bool needsEscapeStandard(const char* data, size_t length);
    static int escapeToBufferStandard(const char* input, size_t inputLength, char* output, size_t outputSize);
    
    // 转义字符映射
    static const char* getEscapeSequence(char c);
    static size_t getEscapeSequenceLength(char c);
};

/**
 * @brief SIMD优化的XML写入器
 * 
 * 结合SIMD转义和缓冲写入的高性能XML写入器
 */
class TXSIMDXmlWriter {
public:
    explicit TXSIMDXmlWriter(size_t bufferSize = 64 * 1024);
    ~TXSIMDXmlWriter();

    /**
     * @brief 写入原始数据（不转义）
     */
    void writeRaw(const char* data, size_t length);
    void writeRaw(std::string_view text);

    /**
     * @brief 写入转义后的文本
     */
    void writeEscaped(std::string_view text);

    /**
     * @brief 获取缓冲区数据
     */
    std::string_view getBuffer() const;

    /**
     * @brief 清空缓冲区
     */
    void clear();

    /**
     * @brief 获取当前缓冲区使用量
     */
    size_t size() const { return bufferUsed_; }

private:
    char* buffer_;
    size_t bufferSize_;
    size_t bufferUsed_;
    
    void ensureCapacity(size_t additionalSize);
    void flushIfNeeded();
};

} // namespace TinaXlsx
