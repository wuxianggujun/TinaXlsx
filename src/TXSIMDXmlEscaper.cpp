//
// @file TXSIMDXmlEscaper.cpp
// @brief SIMD优化的XML转义器实现
//

#include "TinaXlsx/TXSIMDXmlEscaper.hpp"
#include <cstring>
#include <algorithm>

namespace TinaXlsx {

// 转义字符映射
const char* TXSIMDXmlEscaper::getEscapeSequence(char c) {
    switch (c) {
        case '<': return "&lt;";
        case '>': return "&gt;";
        case '&': return "&amp;";
        case '"': return "&quot;";
        case '\'': return "&apos;";
        default: return nullptr;
    }
}

size_t TXSIMDXmlEscaper::getEscapeSequenceLength(char c) {
    switch (c) {
        case '<': return 4;  // &lt;
        case '>': return 4;  // &gt;
        case '&': return 5;  // &amp;
        case '"': return 6;  // &quot;
        case '\'': return 6; // &apos;
        default: return 0;
    }
}

TXSIMDXmlEscaper::SIMDLevel TXSIMDXmlEscaper::detectSIMDSupport() {
#ifdef TINAXLSX_HAS_AVX2
    return SIMDLevel::AVX2;
#elif defined(TINAXLSX_HAS_SSE2)
    return SIMDLevel::SSE2;
#else
    return SIMDLevel::None;
#endif
}

bool TXSIMDXmlEscaper::needsEscape(std::string_view text) {
    const char* data = text.data();
    size_t length = text.length();

#ifdef TINAXLSX_HAS_AVX2
    return needsEscapeAVX2(data, length);
#elif defined(TINAXLSX_HAS_SSE2)
    return needsEscapeSSE2(data, length);
#else
    return needsEscapeStandard(data, length);
#endif
}

std::string TXSIMDXmlEscaper::escape(std::string_view text) {
    if (!needsEscape(text)) {
        return std::string(text);
    }

    // 估算输出大小
    size_t maxOutputSize = estimateEscapedLength(text.length());
    std::string result(maxOutputSize, '\0');

    int actualLength = escapeToBuffer(text, result.data(), maxOutputSize);
    if (actualLength >= 0) {
        result.resize(actualLength);
        return result;
    }

    // 回退到标准实现
    return std::string(text); // 简化处理
}

int TXSIMDXmlEscaper::escapeToBuffer(std::string_view text, char* output, size_t outputSize) {
    const char* input = text.data();
    size_t inputLength = text.length();

#ifdef TINAXLSX_HAS_AVX2
    return escapeToBufferAVX2(input, inputLength, output, outputSize);
#elif defined(TINAXLSX_HAS_SSE2)
    return escapeToBufferSSE2(input, inputLength, output, outputSize);
#else
    return escapeToBufferStandard(input, inputLength, output, outputSize);
#endif
}

size_t TXSIMDXmlEscaper::estimateEscapedLength(size_t originalLength) {
    // 最坏情况：所有字符都是&，转义为&amp;（5倍）
    return originalLength * 6;
}

// 标准实现
bool TXSIMDXmlEscaper::needsEscapeStandard(const char* data, size_t length) {
    for (size_t i = 0; i < length; ++i) {
        char c = data[i];
        if (c == '<' || c == '>' || c == '&' || c == '"' || c == '\'') {
            return true;
        }
    }
    return false;
}

int TXSIMDXmlEscaper::escapeToBufferStandard(const char* input, size_t inputLength, char* output, size_t outputSize) {
    size_t outputPos = 0;
    
    for (size_t i = 0; i < inputLength; ++i) {
        char c = input[i];
        const char* escapeSeq = getEscapeSequence(c);
        
        if (escapeSeq) {
            size_t escapeLen = getEscapeSequenceLength(c);
            if (outputPos + escapeLen > outputSize) {
                return -1; // 缓冲区不足
            }
            memcpy(output + outputPos, escapeSeq, escapeLen);
            outputPos += escapeLen;
        } else {
            if (outputPos >= outputSize) {
                return -1; // 缓冲区不足
            }
            output[outputPos++] = c;
        }
    }
    
    return static_cast<int>(outputPos);
}

#ifdef TINAXLSX_HAS_SSE2
bool TXSIMDXmlEscaper::needsEscapeSSE2(const char* data, size_t length) {
    // 🚀 SIMD优化：使用SSE2同时检查16个字符
    const __m128i lt = _mm_set1_epi8('<');
    const __m128i gt = _mm_set1_epi8('>');
    const __m128i amp = _mm_set1_epi8('&');
    const __m128i quote = _mm_set1_epi8('"');
    const __m128i apos = _mm_set1_epi8('\'');
    
    size_t simdLength = length & ~15; // 16字节对齐
    
    for (size_t i = 0; i < simdLength; i += 16) {
        __m128i chunk = _mm_loadu_si128(reinterpret_cast<const __m128i*>(data + i));
        
        __m128i cmp_lt = _mm_cmpeq_epi8(chunk, lt);
        __m128i cmp_gt = _mm_cmpeq_epi8(chunk, gt);
        __m128i cmp_amp = _mm_cmpeq_epi8(chunk, amp);
        __m128i cmp_quote = _mm_cmpeq_epi8(chunk, quote);
        __m128i cmp_apos = _mm_cmpeq_epi8(chunk, apos);
        
        __m128i result = _mm_or_si128(_mm_or_si128(cmp_lt, cmp_gt), 
                                     _mm_or_si128(_mm_or_si128(cmp_amp, cmp_quote), cmp_apos));
        
        if (_mm_movemask_epi8(result) != 0) {
            return true;
        }
    }
    
    // 处理剩余字符
    for (size_t i = simdLength; i < length; ++i) {
        char c = data[i];
        if (c == '<' || c == '>' || c == '&' || c == '"' || c == '\'') {
            return true;
        }
    }
    
    return false;
}

int TXSIMDXmlEscaper::escapeToBufferSSE2(const char* input, size_t inputLength, char* output, size_t outputSize) {
    // 对于SSE2，复杂的转义逻辑回退到标准实现
    // 实际项目中可以实现更复杂的SIMD转义逻辑
    return escapeToBufferStandard(input, inputLength, output, outputSize);
}
#endif

#ifdef TINAXLSX_HAS_AVX2
bool TXSIMDXmlEscaper::needsEscapeAVX2(const char* data, size_t length) {
    // 🚀 SIMD优化：使用AVX2同时检查32个字符
    const __m256i lt = _mm256_set1_epi8('<');
    const __m256i gt = _mm256_set1_epi8('>');
    const __m256i amp = _mm256_set1_epi8('&');
    const __m256i quote = _mm256_set1_epi8('"');
    const __m256i apos = _mm256_set1_epi8('\'');
    
    size_t simdLength = length & ~31; // 32字节对齐
    
    for (size_t i = 0; i < simdLength; i += 32) {
        __m256i chunk = _mm256_loadu_si256(reinterpret_cast<const __m256i*>(data + i));
        
        __m256i cmp_lt = _mm256_cmpeq_epi8(chunk, lt);
        __m256i cmp_gt = _mm256_cmpeq_epi8(chunk, gt);
        __m256i cmp_amp = _mm256_cmpeq_epi8(chunk, amp);
        __m256i cmp_quote = _mm256_cmpeq_epi8(chunk, quote);
        __m256i cmp_apos = _mm256_cmpeq_epi8(chunk, apos);
        
        __m256i result = _mm256_or_si256(_mm256_or_si256(cmp_lt, cmp_gt), 
                                        _mm256_or_si256(_mm256_or_si256(cmp_amp, cmp_quote), cmp_apos));
        
        if (_mm256_movemask_epi8(result) != 0) {
            return true;
        }
    }
    
    // 处理剩余字符
    for (size_t i = simdLength; i < length; ++i) {
        char c = data[i];
        if (c == '<' || c == '>' || c == '&' || c == '"' || c == '\'') {
            return true;
        }
    }
    
    return false;
}

int TXSIMDXmlEscaper::escapeToBufferAVX2(const char* input, size_t inputLength, char* output, size_t outputSize) {
    // 对于AVX2，复杂的转义逻辑回退到标准实现
    // 实际项目中可以实现更复杂的SIMD转义逻辑
    return escapeToBufferStandard(input, inputLength, output, outputSize);
}
#endif

} // namespace TinaXlsx
