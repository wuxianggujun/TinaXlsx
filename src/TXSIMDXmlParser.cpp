//
// @file TXSIMDXmlParser.cpp
// @brief SIMD优化的XML解析器实现
//

#include "TinaXlsx/TXSIMDXmlParser.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCoordinate.hpp"
#include "TinaXlsx/TXCellManager.hpp"
#include "TinaXlsx/TXCompactCell.hpp"  // 包含新的 TXStringPool
#include <chrono>
#include <cstring>
#include <algorithm>

namespace TinaXlsx {

// ==================== FastXmlNode 实现 ====================

std::string_view FastXmlNode::getAttribute(std::string_view attrName) const {
    if (attributes.empty()) return {};
    
    // 🚀 快速属性查找，避免正则表达式
    size_t pos = attributes.find(attrName);
    if (pos == std::string_view::npos) return {};
    
    // 查找 = 号
    pos = attributes.find('=', pos + attrName.length());
    if (pos == std::string_view::npos) return {};
    
    // 跳过 = 和引号
    pos++;
    while (pos < attributes.length() && (attributes[pos] == ' ' || attributes[pos] == '\t')) pos++;
    if (pos >= attributes.length() || (attributes[pos] != '"' && attributes[pos] != '\'')) return {};
    
    char quote = attributes[pos++];
    size_t start = pos;
    
    // 查找结束引号
    size_t end = attributes.find(quote, start);
    if (end == std::string_view::npos) return {};
    
    return attributes.substr(start, end - start);
}

bool FastXmlNode::hasAttribute(std::string_view attrName) const {
    return attributes.find(attrName) != std::string_view::npos;
}

// ==================== TXSIMDXmlParser 实现 ====================

TXSIMDXmlParser::TXSIMDXmlParser() = default;
TXSIMDXmlParser::~TXSIMDXmlParser() = default;

TXSIMDXmlParser::SIMDLevel TXSIMDXmlParser::detectSIMDSupport() {
#ifdef TINAXLSX_HAS_AVX2
    return SIMDLevel::AVX2;
#elif defined(TINAXLSX_HAS_SSE2)
    return SIMDLevel::SSE2;
#else
    return SIMDLevel::None;
#endif
}

size_t TXSIMDXmlParser::parseWorksheet(std::string_view xmlContent, ISIMDXmlCallback& callback) {
    auto startTime = std::chrono::high_resolution_clock::now();
    
    // 重置统计信息
    stats_ = ParseStats{};
    
    size_t result = 0;
    
    // 根据SIMD支持选择实现
    SIMDLevel level = (options_.forceSIMDLevel != SIMDLevel::None) ? 
                      options_.forceSIMDLevel : detectSIMDSupport();
    
    stats_.usedSIMDLevel = level;
    
    switch (level) {
#ifdef TINAXLSX_HAS_AVX2
        case SIMDLevel::AVX2:
            result = parseWorksheetAVX2(xmlContent, callback);
            break;
#endif
#ifdef TINAXLSX_HAS_SSE2
        case SIMDLevel::SSE2:
            result = parseWorksheetSSE2(xmlContent, callback);
            break;
#endif
        default:
            result = parseWorksheetStandard(xmlContent, callback);
            break;
    }
    
    auto endTime = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(endTime - startTime);
    stats_.parseTimeMs = duration.count() / 1000.0;
    stats_.totalNodes = result;
    
    return result;
}

size_t TXSIMDXmlParser::parseSharedStrings(std::string_view xmlContent, ISIMDXmlCallback& callback) {
    // 简化的共享字符串解析
    size_t count = 0;
    const char* data = xmlContent.data();
    size_t length = xmlContent.length();
    
    // 查找所有 <si> 节点
    const char* siTag = "<si>";
    const char* siEndTag = "</si>";
    const char* pos = data;
    const char* end = data + length;
    
    while (pos < end) {
        pos = findTagStandard(pos, end - pos, siTag, 4);
        if (!pos) break;
        
        const char* siEnd = findTagStandard(pos, end - pos, siEndTag, 5);
        if (!siEnd) break;
        
        // 提取文本内容
        const char* tStart = findTagStandard(pos, siEnd - pos, "<t>", 3);
        if (tStart) {
            tStart += 3; // 跳过 <t>
            const char* tEnd = findTagStandard(tStart, siEnd - tStart, "</t>", 4);
            if (tEnd) {
                std::string_view text(tStart, tEnd - tStart);
                callback.onSharedString(count, text);
            }
        }
        
        count++;
        pos = siEnd + 5; // 跳过 </si>
    }
    
    return count;
}

// ==================== 标准实现 ====================

size_t TXSIMDXmlParser::parseWorksheetStandard(std::string_view xmlContent, ISIMDXmlCallback& callback) {
    size_t nodeCount = 0;
    const char* data = xmlContent.data();
    size_t length = xmlContent.length();

    // 🚀 优化：使用更快的字符串查找
    const char* sheetDataStart = std::strstr(data, "<sheetData");
    if (!sheetDataStart) return 0;

    // 查找 <sheetData> 结束
    const char* sheetDataEnd = std::strstr(sheetDataStart, "</sheetData>");
    if (!sheetDataEnd) sheetDataEnd = data + length;

    // 🚀 批量解析优化：预分配向量避免重复分配
    std::vector<FastXmlNode> cellBatch;
    cellBatch.reserve(1000);

    // 解析所有行
    const char* pos = sheetDataStart;
    while (pos < sheetDataEnd) {
        // 🚀 优化：使用更快的标签查找
        pos = std::strstr(pos, "<row");
        if (!pos || pos >= sheetDataEnd) break;

        // 查找行结束
        const char* rowEnd = std::strstr(pos, "</row>");
        if (!rowEnd) break;

        // 解析行节点
        FastXmlNode rowNode = parseNode(pos, rowEnd + 6);
        if (!callback.onRowNode(rowNode)) break;

        // 🚀 批量解析单元格
        cellBatch.clear();
        const char* cellPos = pos;
        while (cellPos < rowEnd) {
            cellPos = std::strstr(cellPos, "<c ");
            if (!cellPos || cellPos >= rowEnd) break;

            // 查找单元格结束
            const char* cellEnd = std::strstr(cellPos, "</c>");
            if (!cellEnd) {
                // 自闭合标签
                cellEnd = std::strchr(cellPos, '>');
                if (!cellEnd) break;
                cellEnd++;
            } else {
                cellEnd += 4;
            }

            // 🚀 批量收集单元格节点
            FastXmlNode cellNode = parseNode(cellPos, cellEnd);
            cellBatch.push_back(cellNode);

            cellPos = cellEnd;
        }

        // 🚀 批量处理单元格
        for (const auto& cellNode : cellBatch) {
            if (!callback.onCellNode(cellNode)) return nodeCount;
            nodeCount++;
        }

        pos = rowEnd + 6;
    }

    return nodeCount;
}

const char* TXSIMDXmlParser::findTagStandard(const char* data, size_t length, const char* tag, size_t tagLen) {
    if (length < tagLen) return nullptr;
    
    for (size_t i = 0; i <= length - tagLen; ++i) {
        if (std::memcmp(data + i, tag, tagLen) == 0) {
            return data + i;
        }
    }
    return nullptr;
}

#ifdef TINAXLSX_HAS_SSE2
size_t TXSIMDXmlParser::parseWorksheetSSE2(std::string_view xmlContent, ISIMDXmlCallback& callback) {
    // 🚀 SSE2优化版本：使用SIMD加速标签查找
    size_t nodeCount = 0;
    const char* data = xmlContent.data();
    size_t length = xmlContent.length();

    // 🚀 SIMD优化：快速查找 <sheetData
    const char* sheetDataStart = findTagSSE2(data, length, "<sheetData", 10);
    if (!sheetDataStart) return parseWorksheetStandard(xmlContent, callback);

    // 查找 <sheetData> 结束
    const char* sheetDataEnd = findTagSSE2(sheetDataStart, length - (sheetDataStart - data), "</sheetData>", 12);
    if (!sheetDataEnd) sheetDataEnd = data + length;

    // 🚀 SIMD优化的批量解析
    std::vector<FastXmlNode> cellBatch;
    cellBatch.reserve(2000); // 更大的批量大小

    const char* pos = sheetDataStart;
    while (pos < sheetDataEnd) {
        // 🚀 SIMD查找 <row
        pos = findTagSSE2(pos, sheetDataEnd - pos, "<row", 4);
        if (!pos) break;

        // 查找行结束
        const char* rowEnd = findTagSSE2(pos, sheetDataEnd - pos, "</row>", 6);
        if (!rowEnd) break;

        // 解析行节点
        FastXmlNode rowNode = parseNode(pos, rowEnd + 6);
        if (!callback.onRowNode(rowNode)) break;

        // 🚀 SIMD批量解析单元格
        cellBatch.clear();
        const char* cellPos = pos;

        // 使用SIMD快速扫描所有单元格
        while (cellPos < rowEnd) {
            cellPos = findTagSSE2(cellPos, rowEnd - cellPos, "<c ", 3);
            if (!cellPos) break;

            const char* cellEnd = findTagSSE2(cellPos, rowEnd - cellPos, "</c>", 4);
            if (!cellEnd) {
                cellEnd = findChar(cellPos, rowEnd - cellPos, '>');
                if (!cellEnd) break;
                cellEnd++;
            } else {
                cellEnd += 4;
            }

            FastXmlNode cellNode = parseNode(cellPos, cellEnd);
            cellBatch.push_back(cellNode);
            cellPos = cellEnd;
        }

        // 批量处理单元格
        for (const auto& cellNode : cellBatch) {
            if (!callback.onCellNode(cellNode)) return nodeCount;
            nodeCount++;
        }

        pos = rowEnd + 6;
    }

    return nodeCount;
}

const char* TXSIMDXmlParser::findTagSSE2(const char* data, size_t length, const char* tag, size_t tagLen) {
    // 🚀 SSE2优化的标签查找
    if (length < tagLen || tagLen == 0) return nullptr;
    
    const char firstChar = tag[0];
    const __m128i first = _mm_set1_epi8(firstChar);
    
    size_t simdLength = length & ~15; // 16字节对齐
    
    for (size_t i = 0; i < simdLength; i += 16) {
        __m128i chunk = _mm_loadu_si128(reinterpret_cast<const __m128i*>(data + i));
        __m128i cmp = _mm_cmpeq_epi8(chunk, first);
        
        int mask = _mm_movemask_epi8(cmp);
        while (mask) {
            int pos = _tzcnt_u32(mask); // MSVC版本的计算尾随零
            size_t actualPos = i + pos;
            
            if (actualPos + tagLen <= length && 
                std::memcmp(data + actualPos, tag, tagLen) == 0) {
                return data + actualPos;
            }
            
            mask &= mask - 1; // 清除最低位的1
        }
    }
    
    // 处理剩余字符
    for (size_t i = simdLength; i <= length - tagLen; ++i) {
        if (std::memcmp(data + i, tag, tagLen) == 0) {
            return data + i;
        }
    }
    
    return nullptr;
}
#endif

#ifdef TINAXLSX_HAS_AVX2
size_t TXSIMDXmlParser::parseWorksheetAVX2(std::string_view xmlContent, ISIMDXmlCallback& callback) {
    // 🚀 AVX2优化版本：使用更宽的SIMD指令
    // 目前回退到SSE2实现
    return parseWorksheetSSE2(xmlContent, callback);
}

const char* TXSIMDXmlParser::findTagAVX2(const char* data, size_t length, const char* tag, size_t tagLen) {
    // 🚀 AVX2优化的标签查找
    if (length < tagLen || tagLen == 0) return nullptr;
    
    const char firstChar = tag[0];
    const __m256i first = _mm256_set1_epi8(firstChar);
    
    size_t simdLength = length & ~31; // 32字节对齐
    
    for (size_t i = 0; i < simdLength; i += 32) {
        __m256i chunk = _mm256_loadu_si256(reinterpret_cast<const __m256i*>(data + i));
        __m256i cmp = _mm256_cmpeq_epi8(chunk, first);
        
        int mask = _mm256_movemask_epi8(cmp);
        while (mask) {
            int pos = _tzcnt_u32(mask); // MSVC版本的计算尾随零
            size_t actualPos = i + pos;
            
            if (actualPos + tagLen <= length && 
                std::memcmp(data + actualPos, tag, tagLen) == 0) {
                return data + actualPos;
            }
            
            mask &= mask - 1;
        }
    }
    
    // 回退到SSE2处理剩余部分
    return findTagSSE2(data + simdLength, length - simdLength, tag, tagLen);
}
#endif

// ==================== 辅助方法 ====================

FastXmlNode TXSIMDXmlParser::parseNode(const char* start, const char* end) {
    FastXmlNode node;
    node.start = start;
    node.end = end;
    
    // 查找标签名
    const char* nameStart = start + 1; // 跳过 <
    const char* nameEnd = findAnyChar(nameStart, end - nameStart, " \t\n\r>", 5);
    if (!nameEnd) nameEnd = end;
    
    node.name = std::string_view(nameStart, nameEnd - nameStart);
    
    // 提取属性
    if (*nameEnd != '>') {
        const char* attrEnd = findChar(nameEnd, end - nameEnd, '>');
        if (attrEnd) {
            node.attributes = std::string_view(nameEnd, attrEnd - nameEnd);
        }
    }
    
    // 提取值
    node.value = extractNodeValue(start, end);
    
    return node;
}

std::string_view TXSIMDXmlParser::extractNodeValue(const char* start, const char* end) {
    // 查找 > 和 <
    const char* valueStart = findChar(start, end - start, '>');
    if (!valueStart) return {};
    valueStart++; // 跳过 >
    
    const char* valueEnd = findChar(valueStart, end - valueStart, '<');
    if (!valueEnd) valueEnd = end;
    
    return std::string_view(valueStart, valueEnd - valueStart);
}

std::string_view TXSIMDXmlParser::extractAttributes(const char* start, const char* end) {
    const char* attrStart = findAnyChar(start, end - start, " \t", 2);
    if (!attrStart) return {};
    
    const char* attrEnd = findChar(attrStart, end - attrStart, '>');
    if (!attrEnd) attrEnd = end;
    
    return std::string_view(attrStart, attrEnd - attrStart);
}

const char* TXSIMDXmlParser::findChar(const char* data, size_t length, char target) {
    const char* result = static_cast<const char*>(std::memchr(data, target, length));
    return result;
}

const char* TXSIMDXmlParser::findAnyChar(const char* data, size_t length, const char* targets, size_t targetCount) {
    for (size_t i = 0; i < length; ++i) {
        for (size_t j = 0; j < targetCount; ++j) {
            if (data[i] == targets[j]) {
                return data + i;
            }
        }
    }
    return nullptr;
}

// ==================== TXSIMDWorksheetParser 实现 ====================

TXSIMDWorksheetParser::TXSIMDWorksheetParser(TXSheet* sheet)
    : sheet_(sheet) {
    cellBatch_.reserve(BATCH_SIZE);
}

size_t TXSIMDWorksheetParser::parse(std::string_view xmlContent) {
    auto startTime = std::chrono::high_resolution_clock::now();

    // 重置统计信息
    stats_ = WorksheetStats{};

    // 使用SIMD解析器
    size_t nodeCount = parser_.parseWorksheet(xmlContent, *this);

    // 刷新剩余批量数据
    flushBatch();

    auto endTime = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(endTime - startTime);
    stats_.parseTimeMs = duration.count() / 1000.0;

    return stats_.totalCells;
}

bool TXSIMDWorksheetParser::onCellNode(const FastXmlNode& node) {
    stats_.totalCells++;

    // 🚀 快速解析单元格属性
    auto ref = node.getAttribute("r");
    if (ref.empty()) {
        stats_.emptySkipped++;
        return true;
    }

    // 解析坐标
    TXCoordinate coord = parseCoordinate(ref);

    // 解析值
    std::string value = parseValue(node);
    if (value.empty()) {
        stats_.emptySkipped++;
        return true;
    }

    // 解析样式
    u32 styleIndex = parseStyleIndex(node);

    // 🚀 添加到批量缓存
    addCellToBatch(coord, value, styleIndex);

    return true;
}

bool TXSIMDWorksheetParser::onRowNode(const FastXmlNode& node) {
    stats_.totalRows++;
    return true;
}

void TXSIMDWorksheetParser::onSharedString(u32 index, std::string_view text) {
    // TODO: 处理共享字符串
}

void TXSIMDWorksheetParser::flushBatch() {
    if (cellBatch_.size() == 0) return;

    // 🚀 超高性能批量设置：使用CellManager的批量接口
    auto& cellManager = sheet_->getCellManager();

    // 转换为CellManager期望的格式
    std::vector<std::pair<TXCellManager::Coordinate, TXCellManager::CellValue>> cellValues;
    cellValues.reserve(cellBatch_.size());

    for (size_t i = 0; i < cellBatch_.size(); ++i) {
        cellValues.emplace_back(cellBatch_.coordinates[i], cellBatch_.values[i]);
    }

    // 🚀 使用真正的批量设置接口
    cellManager.setCellValues(cellValues);

    cellBatch_.clear();
}

void TXSIMDWorksheetParser::addCellToBatch(const TXCoordinate& coord, const std::string& value, u32 styleIndex) {
    cellBatch_.coordinates.push_back(coord);
    cellBatch_.values.push_back(value);
    cellBatch_.styleIndices.push_back(styleIndex);

    if (cellBatch_.size() >= BATCH_SIZE) {
        flushBatch();
    }
}

TXCoordinate TXSIMDWorksheetParser::parseCoordinate(std::string_view ref) {
    // 🚀 超高性能坐标解析：直接解析，避免字符串创建
    if (ref.empty()) return TXCoordinate(row_t(row_t::index_t(1)), column_t(column_t::index_t(1)));

    // 解析列（字母部分）
    column_t::index_t colValue = 0;
    size_t i = 0;
    while (i < ref.length() && std::isalpha(ref[i])) {
        colValue = colValue * 26 + (std::toupper(ref[i]) - 'A' + 1);
        i++;
    }

    // 解析行（数字部分）
    row_t::index_t rowValue = 0;
    while (i < ref.length() && std::isdigit(ref[i])) {
        rowValue = rowValue * 10 + (ref[i] - '0');
        i++;
    }

    return TXCoordinate(row_t(rowValue), column_t(colValue));
}

std::string TXSIMDWorksheetParser::parseValue(const FastXmlNode& cellNode) {
    // 🚀 优化：使用更快的字符串查找
    size_t nodeLength = cellNode.end - cellNode.start;

    // 查找 <v> 节点
    const char* vStart = std::strstr(cellNode.start, "<v>");
    if (vStart && vStart < cellNode.end) {
        vStart += 3; // 跳过 <v>
        const char* vEnd = std::strstr(vStart, "</v>");
        if (vEnd && vEnd < cellNode.end) {
            // 使用新的字符串池管理字符串
            std::string valueStr(vStart, vEnd - vStart);
            TXStringPool::getInstance().intern(valueStr);
            return valueStr;
        }
    }

    // 查找内联字符串 <is><t>
    const char* isStart = std::strstr(cellNode.start, "<is>");
    if (isStart && isStart < cellNode.end) {
        const char* tStart = std::strstr(isStart, "<t>");
        if (tStart && tStart < cellNode.end) {
            tStart += 3;
            const char* tEnd = std::strstr(tStart, "</t>");
            if (tEnd && tEnd < cellNode.end) {
                // 使用新的字符串池管理字符串
                std::string valueStr(tStart, tEnd - tStart);
                TXStringPool::getInstance().intern(valueStr);
                return valueStr;
            }
        }
    }

    return {};
}

u32 TXSIMDWorksheetParser::parseStyleIndex(const FastXmlNode& cellNode) {
    auto sAttr = cellNode.getAttribute("s");
    if (sAttr.empty()) return 0;

    // 🚀 快速数字解析
    u32 result = 0;
    for (char c : sAttr) {
        if (c >= '0' && c <= '9') {
            result = result * 10 + (c - '0');
        } else {
            break;
        }
    }
    return result;
}

} // namespace TinaXlsx
