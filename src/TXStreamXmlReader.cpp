//
// @file TXStreamXmlReader.cpp
// @brief 流式XML读取器实现
//

#include "TinaXlsx/TXStreamXmlReader.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCoordinate.hpp"
#include "TinaXlsx/TXResult.hpp"
#include "TinaXlsx/TXTypes.hpp"
#include "TinaXlsx/TXCellManager.hpp"
#include <pugixml.hpp>
#include <chrono>
#include <sstream>

namespace TinaXlsx {

// ==================== CellData 实现 ====================

void CellData::parseCoordinates() const {
    if (coordinatesParsed) return;
    
    // 解析单元格引用 (如 "A1" -> row=1, col=1)
    if (!ref.empty()) {
        auto coord = TXCoordinate::fromAddress(ref);
        row = coord.getRow();
        col = coord.getCol();
    }
    coordinatesParsed = true;
}

// ==================== TXStreamXmlReader 实现 ====================

TXStreamXmlReader::TXStreamXmlReader() 
    : doc_(std::make_unique<pugi::xml_document>()) {
}

TXStreamXmlReader::~TXStreamXmlReader() = default;

TXResult<void> TXStreamXmlReader::parseWorksheet(TXZipArchiveReader& zipReader, 
                                                const std::string& worksheetPath,
                                                IStreamXmlCallback& callback) {
    // 读取XML数据
    auto xmlData = zipReader.read(worksheetPath);
    if (xmlData.isError()) {
        return Err<void>(xmlData.error().getCode(), "Failed to read " + worksheetPath);
    }
    
    const std::vector<uint8_t>& fileBytes = xmlData.value();
    std::string xmlContent(fileBytes.begin(), fileBytes.end());
    
    return parseWorksheetImpl(xmlContent, callback);
}

TXResult<void> TXStreamXmlReader::parseWorksheetImpl(const std::string& xmlContent, 
                                                    IStreamXmlCallback& callback) {
    // 🚀 性能优化：使用pugi的快速解析模式
    auto parseResult = doc_->load_string(xmlContent.c_str(), 
                                        pugi::parse_default | pugi::parse_trim_pcdata);
    
    if (!parseResult) {
        return Err<void>(TXErrorCode::XmlParseError, 
                        "XML parse error: " + std::string(parseResult.description()));
    }
    
    // 🚀 性能优化：直接遍历sheetData节点，避免XPath查询
    auto sheetData = doc_->child("worksheet").child("sheetData");
    if (!sheetData) {
        return Ok(); // 空工作表
    }
    
    std::vector<RowData> rowBatch;
    rowBatch.reserve(options_.batchSize);
    
    // 🚀 性能优化：直接遍历行节点，避免递归查找
    for (auto rowNode = sheetData.child("row"); rowNode; rowNode = rowNode.next_sibling("row")) {
        if (!parseRowNode(rowNode, callback)) {
            break; // 回调要求停止
        }
    }
    
    return Ok();
}

bool TXStreamXmlReader::parseRowNode(const pugi::xml_node& rowNode, IStreamXmlCallback& callback) {
    RowData rowData;
    
    // 解析行属性
    if (auto rAttr = rowNode.attribute("r")) {
        rowData.rowIndex = rAttr.as_uint();
    }
    
    if (auto htAttr = rowNode.attribute("ht")) {
        rowData.height = htAttr.as_double();
        rowData.customHeight = true;
    }
    
    if (auto hiddenAttr = rowNode.attribute("hidden")) {
        rowData.hidden = hiddenAttr.as_bool();
    }
    
    // 🚀 性能优化：预分配单元格容器
    rowData.cells.reserve(32); // 假设平均每行32个单元格
    
    // 解析单元格
    for (auto cellNode = rowNode.child("c"); cellNode; cellNode = cellNode.next_sibling("c")) {
        CellData cellData = parseCellNode(cellNode);
        
        // 跳过空单元格（如果设置了选项）
        if (options_.skipEmptyCells && cellData.value.empty() && cellData.styleIndex == 0) {
            continue;
        }
        
        rowData.cells.push_back(std::move(cellData));
    }
    
    // 调用回调处理行数据
    return callback.onRowData(rowData);
}

CellData TXStreamXmlReader::parseCellNode(const pugi::xml_node& cellNode) {
    CellData cellData;
    
    // 解析单元格引用
    if (auto rAttr = cellNode.attribute("r")) {
        cellData.ref = rAttr.value();
    }
    
    // 解析单元格类型
    if (auto tAttr = cellNode.attribute("t")) {
        cellData.type = tAttr.value();
    }
    
    // 解析样式索引
    if (auto sAttr = cellNode.attribute("s")) {
        cellData.styleIndex = sAttr.as_uint();
    }
    
    // 解析单元格值
    auto valueNode = cellNode.child("v");
    if (valueNode) {
        cellData.value = valueNode.text().as_string();
    } else {
        // 检查内联字符串
        auto isNode = cellNode.child("is");
        if (isNode) {
            auto tNode = isNode.child("t");
            if (tNode) {
                cellData.value = tNode.text().as_string();
                cellData.type = "inlineStr";
            }
        }
    }
    
    return cellData;
}

TXResult<void> TXStreamXmlReader::parseSharedStrings(TXZipArchiveReader& zipReader,
                                                    IStreamXmlCallback& callback) {
    // 读取共享字符串XML
    auto xmlData = zipReader.read("xl/sharedStrings.xml");
    if (xmlData.isError()) {
        return Ok(); // 没有共享字符串文件
    }
    
    const std::vector<uint8_t>& fileBytes = xmlData.value();
    std::string xmlContent(fileBytes.begin(), fileBytes.end());
    
    auto parseResult = doc_->load_string(xmlContent.c_str());
    if (!parseResult) {
        return Err<void>(TXErrorCode::XmlParseError, 
                        "Failed to parse sharedStrings.xml");
    }
    
    auto sst = doc_->child("sst");
    if (!sst) {
        return Ok();
    }
    
    u32 index = 0;
    for (auto si = sst.child("si"); si; si = si.next_sibling("si")) {
        std::string text;
        
        // 处理简单文本
        auto t = si.child("t");
        if (t) {
            text = t.text().as_string();
        } else {
            // 处理富文本（简化处理）
            for (auto r = si.child("r"); r; r = r.next_sibling("r")) {
                auto rt = r.child("t");
                if (rt) {
                    text += rt.text().as_string();
                }
            }
        }
        
        callback.onSharedString(index++, text);
    }
    
    return Ok();
}

TXResult<void> TXStreamXmlReader::parseStyles(TXZipArchiveReader& zipReader,
                                             IStreamXmlCallback& callback) {
    // 简化的样式解析实现
    auto xmlData = zipReader.read("xl/styles.xml");
    if (xmlData.isError()) {
        return Ok(); // 没有样式文件
    }
    
    // TODO: 实现详细的样式解析
    // 目前返回成功，后续可以扩展
    return Ok();
}

// ==================== TXFastWorksheetLoader 实现 ====================

TXFastWorksheetLoader::TXFastWorksheetLoader(TXSheet* sheet) 
    : sheet_(sheet) {
    cellBatch_.reserve(BATCH_SIZE);
}

TXResult<void> TXFastWorksheetLoader::load(TXZipArchiveReader& zipReader,
                                          const std::string& worksheetPath) {
    auto startTime = std::chrono::high_resolution_clock::now();

    // 重置统计信息
    stats_ = LoadStats{};

    // 🚀 性能优化：设置流式解析选项
    TXStreamXmlReader::ParseOptions options;
    options.skipEmptyCells = true;
    options.batchSize = 1000;
    reader_.setOptions(options);

    // 使用流式解析器加载
    auto result = reader_.parseWorksheet(zipReader, worksheetPath, *this);

    auto endTime = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(endTime - startTime);
    stats_.loadTimeMs = duration.count() / 1000.0;

    return result;
}

bool TXFastWorksheetLoader::onRowData(const RowData& rowData) {
    stats_.totalRows++;

    // 🚀 性能优化：直接按行批量设置，避免逐个setCellValue
    std::vector<std::pair<TXCoordinate, std::string>> rowCells;
    rowCells.reserve(rowData.cells.size());

    for (const auto& cellData : rowData.cells) {
        stats_.totalCells++;

        if (cellData.value.empty() && cellData.styleIndex == 0) {
            stats_.emptySkipped++;
            continue;
        }

        // 🚀 解析坐标一次，避免重复解析
        cellData.parseCoordinates();
        TXCoordinate coord(cellData.row, cellData.col);
        rowCells.emplace_back(coord, cellData.value);
    }

    // 🚀 批量设置整行数据
    if (!rowCells.empty()) {
        batchSetCells(rowCells);
    }

    return true; // 继续处理
}

void TXFastWorksheetLoader::onSharedString(u32 index, const std::string& text) {
    // TODO: 处理共享字符串
}

void TXFastWorksheetLoader::onStyleData(u32 styleIndex, const std::string& styleData) {
    // TODO: 处理样式数据
}

void TXFastWorksheetLoader::batchSetCells(const std::vector<std::pair<TXCoordinate, std::string>>& cells) {
    // 🚀 性能优化：直接逐个设置，但避免字符串解析
    auto& cellManager = sheet_->getCellManager();

    for (const auto& [coord, value] : cells) {
        // 🚀 直接使用坐标对象，避免字符串解析
        cellManager.setCellValue(coord, value);
    }
}

void TXFastWorksheetLoader::flushBatch() {
    // 这个方法现在不再使用，保留用于兼容性
}

} // namespace TinaXlsx
