#include "TinaXlsx/TXWorksheetWriter.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXCoordinate.hpp"
#include "TinaXlsx/TXRange.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include <variant>

namespace TinaXlsx {

TXWorksheetWriter::TXWorksheetWriter() = default;

TXWorksheetWriter::~TXWorksheetWriter() = default;

bool TXWorksheetWriter::writeToZip(TXZipArchiveWriter& zip, const TXSheet* sheet, std::size_t sheetIndex) {
    if (!sheet) {
        lastError_ = "Sheet is null";
        return false;
    }
    
    std::string xml = generateXml(sheet);
    if (xml.empty()) {
        lastError_ = "Failed to generate XML for sheet";
        return false;
    }
    
    std::string filename = "xl/worksheets/sheet" + std::to_string(sheetIndex) + ".xml";
    std::vector<uint8_t> xml_data(xml.begin(), xml.end());
    if (!zip.write(filename, xml_data)) {
        lastError_ = "Failed to write worksheet to zip: " + zip.lastError();
        return false;
    }
    
    lastError_.clear();
    return true;
}

std::string TXWorksheetWriter::generateXml(const TXSheet* sheet) {
    if (!sheet) {
        lastError_ = "Sheet is null";
        return "";
    }
    
    TXXmlWriter writer;
    
    try {
        writeWorksheetHeader(writer, sheet);
        writeDimension(writer, sheet);
        writeSheetData(writer, sheet);
        writeMergeCells(writer, sheet);
        
        // 结束worksheet元素
        writer.endElement();
        
        lastError_.clear();
        return writer.toString();
    } catch (const std::exception& e) {
        lastError_ = "Exception during XML generation: " + std::string(e.what());
        return "";
    }
}

const std::string& TXWorksheetWriter::getLastError() const {
    return lastError_;
}

void TXWorksheetWriter::writeWorksheetHeader(TXXmlWriter& writer, const TXSheet* sheet) {
    writer.startDocument();
    
    // 开始worksheet元素并添加命名空间
    std::vector<std::pair<std::string, std::string>> attributes = {
        {"xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"},
        {"xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    };
    writer.startElement("worksheet", attributes);
}

void TXWorksheetWriter::writeDimension(TXXmlWriter& writer, const TXSheet* sheet) {
    auto usedRange = sheet->getUsedRange();
    std::string dimensionRef = usedRange.isValid() ? usedRange.toAddress() : "A1";
    
    std::vector<std::pair<std::string, std::string>> attributes = {
        {"ref", dimensionRef}
    };
    writer.startElement("dimension", attributes);
    writer.endElement(); // dimension
}

void TXWorksheetWriter::writeSheetData(TXXmlWriter& writer, const TXSheet* sheet) {
    writer.startElement("sheetData");
    
    auto usedRange = sheet->getUsedRange();
    if (usedRange.isValid()) {
        for (row_t row = usedRange.getStart().getRow(); row <= usedRange.getEnd().getRow(); ++row) {
            if (writeRow(writer, sheet, row, usedRange)) {
                // 行写入成功
            }
        }
    }
    
    writer.endElement(); // sheetData
}

bool TXWorksheetWriter::writeRow(TXXmlWriter& writer, const TXSheet* sheet, row_t row, const TXRange& usedRange) {
    bool rowHasData = false;
    
    // 先检查这一行是否有数据
    for (column_t col = usedRange.getStart().getCol(); col <= usedRange.getEnd().getCol(); ++col) {
        const TXCell* cell = sheet->getCell(row, col);
        if (cell && !cell->isEmpty()) {
            rowHasData = true;
            break;
        }
    }
    
    if (!rowHasData) {
        return false; // 这一行没有数据，不写入
    }
    
    // 开始行元素
    std::vector<std::pair<std::string, std::string>> rowAttributes = {
        {"r", std::to_string(row.index())},
        {"spans", "1:1024"}
    };
    writer.startElement("row", rowAttributes);
    
    // 写入这一行的所有单元格
    for (column_t col = usedRange.getStart().getCol(); col <= usedRange.getEnd().getCol(); ++col) {
        const TXCell* cell = sheet->getCell(row, col);
        if (cell && !cell->isEmpty()) {
            TXCoordinate coord(row, col);
            std::string cellRef = coord.toAddress();
            writeCell(writer, cell, cellRef);
        }
    }
    
    writer.endElement(); // row
    return true;
}

void TXWorksheetWriter::writeCell(TXXmlWriter& writer, const TXCell* cell, const std::string& cellRef) {
    std::vector<std::pair<std::string, std::string>> cellAttributes = {
        {"r", cellRef}
    };
    
    const auto& value = cell->getValue();
    
    // 根据数据类型设置不同的类型和值
    if (std::holds_alternative<std::string>(value)) {
        // 字符串类型 - 使用内联字符串
        cellAttributes.emplace_back("t", "inlineStr");
        
        // 添加样式属性（如果有）
        if (cell->hasStyle()) {
            cellAttributes.emplace_back("s", std::to_string(cell->getStyleIndex()));
        }
        
        writer.startElement("c", cellAttributes);
        
        // 写入内联字符串
        writer.startElement("is");
        writer.startElement("t");
        writer.addText(std::get<std::string>(value), true);
        writer.endElement(); // t
        writer.endElement(); // is
        
    } else if (std::holds_alternative<double>(value)) {
        // 数值类型
        cellAttributes.emplace_back("t", "n");
        
        if (cell->hasStyle()) {
            cellAttributes.emplace_back("s", std::to_string(cell->getStyleIndex()));
        }
        
        writer.startElement("c", cellAttributes);
        writer.startElement("v");
        writer.addText(std::to_string(std::get<double>(value)), false);
        writer.endElement(); // v
        
    } else if (std::holds_alternative<int64_t>(value)) {
        // 整数类型
        cellAttributes.emplace_back("t", "n");
        
        if (cell->hasStyle()) {
            cellAttributes.emplace_back("s", std::to_string(cell->getStyleIndex()));
        }
        
        writer.startElement("c", cellAttributes);
        writer.startElement("v");
        writer.addText(std::to_string(std::get<int64_t>(value)), false);
        writer.endElement(); // v
        
    } else if (std::holds_alternative<bool>(value)) {
        // 布尔类型
        cellAttributes.emplace_back("t", "b");
        
        if (cell->hasStyle()) {
            cellAttributes.emplace_back("s", std::to_string(cell->getStyleIndex()));
        }
        
        writer.startElement("c", cellAttributes);
        writer.startElement("v");
        writer.addText(std::get<bool>(value) ? "1" : "0", false);
        writer.endElement(); // v
        
    } else {
        // 其他类型或空值，跳过
        return;
    }
    
    writer.endElement(); // c
}

void TXWorksheetWriter::writeMergeCells(TXXmlWriter& writer, const TXSheet* sheet) {
    auto mergeRegions = sheet->getAllMergeRegions();
    if (mergeRegions.empty()) {
        return; // 没有合并单元格
    }
    
    std::vector<std::pair<std::string, std::string>> attributes = {
        {"count", std::to_string(mergeRegions.size())}
    };
    writer.startElement("mergeCells", attributes);
    
    for (const auto& region : mergeRegions) {
        std::vector<std::pair<std::string, std::string>> mergeCellAttributes = {
            {"ref", region.toAddress()}
        };
        writer.startElement("mergeCell", mergeCellAttributes);
        writer.endElement(); // mergeCell
    }
    
    writer.endElement(); // mergeCells
}

} // namespace TinaXlsx 