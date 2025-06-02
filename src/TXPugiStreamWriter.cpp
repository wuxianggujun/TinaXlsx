//
// @file TXPugiStreamWriter.cpp
// @brief 基于pugixml xml_writer的高性能流式XML写入器实现
//

#include "TinaXlsx/TXPugiStreamWriter.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include <sstream>
#include <iomanip>

namespace TinaXlsx {

// TXBufferedXmlWriter 实现
TXBufferedXmlWriter::TXBufferedXmlWriter(size_t bufferSize) 
    : totalBytes_(0) {
    buffer_.reserve(bufferSize);
}

void TXBufferedXmlWriter::write(const void* data, size_t size) {
    const uint8_t* bytes = static_cast<const uint8_t*>(data);
    buffer_.insert(buffer_.end(), bytes, bytes + size);
    totalBytes_ += size;
}

void TXBufferedXmlWriter::clear() {
    buffer_.clear();
    totalBytes_ = 0;
}

// TXPugiWorksheetWriter 实现
TXPugiWorksheetWriter::TXPugiWorksheetWriter(size_t bufferSize)
    : writer_(std::make_unique<TXBufferedXmlWriter>(bufferSize))
    , doc_(std::make_unique<pugi::xml_document>())
    , worksheetStarted_(false)
    , sheetDataStarted_(false)
    , rowStarted_(false) {
    stats_ = {};
}

void TXPugiWorksheetWriter::startWorksheet(const std::string& usedRangeRef, bool hasCustomColumns) {
    if (worksheetStarted_) {
        return; // 已经开始
    }
    
    // 创建根节点
    auto worksheet = doc_->append_child("worksheet");
    worksheet.append_attribute("xmlns").set_value("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    worksheet.append_attribute("xmlns:r").set_value("http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    // 添加维度信息
    auto dimension = worksheet.append_child("dimension");
    dimension.append_attribute("ref").set_value(usedRangeRef.c_str());
    
    // 如果有自定义列宽，预留cols节点
    if (hasCustomColumns) {
        worksheet.append_child("cols");
    }
    
    worksheetStarted_ = true;
}

void TXPugiWorksheetWriter::writeColumnWidth(u32 columnIndex, double width) {
    if (!worksheetStarted_) {
        return;
    }
    
    auto worksheet = doc_->document_element();
    auto cols = worksheet.child("cols");
    if (!cols) {
        cols = worksheet.append_child("cols");
    }
    
    auto col = cols.append_child("col");
    col.append_attribute("min").set_value(columnIndex);
    col.append_attribute("max").set_value(columnIndex);
    col.append_attribute("width").set_value(width);
    col.append_attribute("customWidth").set_value("1");
}

void TXPugiWorksheetWriter::startSheetData() {
    if (!worksheetStarted_ || sheetDataStarted_) {
        return;
    }
    
    auto worksheet = doc_->document_element();
    worksheet.append_child("sheetData");
    sheetDataStarted_ = true;
}

void TXPugiWorksheetWriter::startRow(u32 rowNumber) {
    if (!sheetDataStarted_ || rowStarted_) {
        return;
    }
    
    auto worksheet = doc_->document_element();
    auto sheetData = worksheet.child("sheetData");
    auto row = sheetData.append_child("row");
    row.append_attribute("r").set_value(rowNumber);
    
    rowStarted_ = true;
    stats_.rowsWritten++;
}

void TXPugiWorksheetWriter::writeCellInlineString(const std::string& cellRef, const std::string& value, u32 styleIndex) {
    if (!rowStarted_) {
        return;
    }
    
    auto worksheet = doc_->document_element();
    auto sheetData = worksheet.child("sheetData");
    auto currentRow = sheetData.last_child(); // 当前行
    
    auto cell = currentRow.append_child("c");
    cell.append_attribute("r").set_value(cellRef.c_str());
    cell.append_attribute("t").set_value("inlineStr");
    
    if (styleIndex > 0) {
        cell.append_attribute("s").set_value(styleIndex);
    }
    
    auto is = cell.append_child("is");
    auto t = is.append_child("t");
    t.text().set(escapeXmlText(value).c_str());
    
    stats_.cellsWritten++;
}

void TXPugiWorksheetWriter::writeCellSharedString(const std::string& cellRef, u32 stringIndex, u32 styleIndex) {
    if (!rowStarted_) {
        return;
    }
    
    auto worksheet = doc_->document_element();
    auto sheetData = worksheet.child("sheetData");
    auto currentRow = sheetData.last_child(); // 当前行
    
    auto cell = currentRow.append_child("c");
    cell.append_attribute("r").set_value(cellRef.c_str());
    cell.append_attribute("t").set_value("s");
    
    if (styleIndex > 0) {
        cell.append_attribute("s").set_value(styleIndex);
    }
    
    auto v = cell.append_child("v");
    v.text().set(std::to_string(stringIndex).c_str());
    
    stats_.cellsWritten++;
}

void TXPugiWorksheetWriter::writeCellNumber(const std::string& cellRef, double value, u32 styleIndex) {
    if (!rowStarted_) {
        return;
    }
    
    auto worksheet = doc_->document_element();
    auto sheetData = worksheet.child("sheetData");
    auto currentRow = sheetData.last_child(); // 当前行
    
    auto cell = currentRow.append_child("c");
    cell.append_attribute("r").set_value(cellRef.c_str());
    
    if (styleIndex > 0) {
        cell.append_attribute("s").set_value(styleIndex);
    }
    
    auto v = cell.append_child("v");
    
    // 使用高精度数值转换
    std::ostringstream oss;
    oss << std::setprecision(15) << value;
    v.text().set(oss.str().c_str());
    
    stats_.cellsWritten++;
}

void TXPugiWorksheetWriter::writeCellInteger(const std::string& cellRef, int64_t value, u32 styleIndex) {
    if (!rowStarted_) {
        return;
    }
    
    auto worksheet = doc_->document_element();
    auto sheetData = worksheet.child("sheetData");
    auto currentRow = sheetData.last_child(); // 当前行
    
    auto cell = currentRow.append_child("c");
    cell.append_attribute("r").set_value(cellRef.c_str());
    
    if (styleIndex > 0) {
        cell.append_attribute("s").set_value(styleIndex);
    }
    
    auto v = cell.append_child("v");
    v.text().set(std::to_string(value).c_str());
    
    stats_.cellsWritten++;
}

void TXPugiWorksheetWriter::writeCellBoolean(const std::string& cellRef, bool value, u32 styleIndex) {
    if (!rowStarted_) {
        return;
    }
    
    auto worksheet = doc_->document_element();
    auto sheetData = worksheet.child("sheetData");
    auto currentRow = sheetData.last_child(); // 当前行
    
    auto cell = currentRow.append_child("c");
    cell.append_attribute("r").set_value(cellRef.c_str());
    cell.append_attribute("t").set_value("b");
    
    if (styleIndex > 0) {
        cell.append_attribute("s").set_value(styleIndex);
    }
    
    auto v = cell.append_child("v");
    v.text().set(value ? "1" : "0");
    
    stats_.cellsWritten++;
}

void TXPugiWorksheetWriter::endRow() {
    if (!rowStarted_) {
        return;
    }
    
    rowStarted_ = false;
}

void TXPugiWorksheetWriter::endSheetData() {
    if (!sheetDataStarted_) {
        return;
    }
    
    sheetDataStarted_ = false;
}

void TXPugiWorksheetWriter::endWorksheet() {
    if (!worksheetStarted_) {
        return;
    }
    
    // 使用pugixml的xml_writer接口生成XML
    writer_->clear();
    
    // 写入XML声明
    const char* xmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
    writer_->write(xmlDecl, strlen(xmlDecl));
    
    // 使用pugixml的save方法写入到我们的writer
    doc_->save(*writer_, "", pugi::format_raw | pugi::format_no_declaration, pugi::encoding_utf8);
    
    stats_.totalBytesWritten = writer_->getTotalBytesWritten();
    worksheetStarted_ = false;
}

TXResult<void> TXPugiWorksheetWriter::writeToZip(TXZipArchiveWriter& zipWriter, const std::string& partName) {
    if (worksheetStarted_) {
        endWorksheet(); // 确保工作表已结束
    }
    
    auto writeResult = zipWriter.write(partName, writer_->getBuffer());
    if (writeResult.isError()) {
        return Err<void>(writeResult.error().getCode(), 
                       "Failed to write " + partName + ": " + writeResult.error().getMessage());
    }
    
    return Ok();
}

const std::vector<uint8_t>& TXPugiWorksheetWriter::getXmlData() const {
    return writer_->getBuffer();
}

void TXPugiWorksheetWriter::reset() {
    writer_->clear();
    doc_->reset();
    stats_ = {};
    worksheetStarted_ = false;
    sheetDataStarted_ = false;
    rowStarted_ = false;
}

std::string TXPugiWorksheetWriter::escapeXmlText(const std::string& text) {
    std::string result;
    result.reserve(text.length() * 1.1); // 预留一些空间给转义字符
    
    for (char c : text) {
        switch (c) {
            case '<':
                result += "&lt;";
                break;
            case '>':
                result += "&gt;";
                break;
            case '&':
                result += "&amp;";
                break;
            case '"':
                result += "&quot;";
                break;
            case '\'':
                result += "&apos;";
                break;
            default:
                result += c;
                break;
        }
    }
    
    return result;
}

// TXWorksheetWriterFactory 实现
std::unique_ptr<TXPugiWorksheetWriter> TXWorksheetWriterFactory::createWriter(size_t estimatedCells) {
    // 根据预估单元格数量调整缓冲区大小
    size_t bufferSize = 64 * 1024; // 默认64KB
    
    if (estimatedCells > 50000) {
        bufferSize = 1024 * 1024; // 1MB
    } else if (estimatedCells > 10000) {
        bufferSize = 512 * 1024; // 512KB
    } else if (estimatedCells > 5000) {
        bufferSize = 256 * 1024; // 256KB
    }
    
    return std::make_unique<TXPugiWorksheetWriter>(bufferSize);
}

bool TXWorksheetWriterFactory::shouldUseStreamWriter(size_t estimatedCells) {
    return estimatedCells >= STREAM_WRITER_THRESHOLD;
}

} // namespace TinaXlsx
