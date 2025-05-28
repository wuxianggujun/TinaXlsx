//
// @file TXWorksheetWriter.cpp - 新架构实现
//

#include "TinaXlsx/TXWorksheetWriter.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXRange.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include <sstream>

namespace TinaXlsx {

TXWorksheetWriter::TXWorksheetWriter() 
    : lastError_(""), xmlWriter_(std::make_unique<TXXmlWriter>()) {}

TXWorksheetWriter::~TXWorksheetWriter() = default;

bool TXWorksheetWriter::writeWorksheetToFile(const std::string& xlsxFilePath, const TXSheet* sheet, 
                                           std::size_t sheetIndex, bool appendMode) {
    if (!sheet) {
        lastError_ = "Sheet is null";
        return false;
    }
    
    // 打开 ZIP 文件用于写入
    TXZipArchiveWriter zipWriter;
    if (!zipWriter.open(xlsxFilePath, appendMode)) {
        lastError_ = "Failed to open XLSX file for writing: " + zipWriter.lastError();
        return false;
    }
    
    // 构建工作表 XML
    XmlNodeBuilder worksheetXml = buildWorksheetXml(sheet);
    
    // 设置到 XML 写入器
    xmlWriter_->setRootNode(worksheetXml);
    
    // 获取工作表的 XML 路径
    std::string xmlPath = getWorksheetXmlPath(sheetIndex);
    
    // 写入到 ZIP 文件
    if (!xmlWriter_->writeToZip(zipWriter, xmlPath)) {
        lastError_ = "Failed to write worksheet XML: " + xmlWriter_->getLastError();
        return false;
    }
    
    return true;
}

bool TXWorksheetWriter::writeWorksheetToZip(TXZipArchiveWriter& zipWriter, const TXSheet* sheet, std::size_t sheetIndex) {
    if (!sheet) {
        lastError_ = "Sheet is null";
        return false;
    }
    
    // 构建工作表 XML
    XmlNodeBuilder worksheetXml = buildWorksheetXml(sheet);
    
    // 设置到 XML 写入器
    xmlWriter_->setRootNode(worksheetXml);
    
    // 获取工作表的 XML 路径
    std::string xmlPath = getWorksheetXmlPath(sheetIndex);
    
    // 写入到 ZIP 文件
    if (!xmlWriter_->writeToZip(zipWriter, xmlPath)) {
        lastError_ = "Failed to write worksheet XML: " + xmlWriter_->getLastError();
        return false;
    }
    
    return true;
}

std::string TXWorksheetWriter::generateXml(const TXSheet* sheet) {
    if (!sheet) {
        lastError_ = "Sheet is null";
        return "";
    }
    
    // 构建工作表 XML
    XmlNodeBuilder worksheetXml = buildWorksheetXml(sheet);
    
    // 设置到 XML 写入器
    xmlWriter_->setRootNode(worksheetXml);
    
    // 生成 XML 字符串
    return xmlWriter_->generateXmlString();
}

XmlNodeBuilder TXWorksheetWriter::buildWorksheetXml(const TXSheet* sheet) {
    // 创建根节点
    XmlNodeBuilder worksheet("worksheet");
    worksheet.addAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
             .addAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    
    // 添加维度信息
    worksheet.addChild(buildDimensionNode(sheet));
    
    // 添加工作表数据
    worksheet.addChild(buildSheetDataNode(sheet));
    
    // 添加合并单元格（如果有）
    auto mergeCellsNode = buildMergeCellsNode(sheet);
    if (!mergeCellsNode.getChildren().empty() || !mergeCellsNode.getAttributes().empty()) {
        worksheet.addChild(mergeCellsNode);
    }
    
    return worksheet;
}

XmlNodeBuilder TXWorksheetWriter::buildDimensionNode(const TXSheet* sheet) {
    TXRange usedRange = sheet->getUsedRange();
    
    XmlNodeBuilder dimension("dimension");
    if (usedRange.isValid()) {
        dimension.addAttribute("ref", usedRange.toAddress());
    } else {
        dimension.addAttribute("ref", "A1:A1");
    }
    
    return dimension;
}

XmlNodeBuilder TXWorksheetWriter::buildSheetDataNode(const TXSheet* sheet) {
    XmlNodeBuilder sheetData("sheetData");
    
    TXRange usedRange = sheet->getUsedRange();
    if (!usedRange.isValid()) {
        return sheetData;
    }
    
    // 遍历所有使用的行
    for (row_t row = usedRange.getStart().getRow(); row <= usedRange.getEnd().getRow(); ++row) {
        XmlNodeBuilder rowNode = buildRowNode(sheet, row, usedRange);
        
        // 只添加非空行
        if (!rowNode.getChildren().empty()) {
            sheetData.addChild(rowNode);
        }
    }
    
    return sheetData;
}

XmlNodeBuilder TXWorksheetWriter::buildRowNode(const TXSheet* sheet, row_t row, const TXRange& usedRange) {
    XmlNodeBuilder rowNode("row");
    rowNode.addAttribute("r", std::to_string(row.index()));
    
    // 遍历这一行的所有列
    for (column_t col = usedRange.getStart().getCol(); col <= usedRange.getEnd().getCol(); ++col) {
        const TXCell* cell = sheet->getCell(row, col);
        
        if (cell && !cell->isEmpty()) {
            std::string cellRef = column_t::column_string_from_index(col.index()) + std::to_string(row.index());
            XmlNodeBuilder cellNode = buildCellNode(cell, cellRef);
            rowNode.addChild(cellNode);
        }
    }
    
    return rowNode;
}

XmlNodeBuilder TXWorksheetWriter::buildCellNode(const TXCell* cell, const std::string& cellRef) {
    XmlNodeBuilder cellNode("c");
    cellNode.addAttribute("r", cellRef);
    
    // 根据单元格类型设置属性和值
    cell_value_t value = cell->getValue();
    
    if (std::holds_alternative<std::string>(value)) {
        // 字符串类型
        cellNode.addAttribute("t", "inlineStr");
        
        XmlNodeBuilder isNode("is");
        XmlNodeBuilder tNode("t");
        tNode.setText(std::get<std::string>(value));
        isNode.addChild(tNode);
        cellNode.addChild(isNode);
        
    } else if (std::holds_alternative<double>(value)) {
        // 浮点数类型
        XmlNodeBuilder vNode("v");
        vNode.setText(std::to_string(std::get<double>(value)));
        cellNode.addChild(vNode);
        
    } else if (std::holds_alternative<int64_t>(value)) {
        // 整数类型
        XmlNodeBuilder vNode("v");
        vNode.setText(std::to_string(std::get<int64_t>(value)));
        cellNode.addChild(vNode);
        
    } else if (std::holds_alternative<bool>(value)) {
        // 布尔类型
        cellNode.addAttribute("t", "b");
        XmlNodeBuilder vNode("v");
        vNode.setText(std::get<bool>(value) ? "1" : "0");
        cellNode.addChild(vNode);
    }
    
    return cellNode;
}

XmlNodeBuilder TXWorksheetWriter::buildMergeCellsNode(const TXSheet* sheet) {
    XmlNodeBuilder mergeCells("mergeCells");
    
    // 获取合并单元格列表（这需要在 TXSheet 中实现）
    // 暂时返回空节点
    // TODO: 实现合并单元格的获取和写入
    
    return mergeCells;
}

std::string TXWorksheetWriter::getWorksheetXmlPath(std::size_t sheetIndex) const {
    return "xl/worksheets/sheet" + std::to_string(sheetIndex) + ".xml";
}

const std::string& TXWorksheetWriter::getLastError() const {
    return lastError_;
}

} // namespace TinaXlsx