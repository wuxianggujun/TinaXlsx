//
// @file TXWorksheetReader.cpp - 新架构实现
//

#include "TinaXlsx/TXWorksheetReader.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXCoordinate.hpp"
#include "TinaXlsx/TXRange.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include <variant>
#include <regex>

namespace TinaXlsx {

TXWorksheetReader::TXWorksheetReader() 
    : lastError_(""), xmlReader_(std::make_unique<TXXmlReader>()) {}

TXWorksheetReader::~TXWorksheetReader() = default;

bool TXWorksheetReader::readWorksheetFromFile(const std::string& xlsxFilePath, TXSheet* sheet, std::size_t sheetIndex) {
    if (!sheet) {
        lastError_ = "Sheet is null";
        return false;
    }
    
    // 打开 ZIP 文件
    TXZipArchiveReader zipReader;
    if (!zipReader.open(xlsxFilePath)) {
        lastError_ = "Failed to open XLSX file: " + zipReader.lastError();
        return false;
    }
    
    // 构造工作表文件名
    std::string xmlPath = getWorksheetXmlPath(sheetIndex);
    
    // 使用 XML 读取器从 ZIP 中读取
    if (!xmlReader_->readFromZip(zipReader, xmlPath)) {
        lastError_ = "Failed to read worksheet XML: " + xmlReader_->getLastError();
        return false;
    }
    
    // 解析工作表数据
    return parseWorksheetData(sheet, *xmlReader_);
}

bool TXWorksheetReader::parseFromXml(TXSheet* sheet, const std::string& xmlContent) {
    if (!sheet) {
        lastError_ = "Sheet is null";
        return false;
    }
    
    // 重置 XML 读取器
    xmlReader_->reset();
    
    if (!xmlReader_->parseFromString(xmlContent)) {
        lastError_ = "Failed to parse XML: " + xmlReader_->getLastError();
        return false;
    }
    
    return parseWorksheetData(sheet, *xmlReader_);
}

bool TXWorksheetReader::parseWorksheetData(TXSheet* sheet, const TXXmlReader& xmlReader) {
    // 获取根节点
    auto rootNode = xmlReader.getRootNode();
    if (rootNode.name != "worksheet") {
        lastError_ = "Invalid worksheet XML: root node is not 'worksheet'";
        return false;
    }
    
    // 查找 sheetData 节点
    auto sheetDataNodes = xmlReader.findNodes("//sheetData");
    if (sheetDataNodes.empty()) {
        // 空工作表是允许的
        return true;
    }
    
    const auto& sheetDataNode = sheetDataNodes[0];
    
    // 解析每一行
    for (const auto& rowNode : sheetDataNode.children) {
        if (rowNode.name == "row") {
            if (!parseRowData(sheet, rowNode)) {
                return false;
            }
        }
    }
    
    // 解析合并单元格（如果存在）
    auto mergeCellsNodes = xmlReader.findNodes("//mergeCells");
    if (!mergeCellsNodes.empty()) {
        if (!parseMergeCells(sheet, mergeCellsNodes[0])) {
            return false;
        }
    }
    
    return true;
}

bool TXWorksheetReader::parseRowData(TXSheet* sheet, const XmlNodeInfo& rowNode) {
    // 解析行中的每个单元格
    for (const auto& cellNode : rowNode.children) {
        if (cellNode.name == "c") { // 'c' stands for cell
            if (!parseCellData(sheet, cellNode)) {
                return false;
            }
        }
    }
    return true;
}

bool TXWorksheetReader::parseCellData(TXSheet* sheet, const XmlNodeInfo& cellNode) {
    // 获取单元格引用 (r 属性)
    auto it = cellNode.attributes.find("r");
    if (it == cellNode.attributes.end()) {
        lastError_ = "Cell node missing 'r' attribute";
        return false;
    }
    
    std::string cellRef = it->second;
    TXCoordinate coord = parseAddress(cellRef);
    
    // 获取单元格类型 (t 属性)
    std::string cellType = "n"; // 默认为数字类型
    auto typeIt = cellNode.attributes.find("t");
    if (typeIt != cellNode.attributes.end()) {
        cellType = typeIt->second;
    }
    
    // 解析单元格值
    cell_value_t cellValue = parseCellValue(cellNode, cellType);
    
    // 设置到工作表中
    sheet->setCellValue(coord.getRow(), coord.getCol(), cellValue);
    
    return true;
}

bool TXWorksheetReader::parseMergeCells(TXSheet* sheet, const XmlNodeInfo& mergeCellsNode) {
    for (const auto& mergeCellNode : mergeCellsNode.children) {
        if (mergeCellNode.name == "mergeCell") {
            auto refIt = mergeCellNode.attributes.find("ref");
            if (refIt != mergeCellNode.attributes.end()) {
                std::string rangeRef = refIt->second;
                try {
                    TXRange range(rangeRef);
                    sheet->mergeCells(range);
                } catch (const std::exception& e) {
                    lastError_ = "Invalid merge cell range: " + rangeRef;
                    return false;
                }
            }
        }
    }
    return true;
}

cell_value_t TXWorksheetReader::parseCellValue(const XmlNodeInfo& cellNode, const std::string& cellType) {
    // 查找值节点
    for (const auto& child : cellNode.children) {
        if (child.name == "v") { // 'v' stands for value
            std::string valueText = child.value;
            
            if (cellType == "s") {
                // 共享字符串（需要进一步实现）
                return std::string(valueText);
            } else if (cellType == "str") {
                // 内联字符串
                return std::string(valueText);
            } else if (cellType == "b") {
                // 布尔值
                return valueText == "1";
            } else {
                // 数字类型
                try {
                    if (valueText.find('.') != std::string::npos) {
                        return std::stod(valueText);
                    } else {
                        return static_cast<int64_t>(std::stoll(valueText));
                    }
                } catch (const std::exception&) {
                    return std::string(valueText);
                }
            }
        }
    }
    
    // 如果没有找到值节点，返回空字符串
    return std::string("");
}

TXCoordinate TXWorksheetReader::parseAddress(const std::string& address) {
    // 简单的地址解析实现
    std::regex addrRegex(R"(([A-Z]+)(\d+))");
    std::smatch match;
    
    if (std::regex_match(address, match, addrRegex)) {
        std::string columnStr = match[1].str();
        int rowNum = std::stoi(match[2].str());
        
        // 转换列字母为列号
        column_t col(column_t::column_index_from_string(columnStr));
        row_t row(rowNum);
        
        return TXCoordinate(row, col);
    }
    
    // 解析失败，返回默认坐标
    return TXCoordinate(row_t(1), column_t(1));
}

std::string TXWorksheetReader::getWorksheetXmlPath(std::size_t sheetIndex) const {
    return "xl/worksheets/sheet" + std::to_string(sheetIndex) + ".xml";
}

const std::string& TXWorksheetReader::getLastError() const {
    return lastError_;
}

} // namespace TinaXlsx