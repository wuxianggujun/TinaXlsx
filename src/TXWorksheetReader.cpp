#include "TinaXlsx/TXWorksheetReader.hpp"
#include "TinaXlsx/TXSheet.hpp"
#include "TinaXlsx/TXCell.hpp"
#include "TinaXlsx/TXCoordinate.hpp"
#include "TinaXlsx/TXRange.hpp"
#include "TinaXlsx/TXZipArchive.hpp"
#include "TinaXlsx/TXXmlReader.hpp"
#include <variant>
#include <regex>

namespace TinaXlsx {

TXWorksheetReader::TXWorksheetReader() : lastError_("") {}

TXWorksheetReader::~TXWorksheetReader() = default;

bool TXWorksheetReader::readFromZip(TXZipArchiveReader& zip, TXSheet* sheet, std::size_t sheetIndex) {
    if (!sheet) {
        lastError_ = "Sheet is null";
        return false;
    }
    
    // 构造工作表文件名
    std::string filename = "xl/worksheets/sheet" + std::to_string(sheetIndex) + ".xml";
    
    if (!zip.has(filename)) {
        lastError_ = "Worksheet file not found: " + filename;
        return false;
    }
    
    std::string xmlContent = zip.readString(filename);
    if (xmlContent.empty()) {
        lastError_ = "Failed to read worksheet file: " + zip.lastError();
        return false;
    }
    
    return parseFromXml(sheet, xmlContent);
}

bool TXWorksheetReader::parseFromXml(TXSheet* sheet, const std::string& xmlContent) {
    if (!sheet) {
        lastError_ = "Sheet is null";
        return false;
    }
    
    TXXmlReader reader;
    if (!reader.loadFromString(xmlContent)) {
        lastError_ = "Failed to parse XML: " + reader.getLastError();
        return false;
    }
    
    const TXXmlNode* rootNode = reader.getRootNode();
    if (!rootNode || rootNode->name != "worksheet") {
        lastError_ = "Invalid worksheet XML: missing root element";
        return false;
    }
    
    return parseWorksheetData(sheet, rootNode);
}

const std::string& TXWorksheetReader::getLastError() const {
    return lastError_;
}

bool TXWorksheetReader::parseWorksheetData(TXSheet* sheet, const TXXmlNode* rootNode) {
    // 解析sheetData节点
    const TXXmlNode* sheetDataNode = rootNode->findChild("sheetData");
    if (sheetDataNode) {
        // 遍历所有行
        auto rowNodes = sheetDataNode->findChildren("row");
        for (const auto* rowNode : rowNodes) {
            if (!parseRowData(sheet, rowNode)) {
                return false;
            }
        }
    }
    
    // 解析合并单元格
    const TXXmlNode* mergeCellsNode = rootNode->findChild("mergeCells");
    if (mergeCellsNode) {
        if (!parseMergeCells(sheet, mergeCellsNode)) {
            return false;
        }
    }
    
    lastError_.clear();
    return true;
}

bool TXWorksheetReader::parseRowData(TXSheet* sheet, const TXXmlNode* rowNode) {
    if (!rowNode) return false;
    
    // 遍历行中的所有单元格
    auto cellNodes = rowNode->findChildren("c");
    for (const auto* cellNode : cellNodes) {
        if (!parseCellData(sheet, cellNode)) {
            return false;
        }
    }
    
    return true;
}

bool TXWorksheetReader::parseCellData(TXSheet* sheet, const TXXmlNode* cellNode) {
    if (!cellNode) return false;
    
    // 获取单元格地址
    std::string cellAddress = cellNode->getAttribute("r");
    if (cellAddress.empty()) {
        lastError_ = "Cell missing address attribute";
        return false;
    }
    
    TXCoordinate coord = parseAddress(cellAddress);
    if (!coord.isValid()) {
        lastError_ = "Invalid cell address: " + cellAddress;
        return false;
    }
    
    // 获取单元格类型
    std::string cellType = cellNode->getAttribute("t", "n"); // 默认为数值类型
    
    // 解析单元格值
    cell_value_t value = parseCellValue(cellNode, cellType);
    
    // 将值设置到工作表中
    if (std::holds_alternative<std::string>(value)) {
        sheet->setCellValue(coord.getRow(), coord.getCol(), std::get<std::string>(value));
    } else if (std::holds_alternative<double>(value)) {
        sheet->setCellValue(coord.getRow(), coord.getCol(), std::get<double>(value));
    } else if (std::holds_alternative<int64_t>(value)) {
        sheet->setCellValue(coord.getRow(), coord.getCol(), std::get<int64_t>(value));
    } else if (std::holds_alternative<bool>(value)) {
        sheet->setCellValue(coord.getRow(), coord.getCol(), std::get<bool>(value));
    }
    
    return true;
}

bool TXWorksheetReader::parseMergeCells(TXSheet* sheet, const TXXmlNode* mergeCellsNode) {
    if (!mergeCellsNode) return false;
    
    auto mergeCellNodes = mergeCellsNode->findChildren("mergeCell");
    for (const auto* mergeCellNode : mergeCellNodes) {
        std::string rangeRef = mergeCellNode->getAttribute("ref");
        if (!rangeRef.empty()) {
            // 解析范围并添加到工作表
            TXRange range = TXRange::fromAddress(rangeRef);
            if (range.isValid()) {
                sheet->mergeCells(range);
            }
        }
    }
    
    return true;
}

cell_value_t TXWorksheetReader::parseCellValue(const TXXmlNode* cellNode, const std::string& cellType) {
    if (cellType == "inlineStr") {
        // 内联字符串
        const TXXmlNode* isNode = cellNode->findChild("is");
        if (isNode) {
            const TXXmlNode* tNode = isNode->findChild("t");
            if (tNode) {
                return cell_value_t(tNode->text);
            }
        }
    } else if (cellType == "str") {
        // 字符串（通过共享字符串表）
        const TXXmlNode* vNode = cellNode->findChild("v");
        if (vNode) {
            // 这里应该查找共享字符串表，简化处理直接返回文本
            return cell_value_t(vNode->text);
        }
    } else if (cellType == "b") {
        // 布尔值
        const TXXmlNode* vNode = cellNode->findChild("v");
        if (vNode) {
            return cell_value_t(vNode->text == "1");
        }
    } else {
        // 数值类型 (cellType == "n" 或默认)
        const TXXmlNode* vNode = cellNode->findChild("v");
        if (vNode) {
            try {
                // 尝试解析为整数
                if (vNode->text.find('.') == std::string::npos) {
                    return cell_value_t(std::stoll(vNode->text));
                } else {
                    // 包含小数点，解析为浮点数
                    return cell_value_t(std::stod(vNode->text));
                }
            } catch (const std::exception&) {
                // 解析失败，返回字符串
                return cell_value_t(vNode->text);
            }
        }
    }
    
    // 默认返回空字符串
    return cell_value_t(std::string(""));
}

TXCoordinate TXWorksheetReader::parseAddress(const std::string& address) {
    // 使用正则表达式解析单元格地址，例如 "A1", "AB123"
    std::regex addressRegex(R"(([A-Z]+)(\d+))");
    std::smatch match;
    
    if (std::regex_match(address, match, addressRegex)) {
        std::string colStr = match[1].str();
        std::string rowStr = match[2].str();
        
        // 解析列号（A=1, B=2, ..., Z=26, AA=27, ...）
        column_t col(colStr);
        
        // 解析行号
        row_t row(static_cast<row_t::index_t>(std::stoi(rowStr)));
        
        return TXCoordinate(row, col);
    }
    
    // 解析失败，返回无效坐标
    return TXCoordinate();
}

} // namespace TinaXlsx 