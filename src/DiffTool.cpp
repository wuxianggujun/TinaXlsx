/**
 * @file DiffTool.cpp
 * @brief Excel差异对比工具实现
 */

#include "TinaXlsx/DiffTool.hpp"
#include "TinaXlsx/Writer.hpp"
#include "TinaXlsx/Reader.hpp"
#include "TinaXlsx/Worksheet.hpp"
#include "TinaXlsx/Utils.hpp"
#include <algorithm>
#include <unordered_map>
#include <unordered_set>

namespace TinaXlsx {

DiffTool::DiffTool(const CompareOptions& options)
    : options_(options) {
}

void DiffTool::setOptions(const CompareOptions& options) {
    options_ = options;
}

DiffTool::DiffResult DiffTool::compareData(
    const TableData& table1,
    const TableData& table2) {
    
    std::vector<RowDiff> rowDiffs = myersDiff(table1, table2);
    
    DiffResult result;
    result.rowDiffs = std::move(rowDiffs);
    
    // 统计差异
    for (const auto& diff : result.rowDiffs) {
        switch (diff.type) {
            case DiffType::Added: result.addedRowCount++; break;
            case DiffType::Deleted: result.deletedRowCount++; break;
            case DiffType::Modified: result.modifiedRowCount++; break;
            case DiffType::Unchanged: result.unchangedRowCount++; break;
        }
    }
    
    // 计算整体相似度
    if (result.rowDiffs.empty()) {
        result.overallSimilarity = 1.0;
    } else {
        result.overallSimilarity = static_cast<double>(result.unchangedRowCount) / result.rowDiffs.size();
    }
    
    return result;
}

DiffTool::DiffResult DiffTool::compareFiles(
    const std::string& file1, 
    const std::string& file2,
    const std::string& sheet1,
    const std::string& sheet2) {
    
    // 读取第一个文件
    Reader reader1(file1);
    if (!sheet1.empty()) {
        reader1.openSheet(sheet1);
    }
    TableData table1 = reader1.readAll();
    
    // 读取第二个文件
    Reader reader2(file2);
    if (!sheet2.empty()) {
        reader2.openSheet(sheet2);
    }
    TableData table2 = reader2.readAll();
    
    return compareData(table1, table2);
}

bool DiffTool::exportResult(
    const DiffResult& result,
    const std::string& outputFile,
    const TableData& table1,
    const TableData& table2,
    const std::vector<std::string>& columnHeaders) {
    
    try {
        Writer writer(outputFile);
        
        // 创建工作表
        auto summarySheet = writer.createWorksheet("差异摘要");
        
        // 创建格式
        auto headerFormat = writer.createFormat();
        headerFormat->setFontName("黑体")
            .setFontSize(11)
            .setBold()
            .setAlignment(Alignment::Center)
            .setBackgroundColor(0xD9D9D9)
            .setBorder(BorderStyle::Thin);
        
        // 简化实现，只创建基本的摘要信息
        summarySheet->writeString({0, 0}, "差异摘要", headerFormat.get());
        summarySheet->writeString({1, 0}, "新增行数");
        summarySheet->writeInteger({1, 1}, static_cast<int>(result.addedRowCount));
        summarySheet->writeString({2, 0}, "删除行数");
        summarySheet->writeInteger({2, 1}, static_cast<int>(result.deletedRowCount));
        summarySheet->writeString({3, 0}, "修改行数");
        summarySheet->writeInteger({3, 1}, static_cast<int>(result.modifiedRowCount));
        summarySheet->writeString({4, 0}, "相同行数");
        summarySheet->writeInteger({4, 1}, static_cast<int>(result.unchangedRowCount));
        
        return writer.save();
    }
    catch (const std::exception&) {
        return false;
    }
}

bool DiffTool::exportHtmlReport(
    const DiffResult& result,
    const TableData& table1,
    const TableData& table2,
    const std::string& outputFile,
    const std::string& title) {
    
    // 简化实现，暂时返回false
    return false;
}

double DiffTool::calculateRowSimilarity(const RowData& row1, const RowData& row2) const {
    if (row1.empty() && row2.empty()) {
        return 1.0;
    }
    
    if (row1.empty() || row2.empty()) {
        return 0.0;
    }
    
    size_t maxSize = std::max(row1.size(), row2.size());
    size_t minSize = std::min(row1.size(), row2.size());
    
    double similarity = 0.0;
    for (size_t i = 0; i < minSize; ++i) {
        if (areValuesEqual(row1[i], row2[i])) {
            similarity += 1.0;
        }
    }
    
    // 考虑长度差异的惩罚
    return similarity / maxSize;
}

double DiffTool::calculateStringSimilarity(const std::string& str1, const std::string& str2) const {
    if (str1 == str2) {
        return 1.0;
    }
    
    if (str1.empty() || str2.empty()) {
        return 0.0;
    }
    
    // 简化的相似度计算
    size_t maxLen = std::max(str1.length(), str2.length());
    size_t commonChars = 0;
    
    for (size_t i = 0; i < std::min(str1.length(), str2.length()); ++i) {
        if (str1[i] == str2[i]) {
            commonChars++;
        }
    }
    
    return static_cast<double>(commonChars) / maxLen;
}

// 私有方法实现
std::vector<DiffTool::RowDiff> DiffTool::myersDiff(const TableData& table1, const TableData& table2) const {
    std::vector<RowDiff> result;
    
    size_t n = table1.size();
    size_t m = table2.size();
    
    // 简化的逐行比较实现
    size_t i = 0, j = 0;
    
    while (i < n && j < m) {
        double similarity = calculateRowSimilarity(table1[i], table2[j]);
        
        if (similarity >= options_.similarityThreshold) {
            // 相似行
            RowDiff diff;
            diff.oldRowIndex = i;
            diff.newRowIndex = j;
            diff.type = (table1[i] == table2[j]) ? DiffType::Unchanged : DiffType::Modified;
            diff.oldRowData = table1[i];
            diff.newRowData = table2[j];
            diff.similarity = similarity;
            
            if (options_.enableCellLevelDiff && diff.type == DiffType::Modified) {
                diff.cellDiffs = calculateCellDiffs(table1[i], table2[j], i);
            }
            
            result.push_back(diff);
            i++;
            j++;
        } else {
            // 处理不匹配的行
            RowDiff diff1;
            diff1.oldRowIndex = i;
            diff1.newRowIndex = SIZE_MAX;
            diff1.type = DiffType::Deleted;
            diff1.oldRowData = table1[i];
            diff1.similarity = 0.0;
            result.push_back(diff1);
            
            RowDiff diff2;
            diff2.oldRowIndex = SIZE_MAX;
            diff2.newRowIndex = j;
            diff2.type = DiffType::Added;
            diff2.newRowData = table2[j];
            diff2.similarity = 0.0;
            result.push_back(diff2);
            
            i++;
            j++;
        }
    }
    
    // 处理剩余的行
    while (i < n) {
        RowDiff diff;
        diff.oldRowIndex = i;
        diff.newRowIndex = SIZE_MAX;
        diff.type = DiffType::Deleted;
        diff.oldRowData = table1[i];
        diff.similarity = 0.0;
        result.push_back(diff);
        i++;
    }
    
    while (j < m) {
        RowDiff diff;
        diff.oldRowIndex = SIZE_MAX;
        diff.newRowIndex = j;
        diff.type = DiffType::Added;
        diff.newRowData = table2[j];
        diff.similarity = 0.0;
        result.push_back(diff);
        j++;
    }
    
    return result;
}

std::vector<std::pair<int, int>> DiffTool::findBestMatches(const TableData& table1, const TableData& table2) const {
    std::vector<std::pair<int, int>> matches;
    
    for (size_t i = 0; i < table1.size(); ++i) {
        double bestSimilarity = 0.0;
        int bestMatch = -1;
        
        for (size_t j = 0; j < table2.size(); ++j) {
            double similarity = calculateRowSimilarity(table1[i], table2[j]);
            if (similarity > bestSimilarity && similarity >= options_.similarityThreshold) {
                bestSimilarity = similarity;
                bestMatch = j;
            }
        }
        
        if (bestMatch != -1) {
            matches.emplace_back(i, bestMatch);
        }
    }
    
    return matches;
}

std::vector<DiffTool::CellDiff> DiffTool::calculateCellDiffs(const RowData& row1, const RowData& row2, RowIndex rowIndex) const {
    std::vector<CellDiff> diffs;
    
    size_t maxCols = std::max(row1.size(), row2.size());
    
    for (size_t col = 0; col < maxCols; ++col) {
        CellValue oldValue = (col < row1.size()) ? row1[col] : CellValue{};
        CellValue newValue = (col < row2.size()) ? row2[col] : CellValue{};
        
        if (!areValuesEqual(oldValue, newValue)) {
            CellDiff diff;
            diff.position = {rowIndex, static_cast<ColumnIndex>(col)};
            diff.oldValue = oldValue;
            diff.newValue = newValue;
            
            if (col >= row1.size()) {
                diff.type = DiffType::Added;
            } else if (col >= row2.size()) {
                diff.type = DiffType::Deleted;
            } else {
                diff.type = DiffType::Modified;
            }
            
            diffs.push_back(diff);
        }
    }
    
    return diffs;
}

CellValue DiffTool::preprocessCellValue(const CellValue& value) const {
    return value;
}

TableData DiffTool::preprocessTable(const TableData& table) const {
    TableData result = table;
    
    if (options_.ignoreEmptyRows) {
        result.erase(
            std::remove_if(result.begin(), result.end(),
                [this](const RowData& row) { return isEmptyRow(row); }),
            result.end());
    }
    
    return result;
}

bool DiffTool::areValuesEqual(const CellValue& value1, const CellValue& value2) const {
    if (options_.ignoreCase) {
        // 如果忽略大小写，将字符串转换为小写比较
        auto str1 = Utils::Convert::cellValueToString(value1);
        auto str2 = Utils::Convert::cellValueToString(value2);
        return Utils::String::toLower(str1) == Utils::String::toLower(str2);
    }
    
    return value1 == value2;
}

bool DiffTool::isEmptyRow(const RowData& row) const {
    return std::all_of(row.begin(), row.end(), [](const CellValue& cell) {
        return std::holds_alternative<std::monostate>(cell) ||
               (std::holds_alternative<std::string>(cell) && std::get<std::string>(cell).empty());
    });
}

} // namespace TinaXlsx 