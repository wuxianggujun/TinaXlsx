/**
 * @file test_reader_comprehensive.cpp
 * @brief TinaXlsx Reader 全面功能测试
 * 整合所有Reader相关的测试，包括基础功能、错误处理、性能等
 */

#include <gtest/gtest.h>
#include <TinaXlsx/Reader.hpp>
#include <TinaXlsx/Writer.hpp>
#include <TinaXlsx/Exception.hpp>
#include <filesystem>
#include <fstream>
#include <chrono>

using namespace TinaXlsx;

class ReaderComprehensiveTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 创建测试目录
        testDir_ = std::filesystem::temp_directory_path() / "tinaxlsx_test";
        std::filesystem::create_directories(testDir_);
        
        // 创建测试文件路径
        simpleTestFile_ = testDir_ / "simple_test.xlsx";
        largeTestFile_ = testDir_ / "large_test.xlsx";
        
        // 创建简单测试文件
        createSimpleTestFile();
    }
    
    void TearDown() override {
        // 清理测试文件
        std::error_code ec;
        std::filesystem::remove_all(testDir_, ec);
    }
    
    void createSimpleTestFile() {
        try {
            Writer writer(simpleTestFile_.string());
            
            // 写入简单数据
            writer.writeCell(0, 0, "姓名");
            writer.writeCell(0, 1, "年龄");
            writer.writeCell(0, 2, "分数");
            
            writer.writeCell(1, 0, "张三");
            writer.writeCell(1, 1, 25);
            writer.writeCell(1, 2, 95.5);
            
            writer.writeCell(2, 0, "李四");
            writer.writeCell(2, 1, 30);
            writer.writeCell(2, 2, 87.3);
            
            writer.writeCell(3, 0, "王五");
            writer.writeCell(3, 1, 28);
            writer.writeCell(3, 2, 92.1);
            
            // 空行
            writer.writeCell(5, 0, "总计");
            writer.writeCell(5, 2, 274.9);
            
            writer.save();
        } catch (const std::exception& e) {
            FAIL() << "Failed to create test file: " << e.what();
        }
    }
    
    void createLargeTestFile(size_t rows = 10000, size_t cols = 50) {
        try {
            Writer writer(largeTestFile_.string());
            
            // 写入大量数据用于性能测试
            for (size_t row = 0; row < rows; ++row) {
                for (size_t col = 0; col < cols; ++col) {
                    if (row == 0) {
                        writer.writeCell(row, col, "Column" + std::to_string(col));
                    } else {
                        writer.writeCell(row, col, row * cols + col);
                    }
                }
            }
            
            writer.save();
        } catch (const std::exception& e) {
            FAIL() << "Failed to create large test file: " << e.what();
        }
    }

protected:
    std::filesystem::path testDir_;
    std::filesystem::path simpleTestFile_;
    std::filesystem::path largeTestFile_;
};

// =============================================================================
// 基础功能测试
// =============================================================================

TEST_F(ReaderComprehensiveTest, BasicConstruction) {
    // 测试正常构造
    EXPECT_NO_THROW({
        Reader reader(simpleTestFile_.string());
    });
    
    // 测试不存在的文件
    EXPECT_THROW({
        Reader reader("nonexistent_file.xlsx");
    }, FileException);
}

TEST_F(ReaderComprehensiveTest, GetSheetNames) {
    Reader reader(simpleTestFile_.string());
    
    auto sheetNames = reader.getWorksheetNames();
    EXPECT_FALSE(sheetNames.empty());
    EXPECT_EQ(sheetNames.size(), 1);
    EXPECT_EQ(sheetNames[0], "Sheet1");
}

TEST_F(ReaderComprehensiveTest, OpenSheet) {
    Reader reader(simpleTestFile_.string());
    
    // 通过名称打开工作表
    EXPECT_TRUE(reader.openWorksheet("Sheet1"));
    EXPECT_FALSE(reader.openWorksheet("NonexistentSheet"));
    
    // 通过索引打开工作表
    EXPECT_TRUE(reader.openWorksheet(0));
    EXPECT_FALSE(reader.openWorksheet(999));
}

TEST_F(ReaderComprehensiveTest, GetDimensions) {
    Reader reader(simpleTestFile_.string());
    ASSERT_TRUE(reader.openWorksheet(0));
    
    auto dimensions = reader.getDimensions();
    
    EXPECT_TRUE(dimensions.has_value());
    
    if (dimensions) {
        EXPECT_GE(dimensions->first, 6);  // 至少6行（0-5）
        EXPECT_GE(dimensions->second, 3);  // 至少3列
    }
}

// =============================================================================
// 数据读取测试
// =============================================================================

TEST_F(ReaderComprehensiveTest, ReadSpecificCell) {
    Reader reader(simpleTestFile_.string());
    ASSERT_TRUE(reader.openWorksheet(0));
    
    // 读取标题行
    auto cell00 = reader.readCell({0, 0});
    ASSERT_TRUE(cell00.has_value());
    EXPECT_EQ(Reader::cellValueToString(*cell00), "姓名");
    
    auto cell01 = reader.readCell({0, 1});
    ASSERT_TRUE(cell01.has_value());
    EXPECT_EQ(Reader::cellValueToString(*cell01), "年龄");
    
    // 读取数据行
    auto cell10 = reader.readCell({1, 0});
    ASSERT_TRUE(cell10.has_value());
    EXPECT_EQ(Reader::cellValueToString(*cell10), "张三");
    
    auto cell11 = reader.readCell({1, 1});
    ASSERT_TRUE(cell11.has_value());
    EXPECT_EQ(Reader::cellValueToString(*cell11), "25");
    
    auto cell12 = reader.readCell({1, 2});
    ASSERT_TRUE(cell12.has_value());
    EXPECT_EQ(Reader::cellValueToString(*cell12), "95.5");
}

TEST_F(ReaderComprehensiveTest, ReadRowData) {
    Reader reader(simpleTestFile_.string());
    ASSERT_TRUE(reader.openWorksheet(0));
    
    // 读取第一行（标题行）
    auto row0 = reader.readRow(0);
    ASSERT_TRUE(row0.has_value());
    EXPECT_GE(row0->size(), 3);
    
    if (row0->size() >= 3) {
        EXPECT_EQ(Reader::cellValueToString((*row0)[0]), "姓名");
        EXPECT_EQ(Reader::cellValueToString((*row0)[1]), "年龄");
        EXPECT_EQ(Reader::cellValueToString((*row0)[2]), "分数");
    }
    
    // 读取数据行
    auto row1 = reader.readRow(1);
    ASSERT_TRUE(row1.has_value());
    EXPECT_GE(row1->size(), 3);
    
    if (row1->size() >= 3) {
        EXPECT_EQ(Reader::cellValueToString((*row1)[0]), "张三");
        EXPECT_EQ(Reader::cellValueToString((*row1)[1]), "25");
        EXPECT_EQ(Reader::cellValueToString((*row1)[2]), "95.5");
    }
}

TEST_F(ReaderComprehensiveTest, ReadRangeData) {
    Reader reader(simpleTestFile_.string());
    ASSERT_TRUE(reader.openWorksheet(0));
    
    // 读取指定范围
    CellRange range{{0, 0}, {2, 2}};  // A1:C3
    auto rangeData = reader.readRange(range);
    
    EXPECT_EQ(rangeData.size(), 3);  // 3行
    
    for (const auto& row : rangeData) {
        EXPECT_GE(row.size(), 3);  // 至少3列
    }
    
    // 验证范围内的数据
    if (!rangeData.empty() && !rangeData[0].empty()) {
        EXPECT_EQ(Reader::cellValueToString(rangeData[0][0]), "姓名");
        EXPECT_EQ(Reader::cellValueToString(rangeData[1][0]), "张三");
        EXPECT_EQ(Reader::cellValueToString(rangeData[2][0]), "李四");
    }
}

TEST_F(ReaderComprehensiveTest, ReadAllData) {
    Reader reader(simpleTestFile_.string());
    ASSERT_TRUE(reader.openWorksheet(0));
    
    // 读取所有数据
    auto allData = reader.readAllData();
    ASSERT_TRUE(allData.has_value());
    EXPECT_GT(allData->size(), 0);
    
    // 验证数据内容
    if (allData->size() > 0) {
        const auto& firstRow = (*allData)[0];
        EXPECT_GE(firstRow.size(), 3);
        
        if (firstRow.size() >= 3) {
            EXPECT_EQ(Reader::cellValueToString(firstRow[0]), "姓名");
        }
    }
}

// =============================================================================
// 流式读取测试
// =============================================================================

TEST_F(ReaderComprehensiveTest, StreamingReadNextRow) {
    Reader reader(simpleTestFile_.string());
    ASSERT_TRUE(reader.openWorksheet(0));
    
    // 逐行读取
    size_t rowCount = 0;
    std::optional<RowData> row;
    
    while ((row = reader.readRow(rowCount)).has_value()) {
        EXPECT_FALSE(row->empty());
        rowCount++;
        
        // 防止无限循环
        if (rowCount > 100) break;
    }
    
    EXPECT_GT(rowCount, 0);
}

TEST_F(ReaderComprehensiveTest, StreamingReadWithCallback) {
    Reader reader(simpleTestFile_.string());
    ASSERT_TRUE(reader.openWorksheet(0));
    
    size_t callbackRowCount = 0;
    size_t totalCells = 0;
    
    // 使用行回调
    auto rowsRead = reader.readWithCallback(
        [&](const CellPosition& pos, const CellValue& value) -> bool {
            return true; // 继续处理
        },
        [&](RowIndex rowIndex, const RowData& rowData) -> bool {
            callbackRowCount++;
            return true; // 继续处理
        }
    );
    
    EXPECT_GT(callbackRowCount, 0);
    EXPECT_GT(totalCells, 0);
    EXPECT_EQ(rowsRead, callbackRowCount);
}

TEST_F(ReaderComprehensiveTest, StreamingReadCellsWithCallback) {
    Reader reader(simpleTestFile_.string());
    ASSERT_TRUE(reader.openWorksheet(0));
    
    size_t cellCount = 0;
    std::vector<std::pair<CellPosition, CellValue>> cells;
    
    // 使用单元格回调
    auto cellsRead = reader.readAllCells([&](const CellPosition& position, const CellValue& value) -> bool {
        cellCount++;
        cells.emplace_back(position, value);
        
        // 验证位置合理性
        EXPECT_GE(position.row, 0);
        EXPECT_GE(position.column, 0);
        
        return true; // 继续读取
    });
    
    EXPECT_GT(cellCount, 0);
    EXPECT_EQ(cellsRead, cellCount);
    EXPECT_EQ(cells.size(), cellCount);
}

// =============================================================================
// 错误处理测试
// =============================================================================

TEST_F(ReaderComprehensiveTest, ErrorHandling) {
    Reader reader(simpleTestFile_.string());
    ASSERT_TRUE(reader.openWorksheet(0));
    
    // 读取超出范围的单元格
    auto invalidCell = reader.readCell({999, 999});
    EXPECT_FALSE(invalidCell.has_value());
    
    // 读取超出范围的行
    auto invalidRow = reader.readRow(999);
    EXPECT_FALSE(invalidRow.has_value());
    
    // 读取无效范围
    CellRange invalidRange{{999, 999}, {1000, 1000}};
    auto invalidRangeData = reader.readRange(invalidRange);
    EXPECT_TRUE(invalidRangeData.empty() || 
                std::all_of(invalidRangeData.begin(), invalidRangeData.end(), 
                           [](const RowData& row) { return row.empty(); }));
}

TEST_F(ReaderComprehensiveTest, FileNotFound) {
    EXPECT_THROW({
        Reader reader("totally_nonexistent_file.xlsx");
    }, FileException);
}

TEST_F(ReaderComprehensiveTest, InvalidFileFormat) {
    // 创建一个非Excel文件
    auto invalidFile = testDir_ / "invalid.xlsx";
    std::ofstream ofs(invalidFile, std::ios::binary);
    ofs << "This is not an Excel file";
    ofs.close();
    
    EXPECT_THROW({
        Reader reader(invalidFile.string());
    }, FileException);
}

// =============================================================================
// 数据类型测试
// =============================================================================

TEST_F(ReaderComprehensiveTest, DataTypeConversion) {
    Reader reader(simpleTestFile_.string());
    ASSERT_TRUE(reader.openWorksheet(0));
    
    // 测试字符串转换
    std::string testStr = "测试字符串";
    auto cellValue1 = Reader::stringToCellValue(testStr);
    EXPECT_EQ(Reader::cellValueToString(cellValue1), testStr);
    
    // 测试数字转换
    std::string numStr = "123.45";
    auto cellValue2 = Reader::stringToCellValue(numStr);
    std::string convertedBack = Reader::cellValueToString(cellValue2);
    // 数字可能在转换过程中格式化，所以检查数值是否相等
    EXPECT_NEAR(std::stod(convertedBack), 123.45, 0.01);
    
    // 测试整数转换
    std::string intStr = "42";
    auto cellValue3 = Reader::stringToCellValue(intStr);
    EXPECT_EQ(Reader::cellValueToString(cellValue3), intStr);
}

TEST_F(ReaderComprehensiveTest, EmptyDataDetection) {
    Reader reader(simpleTestFile_.string());
    ASSERT_TRUE(reader.openWorksheet(0));
    
    // 测试空单元格
    CellValue emptyValue = std::monostate{};
    EXPECT_TRUE(Reader::isEmptyCell(emptyValue));
    
    CellValue stringValue = std::string("test");
    EXPECT_FALSE(Reader::isEmptyCell(stringValue));
    
    // 测试空行检测
    RowData emptyRow;
    EXPECT_TRUE(Reader::isEmptyRow(emptyRow));
    
    RowData rowWithEmpty = {std::monostate{}, std::monostate{}};
    EXPECT_TRUE(Reader::isEmptyRow(rowWithEmpty));
    
    RowData rowWithData = {std::string("data"), std::monostate{}};
    EXPECT_FALSE(Reader::isEmptyRow(rowWithData));
}

// =============================================================================
// 性能测试
// =============================================================================

TEST_F(ReaderComprehensiveTest, DISABLED_PerformanceTest) {
    const size_t testRows = 1000;
    const size_t testCols = 20;
    
    // 创建大文件
    createLargeTestFile(testRows, testCols);
    
    // 测试读取性能
    auto startTime = std::chrono::high_resolution_clock::now();
    
    {
        Reader reader(largeTestFile_.string());
        ASSERT_TRUE(reader.openWorksheet(0));
        
        auto allData = reader.readAllData();
        auto endTime = std::chrono::high_resolution_clock::now();
        
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(endTime - startTime);
        
        ASSERT_TRUE(allData.has_value());
        EXPECT_GT(allData->size(), 0);
        
        std::cout << "读取 " << allData->size() << " 行数据耗时: " 
                  << duration.count() << " ms" << std::endl;
    }
    
    // 性能要求：读取1000行x20列数据应该在1秒内完成
    EXPECT_LT(duration.count(), 1000) << "读取性能不符合要求";
}

TEST_F(ReaderComprehensiveTest, DISABLED_StreamingPerformanceTest) {
    const size_t testRows = 1000;
    const size_t testCols = 20;
    
    createLargeTestFile(testRows, testCols);
    
    // 测试流式读取性能
    auto start = std::chrono::high_resolution_clock::now();
    
    {
        Reader reader(largeTestFile_.string());
        ASSERT_TRUE(reader.openWorksheet(0));
        
        size_t rowCount = 0;
        size_t cellCount = 0;
        
        auto processedCells = reader.readWithCallback(
            [&](const CellPosition& pos, const CellValue& value) -> bool {
                cellCount++;
                return true;
            },
            [&](RowIndex rowIndex, const RowData& rowData) -> bool {
                rowCount++;
                return true;
            }
        );
        
        auto endTime = std::chrono::high_resolution_clock::now();
        auto duration = std::chrono::duration_cast<std::chrono::milliseconds>(endTime - start);
        
        EXPECT_GT(rowCount, 0);
        EXPECT_GT(cellCount, 0);
        
        std::cout << "流式读取 " << rowCount << " 行, " << cellCount 
                  << " 个单元格耗时: " << duration.count() << " ms" << std::endl;
    }
    
    // 流式读取应该更快
    EXPECT_LT(duration.count(), 800) << "流式读取性能不符合要求";
}

// =============================================================================
// 内存测试
// =============================================================================

TEST_F(ReaderComprehensiveTest, DISABLED_MemoryUsageTest) {
    const size_t testRows = 5000;
    const size_t testCols = 50;
    
    createLargeTestFile(testRows, testCols);
    
    // 多次创建和销毁Reader，检查内存泄漏
    for (int i = 0; i < 10; ++i) {
        Reader reader(largeTestFile_.string());
        ASSERT_TRUE(reader.openWorksheet(0));
        
        // 读取部分数据
        auto partialData = reader.readRange({{0, 0}, {100, 10}});
        ASSERT_TRUE(partialData.has_value());
        EXPECT_GT(partialData->size(), 0);
    }
    
    // 如果有内存泄漏，这个测试在调试模式下应该会被检测到
    SUCCEED() << "内存使用测试完成";
}

// =============================================================================
// 特殊情况测试
// =============================================================================

TEST_F(ReaderComprehensiveTest, EmptyFileHandling) {
    // 创建一个只有标题的Excel文件
    auto emptyFile = testDir_ / "empty_content.xlsx";
    
    try {
        Writer writer(emptyFile.string());
        writer.writeCell(0, 0, "Empty");
        writer.save();
        
        Reader reader(emptyFile.string());
        ASSERT_TRUE(reader.openWorksheet(0));
        
        auto allData = reader.readAllData();
        
        // 空文件应该返回空数据或者只有空行
        if (allData.has_value()) {
            // 如果有数据，应该都是空行
            for (const auto& row : *allData) {
                EXPECT_TRUE(Reader::isEmptyRow(row));
            }
        }
    } catch (const std::exception& e) {
        FAIL() << "Empty file test failed: " << e.what();
    }
}

TEST_F(ReaderComprehensiveTest, UnicodeHandling) {
    Reader reader(simpleTestFile_.string());
    ASSERT_TRUE(reader.openWorksheet(0));
    
    // 验证中文字符正确读取
    auto cell00 = reader.readCell({0, 0});
    ASSERT_TRUE(cell00.has_value());
    std::string value = Reader::cellValueToString(*cell00);
    EXPECT_EQ(value, "姓名");
    
    auto cell10 = reader.readCell({1, 0});
    ASSERT_TRUE(cell10.has_value());
    value = Reader::cellValueToString(*cell10);
    EXPECT_EQ(value, "张三");
}

// =============================================================================
// 主测试入口
// =============================================================================

int main(int argc, char** argv) {
    ::testing::InitGoogleTest(&argc, argv);
    
    std::cout << "=== TinaXlsx Reader 全面功能测试 ===" << std::endl;
    std::cout << "测试包括：" << std::endl;
    std::cout << "- 基础功能（构造、打开工作表、获取维度）" << std::endl;
    std::cout << "- 数据读取（单元格、行、范围、全部数据）" << std::endl;
    std::cout << "- 流式读取（逐行、回调函数）" << std::endl;
    std::cout << "- 错误处理（文件不存在、格式错误、越界访问）" << std::endl;
    std::cout << "- 数据类型转换" << std::endl;
    std::cout << "- 性能测试（默认关闭，可用--gtest_also_run_disabled_tests启用）" << std::endl;
    std::cout << "- Unicode支持" << std::endl;
    std::cout << std::endl;
    
    return RUN_ALL_TESTS();
} 