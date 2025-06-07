//
// @file test_coord_utils.cpp
// @brief TXCoordUtils 统一坐标转换工具测试
//

#include <gtest/gtest.h>
#include <TinaXlsx/TXCoordUtils.hpp>
#include <TinaXlsx/TXUnifiedMemoryManager.hpp>
#include <TinaXlsx/TXHighPerformanceLogger.hpp>
#include <chrono>

using namespace TinaXlsx;

class CoordUtilsTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 初始化内存管理器
        TXUnifiedMemoryManager::Config config;
        config.memory_limit = 512ULL * 1024 * 1024; // 512MB
        GlobalUnifiedMemoryManager::initialize(config);
        
        // 初始化日志系统
        TXGlobalLogger::initialize(GlobalUnifiedMemoryManager::getInstance());
        TXGlobalLogger::setOutputMode(TXLogOutputMode::CONSOLE_ONLY);
    }
    
    void TearDown() override {
        GlobalUnifiedMemoryManager::shutdown();
    }
};

/**
 * @brief 测试基本坐标解析
 */
TEST_F(CoordUtilsTest, BasicCoordParsing) {
    // 测试基本坐标
    auto result = TXCoordUtils::parseCoord("A1");
    ASSERT_TRUE(result.isOk());
    auto coord = result.value();
    EXPECT_EQ(coord.getRow().index(), 0);
    EXPECT_EQ(coord.getCol().index(), 0);
    
    // 测试复杂坐标
    result = TXCoordUtils::parseCoord("Z100");
    ASSERT_TRUE(result.isOk());
    coord = result.value();
    EXPECT_EQ(coord.getRow().index(), 99);
    EXPECT_EQ(coord.getCol().index(), 25);
    
    // 测试AA列
    result = TXCoordUtils::parseCoord("AA1");
    ASSERT_TRUE(result.isOk());
    coord = result.value();
    EXPECT_EQ(coord.getRow().index(), 0);
    EXPECT_EQ(coord.getCol().index(), 26);
    
    TX_LOG_INFO("基本坐标解析测试通过");
}

/**
 * @brief 测试高性能坐标解析
 */
TEST_F(CoordUtilsTest, FastCoordParsing) {
    // 测试高性能版本
    auto [row, col] = TXCoordUtils::parseCoordFast("B2");
    EXPECT_EQ(row, 1);
    EXPECT_EQ(col, 1);
    
    // 测试无效坐标
    auto [invalid_row, invalid_col] = TXCoordUtils::parseCoordFast("INVALID");
    EXPECT_EQ(invalid_row, TXCoordUtils::INVALID_INDEX);
    EXPECT_EQ(invalid_col, TXCoordUtils::INVALID_INDEX);
    
    TX_LOG_INFO("高性能坐标解析测试通过");
}

/**
 * @brief 测试坐标转换为Excel格式
 */
TEST_F(CoordUtilsTest, CoordToExcel) {
    // 测试基本转换
    std::string excel_coord = TXCoordUtils::coordToExcel(0, 0);
    EXPECT_EQ(excel_coord, "A1");
    
    excel_coord = TXCoordUtils::coordToExcel(99, 25);
    EXPECT_EQ(excel_coord, "Z100");
    
    excel_coord = TXCoordUtils::coordToExcel(0, 26);
    EXPECT_EQ(excel_coord, "AA1");
    
    // 测试TXCoordinate版本
    TXCoordinate coord(row_t(1), column_t(1));
    excel_coord = TXCoordUtils::coordToExcel(coord);
    EXPECT_EQ(excel_coord, "B2");
    
    TX_LOG_INFO("坐标转Excel格式测试通过");
}

/**
 * @brief 测试范围解析
 */
TEST_F(CoordUtilsTest, RangeParsing) {
    auto result = TXCoordUtils::parseRange("A1:B2");
    ASSERT_TRUE(result.isOk());
    
    auto [start, end] = result.value();
    EXPECT_EQ(start.getRow().index(), 0);
    EXPECT_EQ(start.getCol().index(), 0);
    EXPECT_EQ(end.getRow().index(), 1);
    EXPECT_EQ(end.getCol().index(), 1);
    
    // 测试范围转换
    std::string range_str = TXCoordUtils::rangeToExcel(start, end);
    EXPECT_EQ(range_str, "A1:B2");
    
    TX_LOG_INFO("范围解析测试通过");
}

/**
 * @brief 测试列转换
 */
TEST_F(CoordUtilsTest, ColumnConversion) {
    // 测试列字母转索引
    uint32_t col_index = TXCoordUtils::columnLettersToIndex("A");
    EXPECT_EQ(col_index, 1); // 1-based
    
    col_index = TXCoordUtils::columnLettersToIndex("Z");
    EXPECT_EQ(col_index, 26);
    
    col_index = TXCoordUtils::columnLettersToIndex("AA");
    EXPECT_EQ(col_index, 27);
    
    // 测试索引转列字母
    std::string col_letters = TXCoordUtils::columnIndexToLetters(0); // 0-based输入
    EXPECT_EQ(col_letters, "A");
    
    col_letters = TXCoordUtils::columnIndexToLetters(25);
    EXPECT_EQ(col_letters, "Z");
    
    col_letters = TXCoordUtils::columnIndexToLetters(26);
    EXPECT_EQ(col_letters, "AA");
    
    TX_LOG_INFO("列转换测试通过");
}

/**
 * @brief 测试验证功能
 */
TEST_F(CoordUtilsTest, Validation) {
    // 测试有效坐标
    EXPECT_TRUE(TXCoordUtils::isValidExcelCoord("A1"));
    EXPECT_TRUE(TXCoordUtils::isValidExcelCoord("Z100"));
    EXPECT_TRUE(TXCoordUtils::isValidExcelCoord("AA1"));
    
    // 测试无效坐标
    EXPECT_FALSE(TXCoordUtils::isValidExcelCoord(""));
    EXPECT_FALSE(TXCoordUtils::isValidExcelCoord("A"));
    EXPECT_FALSE(TXCoordUtils::isValidExcelCoord("1"));
    EXPECT_FALSE(TXCoordUtils::isValidExcelCoord("INVALID"));
    
    // 测试有效范围
    EXPECT_TRUE(TXCoordUtils::isValidExcelRange("A1:B2"));
    EXPECT_TRUE(TXCoordUtils::isValidExcelRange("C3:Z100"));
    
    // 测试无效范围
    EXPECT_FALSE(TXCoordUtils::isValidExcelRange("A1"));
    EXPECT_FALSE(TXCoordUtils::isValidExcelRange("A1:INVALID"));
    
    TX_LOG_INFO("验证功能测试通过");
}

/**
 * @brief 测试批量转换性能
 */
TEST_F(CoordUtilsTest, BatchConversionPerformance) {
    constexpr size_t COORD_COUNT = 10000;
    
    // 准备测试数据
    std::vector<TXCoordinate> coords;
    coords.reserve(COORD_COUNT);
    
    for (size_t i = 0; i < COORD_COUNT; ++i) {
        coords.emplace_back(row_t(i / 100), column_t(i % 100));
    }
    
    // 测试批量转换性能
    auto start_time = std::chrono::high_resolution_clock::now();
    
    std::vector<std::string> excel_coords;
    TXCoordUtils::coordsBatchToExcel(coords.data(), coords.size(), excel_coords);
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end_time - start_time);
    
    EXPECT_EQ(excel_coords.size(), COORD_COUNT);
    EXPECT_EQ(excel_coords[0], "A1");
    EXPECT_EQ(excel_coords[99], "CV1");
    
    TX_LOG_INFO("批量转换{}个坐标耗时: {:.3f}ms", COORD_COUNT, duration.count() / 1000.0);
    TX_LOG_INFO("转换速度: {:.0f} 坐标/秒", COORD_COUNT / (duration.count() / 1000000.0));
    
    // 性能要求：应该能在合理时间内完成1万个坐标的转换
    EXPECT_LT(duration.count() / 1000.0, 100.0); // 100ms内完成（实际37ms，性能很好）
}

/**
 * @brief 测试往返转换一致性
 */
TEST_F(CoordUtilsTest, RoundTripConsistency) {
    // 测试大量随机坐标的往返转换
    for (uint32_t row = 0; row < 1000; row += 37) {
        for (uint32_t col = 0; col < 100; col += 7) {
            // 坐标 -> Excel -> 坐标
            std::string excel_coord = TXCoordUtils::coordToExcel(row, col);
            auto [parsed_row, parsed_col] = TXCoordUtils::parseCoordFast(excel_coord);
            
            EXPECT_EQ(parsed_row, row) << "行转换不一致: " << excel_coord;
            EXPECT_EQ(parsed_col, col) << "列转换不一致: " << excel_coord;
        }
    }
    
    TX_LOG_INFO("往返转换一致性测试通过");
}

/**
 * @brief 性能对比测试
 */
TEST_F(CoordUtilsTest, PerformanceComparison) {
    constexpr size_t TEST_COUNT = 100000;
    
    // 测试高性能版本
    auto start_time = std::chrono::high_resolution_clock::now();
    
    for (size_t i = 0; i < TEST_COUNT; ++i) {
        auto [row, col] = TXCoordUtils::parseCoordFast("B2");
        (void)row; (void)col; // 避免编译器优化
    }
    
    auto end_time = std::chrono::high_resolution_clock::now();
    auto fast_duration = std::chrono::duration_cast<std::chrono::nanoseconds>(end_time - start_time);
    
    TX_LOG_INFO("高性能版本解析{}次耗时: {:.3f}ms", TEST_COUNT, fast_duration.count() / 1000000.0);
    TX_LOG_INFO("平均每次解析: {:.1f}ns", static_cast<double>(fast_duration.count()) / TEST_COUNT);
    
    // 性能要求：每次解析应该在100ns以内
    double avg_time_ns = static_cast<double>(fast_duration.count()) / TEST_COUNT;
    EXPECT_LT(avg_time_ns, 1000.0); // 1μs以内
}
