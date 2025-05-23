#include <gtest/gtest.h>
#include <gmock/gmock.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include <chrono>
#include <random>
#include <sstream>
#include <filesystem>
#include <memory>

namespace TinaXlsx {
namespace Test {

/**
 * @brief 性能测试基类
 */
class PerformanceTest : public ::testing::Test {
protected:
    void SetUp() override {
        test_dir = "performance_test_files";
        // 创建测试目录
        if (!std::filesystem::exists(test_dir)) {
            std::filesystem::create_directory(test_dir);
        }
        
        // 初始化随机数生成器
        rng.seed(std::chrono::steady_clock::now().time_since_epoch().count());
    }

    void TearDown() override {
        // 清理测试文件
        if (std::filesystem::exists(test_dir)) {
            std::filesystem::remove_all(test_dir);
        }
    }

    // 生成随机字符串
    std::string generateRandomString(size_t length) {
        const std::string charset = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        std::string result;
        result.reserve(length);
        
        for (size_t i = 0; i < length; ++i) {
            result += charset[rng() % charset.size()];
        }
        
        return result;
    }

    // 生成随机数字
    double generateRandomNumber() {
        return std::uniform_real_distribution<double>(-1000000.0, 1000000.0)(rng);
    }

    // 生成随机整数
    int64_t generateRandomInteger() {
        return std::uniform_int_distribution<int64_t>(-1000000, 1000000)(rng);
    }

    // 计时辅助函数
    template<typename Func>
    std::chrono::milliseconds measureTime(Func&& func) {
        auto start = std::chrono::high_resolution_clock::now();
        func();
        auto end = std::chrono::high_resolution_clock::now();
        return std::chrono::duration_cast<std::chrono::milliseconds>(end - start);
    }

    std::string test_dir;
    std::mt19937 rng;
};

/**
 * @brief 测试小数据集写入性能
 */
TEST_F(PerformanceTest, SmallDataWritePerformance) {
    const std::string filename = test_dir + "/small_write_test.xlsx";
    const RowIndex rows = 1000;
    const ColumnIndex cols = 10;
    
    auto elapsed = measureTime([&]() {
        Workbook workbook(filename, Workbook::Mode::Write);
        auto& writer = workbook.getWriter();
        auto worksheet = writer.createWorksheet();
        
        for (RowIndex row = 0; row < rows; ++row) {
            for (ColumnIndex col = 0; col < cols; ++col) {
                CellPosition pos(row, col);
                
                // 交替写入不同类型的数据
                switch (col % 4) {
                    case 0:
                        worksheet->writeCell(pos, CellValue(generateRandomString(10)));
                        break;
                    case 1:
                        worksheet->writeCell(pos, CellValue(generateRandomNumber()));
                        break;
                    case 2:
                        worksheet->writeCell(pos, CellValue(generateRandomInteger()));
                        break;
                    case 3:
                        worksheet->writeCell(pos, CellValue((row + col) % 2 == 0));
                        break;
                }
            }
        }
        
        workbook.close();
    });
    
    std::cout << "Small data write (" << rows << "x" << cols << " cells): " 
              << elapsed.count() << "ms" << std::endl;
    
    // 性能断言：1万个单元格应该在5秒内完成
    EXPECT_LT(elapsed.count(), 5000) << "Small data write took too long";
    
    // 验证文件已创建
    EXPECT_TRUE(std::filesystem::exists(filename));
}

/**
 * @brief 测试中等数据集写入性能
 */
TEST_F(PerformanceTest, MediumDataWritePerformance) {
    const std::string filename = test_dir + "/medium_write_test.xlsx";
    const RowIndex rows = 10000;
    const ColumnIndex cols = 20;
    
    auto elapsed = measureTime([&]() {
        Workbook workbook(filename, Workbook::Mode::Write);
        auto& writer = workbook.getWriter();
        auto worksheet = writer.createWorksheet();
        
        // 批量写入以提高性能
        RowData rowData(cols);
        for (RowIndex row = 0; row < rows; ++row) {
            for (ColumnIndex col = 0; col < cols; ++col) {
                // 生成测试数据
                switch (col % 3) {
                    case 0:
                        rowData[col] = CellValue(generateRandomString(15));
                        break;
                    case 1:
                        rowData[col] = CellValue(generateRandomNumber());
                        break;
                    case 2:
                        rowData[col] = CellValue(generateRandomInteger());
                        break;
                }
            }
            
            // 写入整行数据
            for (ColumnIndex col = 0; col < cols; ++col) {
                CellPosition pos(row, col);
                worksheet->writeCell(pos, rowData[col]);
            }
        }
        
        workbook.close();
    });
    
    std::cout << "Medium data write (" << rows << "x" << cols << " cells): " 
              << elapsed.count() << "ms" << std::endl;
    
    // 性能断言：20万个单元格应该在30秒内完成
    EXPECT_LT(elapsed.count(), 30000) << "Medium data write took too long";
    
    // 验证文件已创建且大小合理
    EXPECT_TRUE(std::filesystem::exists(filename));
    auto fileSize = std::filesystem::file_size(filename);
    EXPECT_GT(fileSize, 1000) << "Generated file seems too small";
}

/**
 * @brief 测试读取性能
 */
TEST_F(PerformanceTest, ReadPerformance) {
    const std::string filename = test_dir + "/read_test.xlsx";
    const RowIndex rows = 5000;
    const ColumnIndex cols = 15;
    
    // 先创建测试文件
    {
        Workbook workbook(filename, Workbook::Mode::Write);
        auto& writer = workbook.getWriter();
        auto worksheet = writer.createWorksheet();
        
        for (RowIndex row = 0; row < rows; ++row) {
            for (ColumnIndex col = 0; col < cols; ++col) {
                CellPosition pos(row, col);
                worksheet->writeCell(pos, CellValue(std::string("Cell_") + std::to_string(row) + "_" + std::to_string(col)));
            }
        }
        workbook.close();
    }
    
    // 测试读取性能
    auto elapsed = measureTime([&]() {
        Workbook workbook(filename, Workbook::Mode::Read);
        auto& reader = workbook.getReader();
        
        // 获取工作表名称
        auto sheetNames = reader.getSheetNames();
        ASSERT_FALSE(sheetNames.empty());
        
        // 打开第一个工作表
        ASSERT_TRUE(reader.openSheet(sheetNames[0]));
        
        // 读取所有数据
        auto tableData = reader.readAll(rows, cols);
        
        // 验证读取的数据量
        EXPECT_EQ(tableData.size(), rows);
        if (!tableData.empty()) {
            EXPECT_EQ(tableData[0].size(), cols);
        }
    });
    
    std::cout << "Read performance (" << rows << "x" << cols << " cells): " 
              << elapsed.count() << "ms" << std::endl;
    
    // 性能断言：读取应该比写入快
    EXPECT_LT(elapsed.count(), 15000) << "Read took too long";
}

/**
 * @brief 测试批量数据处理性能
 */
TEST_F(PerformanceTest, BatchDataProcessing) {
    const std::string filename = test_dir + "/batch_test.xlsx";
    const RowIndex rows = 8000;
    const ColumnIndex cols = 12;
    
    auto elapsed = measureTime([&]() {
        Workbook workbook(filename, Workbook::Mode::Write);
        auto& writer = workbook.getWriter();
        auto worksheet = writer.createWorksheet();
        
        // 准备批量数据
        TableData batchData(rows, RowData(cols));
        
        // 生成批量数据
        for (RowIndex row = 0; row < rows; ++row) {
            for (ColumnIndex col = 0; col < cols; ++col) {
                if (col % 4 == 0) {
                    batchData[row][col] = CellValue(std::string("Batch_") + std::to_string(row * cols + col));
                } else if (col % 4 == 1) {
                    batchData[row][col] = CellValue(static_cast<double>(row * col) * 3.14159);
                } else if (col % 4 == 2) {
                    batchData[row][col] = CellValue(static_cast<int64_t>(row * 1000 + col));
                } else {
                    batchData[row][col] = CellValue((row + col) % 2 == 0);
                }
            }
        }
        
        // 批量写入数据
        for (RowIndex row = 0; row < rows; ++row) {
            for (ColumnIndex col = 0; col < cols; ++col) {
                CellPosition pos(row, col);
                worksheet->writeCell(pos, batchData[row][col]);
            }
        }
        
        workbook.close();
    });
    
    std::cout << "Batch processing (" << rows << "x" << cols << " cells): " 
              << elapsed.count() << "ms" << std::endl;
    
    // 性能断言
    EXPECT_LT(elapsed.count(), 25000) << "Batch processing took too long";
    
    // 验证文件创建
    EXPECT_TRUE(std::filesystem::exists(filename));
}

/**
 * @brief 测试内存使用和资源管理
 */
TEST_F(PerformanceTest, MemoryUsage) {
    const std::string filename = test_dir + "/memory_test.xlsx";
    const RowIndex rows = 3000;
    const ColumnIndex cols = 8;
    
    // 测试多次创建和销毁，检查内存泄漏
    for (int iteration = 0; iteration < 10; ++iteration) {
        std::string iterFilename = test_dir + "/memory_test_" + std::to_string(iteration) + ".xlsx";
        
        {
            Workbook workbook(iterFilename, Workbook::Mode::Write);
            auto& writer = workbook.getWriter();
            auto worksheet = writer.createWorksheet();
            
            for (RowIndex row = 0; row < rows; ++row) {
                for (ColumnIndex col = 0; col < cols; ++col) {
                    CellPosition pos(row, col);
                    worksheet->writeCell(pos, CellValue(std::string("Iter") + std::to_string(iteration) + 
                                           "_R" + std::to_string(row) + 
                                           "_C" + std::to_string(col)));
                }
            }
            
            // workbook 在作用域结束时自动销毁
        }
        
        // 验证文件创建
        EXPECT_TRUE(std::filesystem::exists(iterFilename));
    }
    
    std::cout << "Memory test completed: 10 iterations with " 
              << rows << "x" << cols << " cells each" << std::endl;
}

/**
 * @brief 测试字符串处理性能
 */
TEST_F(PerformanceTest, StringProcessingPerformance) {
    const std::string filename = test_dir + "/string_test.xlsx";
    const RowIndex rows = 2000;
    const ColumnIndex cols = 5;
    
    auto elapsed = measureTime([&]() {
        Workbook workbook(filename, Workbook::Mode::Write);
        auto& writer = workbook.getWriter();
        auto worksheet = writer.createWorksheet();
        
        for (RowIndex row = 0; row < rows; ++row) {
            for (ColumnIndex col = 0; col < cols; ++col) {
                CellPosition pos(row, col);
                
                // 生成不同长度的字符串
                size_t stringLength = 10 + (row + col) % 50;  // 10-60字符
                std::string longString = generateRandomString(stringLength);
                
                worksheet->writeCell(pos, CellValue(longString));
            }
        }
        
        workbook.close();
    });
    
    std::cout << "String processing (" << rows << "x" << cols << " varied length strings): " 
              << elapsed.count() << "ms" << std::endl;
    
    // 性能断言
    EXPECT_LT(elapsed.count(), 15000) << "String processing took too long";
}

/**
 * @brief 测试数值计算性能
 */
TEST_F(PerformanceTest, NumericProcessingPerformance) {
    const std::string filename = test_dir + "/numeric_test.xlsx";
    const RowIndex rows = 4000;
    const ColumnIndex cols = 6;
    
    auto elapsed = measureTime([&]() {
        Workbook workbook(filename, Workbook::Mode::Write);
        auto& writer = workbook.getWriter();
        auto worksheet = writer.createWorksheet();
        
        for (RowIndex row = 0; row < rows; ++row) {
            for (ColumnIndex col = 0; col < cols; ++col) {
                CellPosition pos(row, col);
                
                switch (col % 3) {
                    case 0:
                        // 写入浮点数
                        worksheet->writeCell(pos, CellValue(static_cast<double>(row * col) * 3.14159265359));
                        break;
                    case 1:
                        // 写入大整数
                        worksheet->writeCell(pos, CellValue(static_cast<int64_t>(row) * 1000000 + col));
                        break;
                    case 2:
                        // 写入计算结果
                        worksheet->writeCell(pos, CellValue(std::sqrt(static_cast<double>(row * row + col * col))));
                        break;
                }
            }
        }
        
        workbook.close();
    });
    
    std::cout << "Numeric processing (" << rows << "x" << cols << " numeric cells): " 
              << elapsed.count() << "ms" << std::endl;
    
    // 性能断言
    EXPECT_LT(elapsed.count(), 20000) << "Numeric processing took too long";
}

} // namespace Test
} // namespace TinaXlsx 