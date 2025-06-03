#include <gtest/gtest.h>
#include "TinaXlsx/TinaXlsx.hpp"
#include "TinaXlsx/TXCompactCell.hpp"
#include <chrono>
#include <vector>
#include <random>
#include <iostream>
#include <iomanip>

using namespace TinaXlsx;

class MemoryOptimizationTest : public ::testing::Test {
protected:
    void SetUp() override {
        gen_.seed(12345);
    }
    
    std::string generateRandomString(size_t length) {
        const std::string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        std::uniform_int_distribution<> dis(0, chars.size() - 1);
        std::string result;
        result.reserve(length);
        for (size_t i = 0; i < length; ++i) {
            result += chars[dis(gen_)];
        }
        return result;
    }
    
    template<typename Func>
    std::chrono::microseconds measureTime(Func&& func) {
        auto start = std::chrono::high_resolution_clock::now();
        func();
        auto end = std::chrono::high_resolution_clock::now();
        return std::chrono::duration_cast<std::chrono::microseconds>(end - start);
    }

protected:
    std::mt19937 gen_;
};

// 测试紧凑型单元格内存使用
TEST_F(MemoryOptimizationTest, CompactCellMemoryUsage) {
    std::cout << "=== 紧凑型单元格内存优化测试 ===\n\n";
    
    const size_t numCells = 100000;
    
    // 测试原始TXCell内存使用
    std::vector<TXCell> originalCells;
    originalCells.reserve(numCells);
    
    auto originalTime = measureTime([&]() {
        for (size_t i = 0; i < numCells; ++i) {
            originalCells.emplace_back();
            if (i % 3 == 0) {
                originalCells.back().setValue(generateRandomString(8));
            } else if (i % 3 == 1) {
                originalCells.back().setValue(static_cast<double>(i));
            } else {
                originalCells.back().setValue(static_cast<int64_t>(i));
            }
        }
    });
    
    // 测试紧凑型TXCompactCell内存使用
    std::vector<TXCompactCell> compactCells;
    compactCells.reserve(numCells);
    
    auto compactTime = measureTime([&]() {
        for (size_t i = 0; i < numCells; ++i) {
            compactCells.emplace_back();
            if (i % 3 == 0) {
                compactCells.back().setValue(generateRandomString(8));
            } else if (i % 3 == 1) {
                compactCells.back().setValue(static_cast<double>(i));
            } else {
                compactCells.back().setValue(static_cast<int64_t>(i));
            }
        }
    });
    
    // 计算内存使用
    size_t originalMemory = originalCells.size() * sizeof(TXCell);
    size_t compactMemory = 0;
    for (const auto& cell : compactCells) {
        compactMemory += cell.getMemoryUsage();
    }
    
    std::cout << "内存使用对比 (" << numCells << " 单元格):\n";
    std::cout << "  原始TXCell: " << std::fixed << std::setprecision(2) 
              << (originalMemory / 1024.0 / 1024.0) << " MB\n";
    std::cout << "  紧凑TXCompactCell: " << std::fixed << std::setprecision(2) 
              << (compactMemory / 1024.0 / 1024.0) << " MB\n";
    
    double memorySaving = 1.0 - static_cast<double>(compactMemory) / originalMemory;
    std::cout << "  内存节省: " << std::fixed << std::setprecision(1) 
              << (memorySaving * 100) << "%\n";
    
    std::cout << "\n创建时间对比:\n";
    std::cout << "  原始TXCell: " << originalTime.count() << "μs\n";
    std::cout << "  紧凑TXCompactCell: " << compactTime.count() << "μs\n";
    
    double timeRatio = static_cast<double>(compactTime.count()) / originalTime.count();
    std::cout << "  时间比率: " << std::fixed << std::setprecision(2) << timeRatio << "x\n\n";
    
    // 验证数据正确性
    EXPECT_EQ(originalCells[0].getValue(), compactCells[0].getValue());
    EXPECT_EQ(originalCells[numCells-1].getValue(), compactCells[numCells-1].getValue());
    
    std::cout << "✅ 数据正确性验证通过\n\n";
}

// 测试紧凑型单元格管理器性能
TEST_F(MemoryOptimizationTest, CompactCellManagerPerformance) {
    std::cout << "=== 紧凑型单元格管理器性能测试 ===\n\n";
    
    const size_t rows = 500;
    const size_t cols = 200;
    const size_t totalCells = rows * cols;
    
    // 生成测试数据
    std::vector<std::vector<cell_value_t>> testData;
    testData.reserve(rows);
    
    for (size_t r = 0; r < rows; ++r) {
        std::vector<cell_value_t> rowData;
        rowData.reserve(cols);
        for (size_t c = 0; c < cols; ++c) {
            if (c % 4 == 0) {
                rowData.emplace_back(generateRandomString(6));
            } else if (c % 4 == 1) {
                rowData.emplace_back(static_cast<double>(r * cols + c));
            } else if (c % 4 == 2) {
                rowData.emplace_back(static_cast<int64_t>(r * cols + c));
            } else {
                rowData.emplace_back(r % 2 == 0);
            }
        }
        testData.push_back(std::move(rowData));
    }
    
    // 测试原始CellManager
    TXCellManager originalManager;
    auto originalTime = measureTime([&]() {
        originalManager.setRangeValues(row_t(1), column_t(1), testData);
    });
    
    // 测试紧凑型CellManager
    TXCompactCellManager compactManager;
    auto compactTime = measureTime([&]() {
        compactManager.setRangeValues(row_t(1), column_t(1), testData);
    });
    
    std::cout << "批量操作性能对比 (" << totalCells << " 单元格):\n";
    std::cout << "  原始CellManager: " << originalTime.count() << "μs\n";
    std::cout << "  紧凑CellManager: " << compactTime.count() << "μs\n";
    
    double speedup = static_cast<double>(originalTime.count()) / compactTime.count();
    std::cout << "  性能提升: " << std::fixed << std::setprecision(2) << speedup << "x\n\n";
    
    // 获取内存统计
    auto memStats = compactManager.getMemoryStats();
    std::cout << "紧凑型管理器内存统计:\n";
    std::cout << "  总单元格: " << memStats.total_cells << "\n";
    std::cout << "  内存使用: " << std::fixed << std::setprecision(2) 
              << (memStats.memory_used / 1024.0 / 1024.0) << " MB\n";
    std::cout << "  内存节省: " << std::fixed << std::setprecision(2) 
              << (memStats.memory_saved / 1024.0 / 1024.0) << " MB\n";
    std::cout << "  压缩比率: " << std::fixed << std::setprecision(1) 
              << (memStats.compact_ratio * 100) << "%\n\n";
    
    std::cout << "✅ 紧凑型管理器测试完成\n\n";
}

// 测试内存压缩功能
TEST_F(MemoryOptimizationTest, MemoryCompactionTest) {
    std::cout << "=== 内存压缩功能测试 ===\n\n";
    
    TXCompactCellManager manager;
    const size_t numCells = 10000;
    
    // 创建大量单元格，其中一些有扩展数据
    std::cout << "创建 " << numCells << " 个单元格...\n";
    
    auto createTime = measureTime([&]() {
        for (size_t i = 0; i < numCells; ++i) {
            TXCoordinate coord(row_t(i / 100 + 1), column_t(i % 100 + 1));
            
            if (i % 10 == 0) {
                // 创建有样式的单元格
                auto* cell = manager.getOrCreateCell(coord);
                cell->setValue(generateRandomString(5));
                cell->setStyleIndex(i % 5 + 1);
            } else {
                // 创建普通单元格
                manager.setCellValue(coord, static_cast<double>(i));
            }
        }
    });
    
    auto beforeStats = manager.getMemoryStats();
    std::cout << "压缩前内存统计:\n";
    std::cout << "  内存使用: " << std::fixed << std::setprecision(2) 
              << (beforeStats.memory_used / 1024.0) << " KB\n";
    
    // 执行内存压缩
    auto compactTime = measureTime([&]() {
        manager.compactMemory();
    });
    
    auto afterStats = manager.getMemoryStats();
    std::cout << "压缩后内存统计:\n";
    std::cout << "  内存使用: " << std::fixed << std::setprecision(2) 
              << (afterStats.memory_used / 1024.0) << " KB\n";
    
    size_t memoryFreed = beforeStats.memory_used - afterStats.memory_used;
    double compressionRatio = static_cast<double>(memoryFreed) / beforeStats.memory_used;
    
    std::cout << "压缩效果:\n";
    std::cout << "  释放内存: " << std::fixed << std::setprecision(2) 
              << (memoryFreed / 1024.0) << " KB\n";
    std::cout << "  压缩比率: " << std::fixed << std::setprecision(1) 
              << (compressionRatio * 100) << "%\n";
    std::cout << "  压缩时间: " << compactTime.count() << "μs\n\n";
    
    std::cout << "✅ 内存压缩测试完成\n\n";
}

// 测试大数据量场景
TEST_F(MemoryOptimizationTest, LargeDataScenario) {
    std::cout << "=== 大数据量场景测试 ===\n\n";
    
    const size_t rows = 1000;
    const size_t cols = 1000;
    const size_t totalCells = rows * cols;
    
    std::cout << "测试场景: " << rows << "x" << cols << " = " << totalCells << " 单元格\n\n";
    
    TXCompactCellManager manager;
    manager.reserve(totalCells); // 预分配内存
    
    // 分批添加数据以模拟实际使用场景
    const size_t batchSize = 10000;
    const size_t numBatches = (totalCells + batchSize - 1) / batchSize;
    
    std::vector<std::chrono::microseconds> batchTimes;
    batchTimes.reserve(numBatches);
    
    std::cout << "分 " << numBatches << " 批添加数据:\n";
    
    for (size_t batch = 0; batch < numBatches; ++batch) {
        size_t startIdx = batch * batchSize;
        size_t endIdx = std::min(startIdx + batchSize, totalCells);
        size_t currentBatchSize = endIdx - startIdx;
        
        auto batchTime = measureTime([&]() {
            for (size_t i = startIdx; i < endIdx; ++i) {
                size_t r = i / cols;
                size_t c = i % cols;
                TXCoordinate coord(row_t(r + 1), column_t(c + 1));
                
                if (i % 5 == 0) {
                    manager.setCellValue(coord, generateRandomString(4));
                } else if (i % 5 == 1) {
                    manager.setCellValue(coord, static_cast<double>(i));
                } else {
                    manager.setCellValue(coord, static_cast<int64_t>(i));
                }
            }
        });
        
        batchTimes.push_back(batchTime);
        
        double timePerCell = static_cast<double>(batchTime.count()) / currentBatchSize;
        std::cout << "批次 " << (batch + 1) << "/" << numBatches 
                  << ": " << batchTime.count() << "μs"
                  << ", 平均: " << std::fixed << std::setprecision(2) << timePerCell << "μs/cell\n";
    }
    
    // 计算性能统计
    auto totalTime = std::accumulate(batchTimes.begin(), batchTimes.end(), 
                                   std::chrono::microseconds(0));
    auto minTime = *std::min_element(batchTimes.begin(), batchTimes.end());
    auto maxTime = *std::max_element(batchTimes.begin(), batchTimes.end());
    auto avgTime = totalTime / numBatches;
    
    std::cout << "\n性能统计:\n";
    std::cout << "  总时间: " << std::fixed << std::setprecision(2) 
              << (totalTime.count() / 1000.0) << "ms\n";
    std::cout << "  平均批次时间: " << avgTime.count() << "μs\n";
    std::cout << "  最快批次: " << minTime.count() << "μs\n";
    std::cout << "  最慢批次: " << maxTime.count() << "μs\n";
    
    double avgTimePerCell = static_cast<double>(totalTime.count()) / totalCells;
    std::cout << "  平均单元格时间: " << std::fixed << std::setprecision(2) 
              << avgTimePerCell << "μs/cell\n";
    
    // 内存统计
    auto memStats = manager.getMemoryStats();
    std::cout << "\n最终内存统计:\n";
    std::cout << "  总单元格: " << memStats.total_cells << "\n";
    std::cout << "  内存使用: " << std::fixed << std::setprecision(2) 
              << (memStats.memory_used / 1024.0 / 1024.0) << " MB\n";
    std::cout << "  每单元格: " << std::fixed << std::setprecision(1) 
              << (static_cast<double>(memStats.memory_used) / totalCells) << " bytes\n";
    
    std::cout << "\n✅ 大数据量测试完成\n\n";
}
