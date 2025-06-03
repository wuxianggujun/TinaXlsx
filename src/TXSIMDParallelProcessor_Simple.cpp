//
// @file TXSIMDParallelProcessor_Simple.cpp
// @brief SIMD + 并行处理器简化实现
//

#include "TinaXlsx/TXSIMDParallelProcessor.hpp"
#include <algorithm>
#include <numeric>
#include <cmath>
#include <iostream>
#include <cstring>

namespace TinaXlsx {

// ==================== TXSIMDParallelProcessor 简化实现 ====================

TXSIMDParallelProcessor::TXSIMDParallelProcessor(const SIMDParallelConfig& config)
    : config_(config) {
    
    // 自动检测线程数
    if (config_.thread_count == 0) {
        config_.thread_count = std::thread::hardware_concurrency();
        if (config_.thread_count == 0) {
            config_.thread_count = 4; // 回退值
        }
    }
    
    std::cout << "TXSIMDParallelProcessor initialized with " << config_.thread_count 
              << " threads, SIMD: " << (config_.enable_simd ? "enabled" : "disabled")
              << " (" << SIMDCapabilities::getSIMDInfo() << ")" << std::endl;
}

TXSIMDParallelProcessor::~TXSIMDParallelProcessor() {
    // 简化版本，无需清理
}

// ==================== 高性能批处理API实现 ====================

void TXSIMDParallelProcessor::ultraFastConvertDoublesToCells(const std::vector<double>& input,
                                                            std::vector<UltraCompactCell>& output) {
    if (input.empty()) return;
    
    auto start = std::chrono::steady_clock::now();
    
    output.resize(input.size());
    
    if (!config_.enable_parallel || input.size() < config_.min_batch_size) {
        // 单线程SIMD处理
        if (config_.enable_simd) {
            TXSIMDProcessor::convertDoublesToCells(input.data(), output.data(), input.size());
        } else {
            for (size_t i = 0; i < input.size(); ++i) {
                output[i] = UltraCompactCell(input[i]);
            }
        }
    } else {
        // 并行处理 - 简化版本，使用线程直接执行
        auto splits = calculateOptimalSplits(input.size());
        std::vector<std::thread> threads;
        threads.reserve(splits.size());
        
        for (const auto& split : splits) {
            size_t start_idx = split.first;
            size_t end_idx = split.second;
            size_t count = end_idx - start_idx;
            
            threads.emplace_back([this, &input, &output, start_idx, count]() {
                if (config_.enable_simd) {
                    TXSIMDProcessor::convertDoublesToCells(
                        input.data() + start_idx,
                        output.data() + start_idx,
                        count
                    );
                } else {
                    for (size_t i = 0; i < count; ++i) {
                        output[start_idx + i] = UltraCompactCell(input[start_idx + i]);
                    }
                }
            });
        }
        
        // 等待所有线程完成
        for (auto& thread : threads) {
            thread.join();
        }
    }
    
    auto end = std::chrono::steady_clock::now();
    updateMetrics("convertDoublesToCells", start, end, input.size());
}

void TXSIMDParallelProcessor::ultraFastConvertInt64sToCells(const std::vector<int64_t>& input,
                                                           std::vector<UltraCompactCell>& output) {
    if (input.empty()) return;
    
    auto start = std::chrono::steady_clock::now();
    
    output.resize(input.size());
    
    if (!config_.enable_parallel || input.size() < config_.min_batch_size) {
        // 单线程SIMD处理
        if (config_.enable_simd) {
            TXSIMDProcessor::convertInt64sToCells(input.data(), output.data(), input.size());
        } else {
            for (size_t i = 0; i < input.size(); ++i) {
                output[i] = UltraCompactCell(input[i]);
            }
        }
    } else {
        // 并行处理 - 简化版本，使用线程直接执行
        auto splits = calculateOptimalSplits(input.size());
        std::vector<std::thread> threads;
        threads.reserve(splits.size());
        
        for (const auto& split : splits) {
            size_t start_idx = split.first;
            size_t end_idx = split.second;
            size_t count = end_idx - start_idx;
            
            threads.emplace_back([this, &input, &output, start_idx, count]() {
                if (config_.enable_simd) {
                    TXSIMDProcessor::convertInt64sToCells(
                        input.data() + start_idx,
                        output.data() + start_idx,
                        count
                    );
                } else {
                    for (size_t i = 0; i < count; ++i) {
                        output[start_idx + i] = UltraCompactCell(input[start_idx + i]);
                    }
                }
            });
        }
        
        // 等待所有线程完成
        for (auto& thread : threads) {
            thread.join();
        }
    }
    
    auto end = std::chrono::steady_clock::now();
    updateMetrics("convertInt64sToCells", start, end, input.size());
}

void TXSIMDParallelProcessor::ultraFastSetCoordinates(std::vector<UltraCompactCell>& cells,
                                                     const std::vector<uint16_t>& rows,
                                                     const std::vector<uint16_t>& cols) {
    if (cells.empty() || rows.size() != cells.size() || cols.size() != cells.size()) {
        return;
    }
    
    auto start = std::chrono::steady_clock::now();
    
    if (!config_.enable_parallel || cells.size() < config_.min_batch_size) {
        // 单线程SIMD处理
        if (config_.enable_simd) {
            TXSIMDProcessor::setCoordinates(cells.data(), rows.data(), cols.data(), cells.size());
        } else {
            for (size_t i = 0; i < cells.size(); ++i) {
                cells[i].setRow(rows[i]);
                cells[i].setCol(cols[i]);
            }
        }
    } else {
        // 并行处理
        auto splits = calculateOptimalSplits(cells.size());
        std::vector<std::thread> threads;
        threads.reserve(splits.size());
        
        for (const auto& split : splits) {
            size_t start_idx = split.first;
            size_t end_idx = split.second;
            size_t count = end_idx - start_idx;
            
            threads.emplace_back([this, &cells, &rows, &cols, start_idx, count]() {
                if (config_.enable_simd) {
                    TXSIMDProcessor::setCoordinates(
                        cells.data() + start_idx,
                        rows.data() + start_idx,
                        cols.data() + start_idx,
                        count
                    );
                } else {
                    for (size_t i = 0; i < count; ++i) {
                        cells[start_idx + i].setRow(rows[start_idx + i]);
                        cells[start_idx + i].setCol(cols[start_idx + i]);
                    }
                }
            });
        }
        
        // 等待所有线程完成
        for (auto& thread : threads) {
            thread.join();
        }
    }
    
    auto end = std::chrono::steady_clock::now();
    updateMetrics("setCoordinates", start, end, cells.size());
}

void TXSIMDParallelProcessor::ultraFastClearCells(std::vector<UltraCompactCell>& cells) {
    if (cells.empty()) return;
    
    auto start = std::chrono::steady_clock::now();
    
    if (!config_.enable_parallel || cells.size() < config_.min_batch_size) {
        // 单线程SIMD处理
        if (config_.enable_simd) {
            TXSIMDProcessor::clearCells(cells.data(), cells.size());
        } else {
            for (auto& cell : cells) {
                cell.clear();
            }
        }
    } else {
        // 并行处理
        auto splits = calculateOptimalSplits(cells.size());
        std::vector<std::thread> threads;
        threads.reserve(splits.size());
        
        for (const auto& split : splits) {
            size_t start_idx = split.first;
            size_t end_idx = split.second;
            size_t count = end_idx - start_idx;
            
            threads.emplace_back([this, &cells, start_idx, count]() {
                if (config_.enable_simd) {
                    TXSIMDProcessor::clearCells(cells.data() + start_idx, count);
                } else {
                    for (size_t i = 0; i < count; ++i) {
                        cells[start_idx + i].clear();
                    }
                }
            });
        }
        
        // 等待所有线程完成
        for (auto& thread : threads) {
            thread.join();
        }
    }
    
    auto end = std::chrono::steady_clock::now();
    updateMetrics("clearCells", start, end, cells.size());
}

void TXSIMDParallelProcessor::ultraFastCopyCells(const std::vector<UltraCompactCell>& src,
                                                std::vector<UltraCompactCell>& dst) {
    if (src.empty()) return;
    
    auto start = std::chrono::steady_clock::now();
    
    dst.resize(src.size());
    
    if (!config_.enable_parallel || src.size() < config_.min_batch_size) {
        // 单线程SIMD处理
        if (config_.enable_simd) {
            TXSIMDProcessor::copyCells(src.data(), dst.data(), src.size());
        } else {
            std::copy(src.begin(), src.end(), dst.begin());
        }
    } else {
        // 并行处理
        auto splits = calculateOptimalSplits(src.size());
        std::vector<std::thread> threads;
        threads.reserve(splits.size());
        
        for (const auto& split : splits) {
            size_t start_idx = split.first;
            size_t end_idx = split.second;
            size_t count = end_idx - start_idx;
            
            threads.emplace_back([this, &src, &dst, start_idx, count]() {
                if (config_.enable_simd) {
                    TXSIMDProcessor::copyCells(
                        src.data() + start_idx,
                        dst.data() + start_idx,
                        count
                    );
                } else {
                    std::copy(src.begin() + start_idx, 
                             src.begin() + start_idx + count,
                             dst.begin() + start_idx);
                }
            });
        }
        
        // 等待所有线程完成
        for (auto& thread : threads) {
            thread.join();
        }
    }
    
    auto end = std::chrono::steady_clock::now();
    updateMetrics("copyCells", start, end, src.size());
}

} // namespace TinaXlsx
