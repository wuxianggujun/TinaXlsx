//
// TinaXlsx 性能分析器
// 用于分析和报告性能测试结果，识别性能瓶颈
//

#pragma once

#include <string>
#include <vector>
#include <map>
#include <chrono>
#include <fstream>
#include <iostream>
#include <algorithm>

namespace TinaXlsx {
namespace Performance {

/**
 * @brief 性能指标数据结构
 */
struct PerformanceMetric {
    std::string name;                    ///< 指标名称
    std::chrono::microseconds duration; ///< 执行时间
    size_t memory_used;                  ///< 内存使用量
    size_t operations_count;             ///< 操作数量
    std::string category;                ///< 分类
    
    // 计算每操作耗时
    double getTimePerOperation() const {
        if (operations_count == 0) return 0.0;
        return static_cast<double>(duration.count()) / operations_count;
    }
    
    // 计算每操作内存使用
    double getMemoryPerOperation() const {
        if (operations_count == 0) return 0.0;
        return static_cast<double>(memory_used) / operations_count;
    }
};

/**
 * @brief 性能分析器类
 */
class PerformanceAnalyzer {
public:
    /**
     * @brief 添加性能指标
     */
    void addMetric(const PerformanceMetric& metric) {
        metrics_.push_back(metric);
    }
    
    /**
     * @brief 生成性能报告
     */
    void generateReport(const std::string& filename = "performance_report.md") {
        std::ofstream report(filename);
        if (!report.is_open()) {
            std::cerr << "无法创建性能报告文件: " << filename << std::endl;
            return;
        }
        
        writeHeader(report);
        writeSummary(report);
        writeDetailedAnalysis(report);
        writeRecommendations(report);
        
        report.close();
        std::cout << "性能报告已生成: " << filename << std::endl;
    }
    
    /**
     * @brief 控制台输出简要报告
     */
    void printSummary() {
        std::cout << "\n=== TinaXlsx 性能分析摘要 ===" << std::endl;
        
        if (metrics_.empty()) {
            std::cout << "没有性能数据" << std::endl;
            return;
        }
        
        // 按类别分组
        std::map<std::string, std::vector<PerformanceMetric>> categories;
        for (const auto& metric : metrics_) {
            categories[metric.category].push_back(metric);
        }
        
        for (const auto& [category, metrics] : categories) {
            std::cout << "\n--- " << category << " ---" << std::endl;
            
            for (const auto& metric : metrics) {
                std::cout << "  " << metric.name << ": " 
                         << metric.duration.count() << "μs";
                
                if (metric.operations_count > 0) {
                    std::cout << " (" << std::fixed << std::setprecision(2) 
                             << metric.getTimePerOperation() << "μs/op)";
                }
                
                std::cout << ", 内存: " << formatMemorySize(metric.memory_used) << std::endl;
            }
        }
        
        // 识别性能问题
        identifyPerformanceIssues();
    }
    
private:
    std::vector<PerformanceMetric> metrics_;
    
    void writeHeader(std::ofstream& report) {
        auto now = std::chrono::system_clock::now();
        auto time_t = std::chrono::system_clock::to_time_t(now);
        
        report << "# TinaXlsx 性能分析报告\n\n";
        report << "**生成时间**: " << std::ctime(&time_t) << "\n";
        report << "**测试项目数**: " << metrics_.size() << "\n\n";
    }
    
    void writeSummary(std::ofstream& report) {
        report << "## 性能摘要\n\n";
        
        if (metrics_.empty()) {
            report << "没有性能数据\n\n";
            return;
        }
        
        // 计算总体统计
        auto total_time = std::chrono::microseconds(0);
        size_t total_memory = 0;
        size_t total_operations = 0;
        
        for (const auto& metric : metrics_) {
            total_time += metric.duration;
            total_memory += metric.memory_used;
            total_operations += metric.operations_count;
        }
        
        report << "| 指标 | 值 |\n";
        report << "|------|----|\n";
        report << "| 总执行时间 | " << total_time.count() << " μs |\n";
        report << "| 总内存使用 | " << formatMemorySize(total_memory) << " |\n";
        report << "| 总操作数 | " << total_operations << " |\n";
        
        if (total_operations > 0) {
            report << "| 平均每操作时间 | " << std::fixed << std::setprecision(2) 
                   << static_cast<double>(total_time.count()) / total_operations << " μs |\n";
        }
        
        report << "\n";
    }
    
    void writeDetailedAnalysis(std::ofstream& report) {
        report << "## 详细分析\n\n";
        
        // 按类别分组
        std::map<std::string, std::vector<PerformanceMetric>> categories;
        for (const auto& metric : metrics_) {
            categories[metric.category].push_back(metric);
        }
        
        for (const auto& [category, metrics] : categories) {
            report << "### " << category << "\n\n";
            
            report << "| 测试项 | 执行时间(μs) | 内存使用 | 操作数 | 每操作时间(μs) |\n";
            report << "|--------|-------------|---------|--------|---------------|\n";
            
            for (const auto& metric : metrics) {
                report << "| " << metric.name 
                       << " | " << metric.duration.count()
                       << " | " << formatMemorySize(metric.memory_used)
                       << " | " << metric.operations_count
                       << " | " << std::fixed << std::setprecision(2) << metric.getTimePerOperation()
                       << " |\n";
            }
            
            report << "\n";
        }
    }
    
    void writeRecommendations(std::ofstream& report) {
        report << "## 性能优化建议\n\n";
        
        // 找出最慢的操作
        auto slowest = std::max_element(metrics_.begin(), metrics_.end(),
            [](const PerformanceMetric& a, const PerformanceMetric& b) {
                return a.getTimePerOperation() < b.getTimePerOperation();
            });
        
        if (slowest != metrics_.end()) {
            report << "### 🔴 性能瓶颈\n\n";
            report << "**最慢操作**: " << slowest->name << "\n";
            report << "- 每操作耗时: " << std::fixed << std::setprecision(2) 
                   << slowest->getTimePerOperation() << " μs\n";
            report << "- 建议: 重点优化此操作的算法复杂度\n\n";
        }
        
        // 找出内存使用最多的操作
        auto memory_heavy = std::max_element(metrics_.begin(), metrics_.end(),
            [](const PerformanceMetric& a, const PerformanceMetric& b) {
                return a.getMemoryPerOperation() < b.getMemoryPerOperation();
            });
        
        if (memory_heavy != metrics_.end()) {
            report << "### 🟡 内存优化\n\n";
            report << "**内存使用最多**: " << memory_heavy->name << "\n";
            report << "- 每操作内存: " << std::fixed << std::setprecision(2) 
                   << memory_heavy->getMemoryPerOperation() << " bytes\n";
            report << "- 建议: 考虑内存池或对象复用策略\n\n";
        }
        
        // 通用建议
        report << "### 🟢 通用优化建议\n\n";
        report << "1. **字符串优化**: 使用字符串池减少重复字符串的内存占用\n";
        report << "2. **批量操作**: 实现批量设置单元格值的API\n";
        report << "3. **内存管理**: 考虑使用内存池管理小对象\n";
        report << "4. **IO优化**: 优化XML生成和ZIP压缩过程\n";
        report << "5. **缓存策略**: 对频繁访问的数据实现缓存\n\n";
    }
    
    void identifyPerformanceIssues() {
        std::cout << "\n=== 性能问题识别 ===" << std::endl;
        
        // 检查是否有异常慢的操作
        const double SLOW_THRESHOLD = 1000.0; // 1000μs per operation
        bool found_issues = false;
        
        for (const auto& metric : metrics_) {
            if (metric.getTimePerOperation() > SLOW_THRESHOLD) {
                if (!found_issues) {
                    std::cout << "🔴 发现性能问题:" << std::endl;
                    found_issues = true;
                }
                std::cout << "  - " << metric.name << ": " 
                         << std::fixed << std::setprecision(2) 
                         << metric.getTimePerOperation() << "μs/op (阈值: " 
                         << SLOW_THRESHOLD << "μs/op)" << std::endl;
            }
        }
        
        if (!found_issues) {
            std::cout << "✅ 未发现明显的性能问题" << std::endl;
        }
    }
    
    std::string formatMemorySize(size_t bytes) {
        const char* units[] = {"B", "KB", "MB", "GB"};
        int unit = 0;
        double size = static_cast<double>(bytes);
        
        while (size >= 1024.0 && unit < 3) {
            size /= 1024.0;
            unit++;
        }
        
        std::ostringstream oss;
        oss << std::fixed << std::setprecision(2) << size << " " << units[unit];
        return oss.str();
    }
};

} // namespace Performance
} // namespace TinaXlsx
