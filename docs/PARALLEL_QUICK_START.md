# TinaXlsx 并行框架快速开始指南

## 🚀 5分钟快速上手

### 1. 编译项目

```bash
# 进入项目目录
cd TinaXlsx

# 创建构建目录
mkdir -p cmake-build-debug
cd cmake-build-debug

# 配置CMake
cmake ..

# 编译项目
cmake --build .
```

### 2. 运行示例

```bash
# 运行高级并行框架示例
cmake --build . --target run_advanced_parallel_example

# 或者直接运行
./examples/AdvancedParallelExample
```

### 3. 运行性能测试

```bash
# 编译性能测试
cmake --build . --target AdvancedParallelPerformanceTests

# 运行测试
./tests/performance/AdvancedParallelPerformanceTests
```

## 💡 基础使用示例

### 示例1：智能并行单元格处理

```cpp
#include "TinaXlsx.hpp"
#include "TXAdvancedParallelFramework.hpp"

int main() {
    // 创建工作簿和工作表
    TXWorkbook workbook;
    auto sheet = workbook.addSheet("ParallelDemo");
    
    // 准备大量数据
    std::vector<std::pair<TXCoordinate, cell_value_t>> cellData;
    for (int row = 1; row <= 10000; ++row) {
        for (int col = 1; col <= 10; ++col) {
            TXCoordinate coord(row_t(row), column_t(col));
            cellData.emplace_back(coord, row * col * 1.5);
        }
    }
    
    // 创建智能并行处理器
    TXSmartParallelCellProcessor processor;
    
    // 并行处理（自动优化）
    auto result = processor.parallelSetCellValues(*sheet, cellData);
    
    if (result.isOk()) {
        std::cout << "成功处理 " << result.value() << " 个单元格" << std::endl;
        
        // 保存文件
        workbook.saveToFile("parallel_demo.xlsx");
        std::cout << "文件保存成功！" << std::endl;
    }
    
    return 0;
}
```

### 示例2：自定义配置

```cpp
// 高性能配置
TXSmartParallelCellProcessor::ProcessorConfig config;
config.numThreads = std::thread::hardware_concurrency();
config.enableAdaptiveBatching = true;
config.enableMemoryPool = true;
config.enableCacheOptimization = true;
config.minBatchSize = 1000;
config.maxBatchSize = 50000;

TXSmartParallelCellProcessor processor(config);

// 使用配置好的处理器
auto result = processor.parallelSetCellValues(*sheet, cellData);
```

### 示例3：无锁线程池直接使用

```cpp
// 创建高性能线程池
TXLockFreeThreadPool::PoolConfig poolConfig;
poolConfig.numThreads = 8;
poolConfig.enableWorkStealing = true;
poolConfig.enableMemoryPool = true;

TXLockFreeThreadPool threadPool(poolConfig);

// 提交任务
std::vector<std::future<double>> futures;
for (int i = 0; i < 1000; ++i) {
    auto future = threadPool.submit([i]() -> double {
        // 计算密集型任务
        double result = 0.0;
        for (int j = 0; j < 10000; ++j) {
            result += std::sin(i * j * 0.001);
        }
        return result;
    }, TXLockFreeThreadPool::TaskPriority::Normal);
    
    futures.push_back(std::move(future));
}

// 收集结果
double total = 0.0;
for (auto& future : futures) {
    total += future.get();
}

std::cout << "计算结果: " << total << std::endl;
```

## 📊 性能对比

### 运行性能测试查看提升效果

```bash
# 运行所有性能测试
cmake --build . --target run_performance_tests

# 查看并行框架专门测试
./tests/performance/AdvancedParallelPerformanceTests --gtest_filter="*Performance*"
```

### 预期性能提升

| 场景 | 数据量 | 预期提升 |
|------|--------|----------|
| 单元格写入 | 10万单元格 | 60%+ |
| 多工作表 | 10个工作表 | 300%+ |
| 大文件处理 | 100万单元格 | 2-5倍 |

## 🔧 常见配置

### 针对不同场景的推荐配置

#### 1. 大量小数据处理

```cpp
TXSmartParallelCellProcessor::ProcessorConfig config;
config.numThreads = std::thread::hardware_concurrency();
config.minBatchSize = 100;
config.maxBatchSize = 5000;
config.enableCacheOptimization = true;
```

#### 2. 少量大数据处理

```cpp
TXSmartParallelCellProcessor::ProcessorConfig config;
config.numThreads = 4; // 减少线程数
config.minBatchSize = 5000;
config.maxBatchSize = 50000;
config.enableMemoryPool = true;
```

#### 3. 内存受限环境

```cpp
TXSmartParallelCellProcessor::ProcessorConfig config;
config.numThreads = 2;
config.enableMemoryPool = true;
config.enableAdaptiveBatching = false; // 固定批量大小
config.minBatchSize = 1000;
config.maxBatchSize = 1000;
```

## 🐛 故障排除

### 常见问题

#### 1. 编译错误

```bash
# 确保C++17支持
cmake .. -DCMAKE_CXX_STANDARD=17

# 检查依赖
cmake .. -DBUILD_TESTS=ON
```

#### 2. 性能不如预期

```cpp
// 检查线程数配置
std::cout << "CPU核心数: " << std::thread::hardware_concurrency() << std::endl;

// 启用性能监控
auto stats = processor.getStats(); // 如果有的话
```

#### 3. 内存使用过高

```cpp
// 启用内存监控
TX_MEMORY_OPERATION_START("并行处理");
// ... 执行操作
TX_MEMORY_OPERATION_END("并行处理");

// 检查内存泄漏
TX_MEMORY_LEAK_DETECTION_REPORT();
```

## 📚 进阶学习

### 1. 深入了解架构

- 阅读 [高级并行框架文档](ADVANCED_PARALLEL_FRAMEWORK.md)
- 查看 [实现总结](PARALLEL_FRAMEWORK_IMPLEMENTATION.md)

### 2. 性能调优

- 阅读 [性能优化指南](PERFORMANCE_OPTIMIZATION.md)
- 运行性能测试分析瓶颈

### 3. 源码学习

```cpp
// 核心文件
include/TinaXlsx/TXAdvancedParallelFramework.hpp
src/TXAdvancedParallelFramework.cpp

// 示例代码
examples/advanced_parallel_example.cpp

// 测试代码
tests/performance/test_advanced_parallel_performance.cpp
```

## 🎯 最佳实践

### 1. 数据准备

```cpp
// 预先分配内存
std::vector<std::pair<TXCoordinate, cell_value_t>> cellData;
cellData.reserve(expectedSize);

// 按行列顺序组织数据（缓存友好）
std::sort(cellData.begin(), cellData.end(), 
    [](const auto& a, const auto& b) {
        if (a.first.row() != b.first.row()) {
            return a.first.row() < b.first.row();
        }
        return a.first.column() < b.first.column();
    });
```

### 2. 错误处理

```cpp
auto result = processor.parallelSetCellValues(*sheet, cellData);
if (result.isError()) {
    std::cerr << "并行处理失败: " << result.error().getMessage() << std::endl;
    // 回退到串行处理
    for (const auto& [coord, value] : cellData) {
        sheet->setCellValue(coord.row(), coord.column(), value);
    }
}
```

### 3. 性能监控

```cpp
auto startTime = std::chrono::high_resolution_clock::now();

auto result = processor.parallelSetCellValues(*sheet, cellData);

auto endTime = std::chrono::high_resolution_clock::now();
auto duration = std::chrono::duration_cast<std::chrono::microseconds>(endTime - startTime);

if (result.isOk()) {
    size_t cellCount = result.value();
    double cellsPerSecond = cellCount * 1000000.0 / duration.count();
    std::cout << "处理速度: " << cellsPerSecond << " 单元格/秒" << std::endl;
}
```

## 🚀 开始使用

现在您已经了解了基础用法，可以开始在您的项目中使用TinaXlsx高级并行框架了！

1. 从简单的示例开始
2. 根据您的数据特点调整配置
3. 运行性能测试验证效果
4. 逐步集成到生产环境

如有问题，请参考详细文档或查看示例代码。祝您使用愉快！
