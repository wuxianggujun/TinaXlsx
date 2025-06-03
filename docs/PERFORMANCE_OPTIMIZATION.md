# TinaXlsx 性能优化指南

本文档提供了TinaXlsx库的性能优化策略和最佳实践。

## 🎯 性能测试

### 运行性能测试

```bash
# 编译性能测试
cmake --build cmake-build-debug --target ExtremePerformanceTests

# 运行性能测试
cmake --build cmake-build-debug --target run_performance_tests

# 或者直接运行
./cmake-build-debug/tests/performance/ExtremePerformanceTests.exe
```

### 性能测试覆盖范围

我们的极致性能测试包括：

1. **大量数据写入测试** - 5万行×20列数据写入
2. **大文件读取测试** - 大文件加载和随机访问
3. **多工作表测试** - 50个工作表并发创建
4. **字符串池测试** - 大量重复字符串的内存优化
5. **样式性能测试** - 多种样式的应用性能
6. **内存泄漏检测** - 100次迭代的内存使用监控
7. **并发安全测试** - 多线程环境下的稳定性
8. **极限单元格测试** - 10万行×50列的极限测试
9. **性能回归测试** - 基准性能指标监控

## 🔍 性能瓶颈识别

### 常见性能问题

1. **字符串重复创建**
   - 问题：频繁创建相同的字符串对象
   - 解决：使用字符串池（SharedStringTable）

2. **内存碎片化**
   - 问题：大量小对象分配导致内存碎片
   - 解决：使用内存池或对象池

3. **XML生成效率**
   - 问题：逐个生成XML节点效率低
   - 解决：批量生成或使用流式写入

4. **ZIP压缩性能**
   - 问题：压缩大文件时CPU占用高
   - 解决：调整压缩级别或使用多线程压缩

## ⚡ 优化策略

### 1. 内存优化

#### 字符串池优化
```cpp
// 当前实现（需要优化）
for (int i = 0; i < 10000; ++i) {
    sheet->setCellValue(row_t(i), column_t(1), "重复字符串");
}

// 优化建议：预先缓存字符串
auto cached_string = workbook.cacheString("重复字符串");
for (int i = 0; i < 10000; ++i) {
    sheet->setCellValue(row_t(i), column_t(1), cached_string);
}
```

#### 内存池使用
```cpp
// 建议实现内存池管理器
class MemoryPool {
public:
    template<typename T>
    T* allocate() {
        // 从池中分配对象
    }
    
    template<typename T>
    void deallocate(T* ptr) {
        // 返回对象到池中
    }
};
```

### 2. 批量操作优化

#### 批量设置单元格值
```cpp
// 当前实现（逐个设置）
for (int row = 1; row <= 1000; ++row) {
    for (int col = 1; col <= 10; ++col) {
        sheet->setCellValue(row_t(row), column_t(col), data[row][col]);
    }
}

// 优化建议：批量设置
std::vector<CellData> batch_data;
// ... 填充 batch_data
sheet->setCellValuesBatch(batch_data);
```

#### 批量样式应用
```cpp
// 优化建议：范围样式设置
TXRange range("A1:J1000");
sheet->setRangeStyle(range, style);
```

### 3. IO优化

#### 流式写入
```cpp
// 建议实现流式XML写入器
class StreamXmlWriter {
public:
    void writeStartElement(const std::string& name);
    void writeAttribute(const std::string& name, const std::string& value);
    void writeText(const std::string& text);
    void writeEndElement();
};
```

#### 压缩优化
```cpp
// 建议的压缩配置
struct CompressionConfig {
    int level = 6;          // 压缩级别 (1-9)
    bool use_threading = true;  // 使用多线程压缩
    size_t buffer_size = 64 * 1024;  // 缓冲区大小
};
```

### 4. 算法优化

#### 单元格查找优化
```cpp
// 当前可能的实现（线性查找）
CellValue getCellValue(row_t row, column_t col) {
    for (auto& cell : cells_) {
        if (cell.row == row && cell.col == col) {
            return cell.value;
        }
    }
    return {};
}

// 优化建议：使用哈希表
std::unordered_map<CellAddress, CellValue> cell_map_;
```

#### 样式去重优化
```cpp
// 建议实现样式管理器
class StyleManager {
private:
    std::vector<TXStyle> styles_;
    std::unordered_map<TXStyle, size_t> style_index_map_;
    
public:
    size_t getStyleIndex(const TXStyle& style) {
        auto it = style_index_map_.find(style);
        if (it != style_index_map_.end()) {
            return it->second;  // 返回已存在的样式索引
        }
        
        // 添加新样式
        size_t index = styles_.size();
        styles_.push_back(style);
        style_index_map_[style] = index;
        return index;
    }
};
```

## 📊 性能基准

### 目标性能指标

| 操作类型 | 目标性能 | 当前性能 | 状态 |
|---------|---------|---------|------|
| 单元格写入 | < 100ns/cell | 待测试 | 🔄 |
| 字符串写入 | < 200ns/cell | 待测试 | 🔄 |
| 样式应用 | < 500ns/cell | 待测试 | 🔄 |
| 文件保存 | < 1s/10MB | 待测试 | 🔄 |
| 文件加载 | < 2s/10MB | 待测试 | 🔄 |

### 内存使用目标

| 数据量 | 目标内存使用 | 当前内存使用 | 状态 |
|-------|-------------|-------------|------|
| 1万单元格 | < 10MB | 待测试 | 🔄 |
| 10万单元格 | < 50MB | 待测试 | 🔄 |
| 100万单元格 | < 200MB | 待测试 | 🔄 |

## 🛠️ 性能调试工具

### 1. 内置性能计时器
```cpp
#include "tests/performance/performance_analyzer.hpp"

PerformanceTimer timer("操作名称");
// 执行需要测试的代码
// 析构时自动输出耗时
```

### 2. 内存使用监控
```cpp
size_t initial_memory = getCurrentMemoryUsage();
// 执行操作
size_t final_memory = getCurrentMemoryUsage();
std::cout << "内存增长: " << (final_memory - initial_memory) << " bytes" << std::endl;
```

### 3. 性能分析器
```cpp
Performance::PerformanceAnalyzer analyzer;
analyzer.addMetric({
    "操作名称",
    std::chrono::microseconds(duration),
    memory_used,
    operation_count,
    "分类"
});
analyzer.generateReport("performance_report.md");
```

## 🎯 优化优先级

### 高优先级
1. **字符串池实现** - 减少重复字符串内存占用
2. **批量操作API** - 提供批量设置单元格的接口
3. **内存池管理** - 减少内存分配开销

### 中优先级
1. **XML生成优化** - 使用流式写入提高效率
2. **样式去重** - 避免重复样式定义
3. **压缩优化** - 调整压缩参数平衡速度和大小

### 低优先级
1. **多线程支持** - 并行处理大数据
2. **缓存策略** - 缓存频繁访问的数据
3. **算法优化** - 优化查找和排序算法

## 📈 性能监控

### 持续性能监控
建议在CI/CD流程中集成性能测试：

```yaml
# GitHub Actions 示例
- name: Run Performance Tests
  run: |
    cmake --build build --target ExtremePerformanceTests
    ./build/tests/performance/ExtremePerformanceTests
    
- name: Upload Performance Report
  uses: actions/upload-artifact@v2
  with:
    name: performance-report
    path: performance_report.md
```

### 性能回归检测
定期运行性能测试，监控关键指标：
- 单元格写入速度
- 内存使用量
- 文件大小
- 加载时间

## 🔧 实施建议

1. **先测试，后优化** - 使用性能测试确定真正的瓶颈
2. **渐进式优化** - 一次优化一个问题，避免过度工程
3. **保持兼容性** - 优化时确保API兼容性
4. **文档更新** - 及时更新性能相关文档

通过系统的性能测试和优化，TinaXlsx可以成为一个高性能的Excel文件处理库。
