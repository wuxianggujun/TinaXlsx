# TinaXlsx 使用指南

## 🚀 快速开始

### 基本设置

```cpp
#include "TinaXlsx/TinaXlsx.hpp"

int main() {
    // 初始化库
    if (!TinaXlsx::initialize()) {
        std::cerr << "库初始化失败" << std::endl;
        return -1;
    }
    
    // 你的代码...
    
    // 清理资源
    TinaXlsx::cleanup();
    return 0;
}
```

### 推荐使用方式：内存优先 API

```cpp
using namespace TinaXlsx;

// 创建内存优先工作簿
auto workbook = TXInMemoryWorkbook::create("example.xlsx");
auto& sheet = workbook->createSheet("数据表");

// 批量设置数值 - 高性能方式
std::vector<double> numbers = {1.0, 2.0, 3.0, 4.0, 5.0};
std::vector<TXCoordinate> coords = {
    TXCoordinate(0, 0), TXCoordinate(0, 1), TXCoordinate(0, 2),
    TXCoordinate(0, 3), TXCoordinate(0, 4)
};
sheet.batchSetNumbers(coords, numbers);

// 保存文件
auto result = workbook->save();
if (!result.isSuccess()) {
    std::cerr << "保存失败: " << result.getError().getMessage() << std::endl;
}
```

## 🎯 核心功能使用

### 1. 工作簿和工作表管理

#### 创建工作簿
```cpp
// 内存优先工作簿（推荐）
auto workbook = TXInMemoryWorkbook::create("my_file.xlsx");

// 传统工作簿（兼容性）
TXWorkbook traditional_workbook("legacy_file.xlsx");
```

#### 工作表操作
```cpp
// 创建工作表
auto& sheet1 = workbook->createSheet("销售数据");
auto& sheet2 = workbook->createSheet("统计报表");

// 获取工作表
auto* sheet = workbook->getSheet("销售数据");
auto* sheet_by_index = workbook->getSheet(0);

// 删除工作表
workbook->removeSheet("不需要的表");
```

### 2. 单元格操作

#### 单个单元格设置
```cpp
// 设置数值
sheet.setNumber(TXCoordinate(0, 0), 123.45);

// 设置字符串
sheet.setString(TXCoordinate(0, 1), "Hello World");

// 设置公式
sheet.setFormula(TXCoordinate(0, 2), "=A1*2");

// 设置布尔值
sheet.setBoolean(TXCoordinate(0, 3), true);
```

#### 批量操作（高性能）
```cpp
// 批量设置数值
std::vector<double> values = {1.0, 2.0, 3.0, 4.0, 5.0};
std::vector<TXCoordinate> coords;
for (int i = 0; i < 5; ++i) {
    coords.emplace_back(0, i);
}
sheet.batchSetNumbers(coords, values);

// 批量设置字符串
std::vector<std::string> texts = {"A", "B", "C", "D", "E"};
sheet.batchSetStrings(coords, texts);
```

### 3. 样式设置

#### 基本样式
```cpp
// 创建样式
TXCellStyle style;
style.font.name = "Arial";
style.font.size = 12;
style.font.bold = true;
style.font.color = TXColor::fromRGB(255, 0, 0); // 红色

// 应用样式
sheet.setCellStyle(TXCoordinate(0, 0), style);
```

#### 批量样式应用
```cpp
// 为范围应用样式
TXRange range(TXCoordinate(0, 0), TXCoordinate(9, 4)); // A1:E10
sheet.setRangeStyle(range, style);
```

### 4. 高级功能

#### 合并单元格
```cpp
// 合并单元格范围
TXRange merge_range(TXCoordinate(0, 0), TXCoordinate(0, 2)); // A1:C1
sheet.mergeCells(merge_range);

// 取消合并
sheet.unmergeCells(merge_range);
```

#### 数据验证
```cpp
// 创建数据验证规则
TXDataValidation validation;
validation.setType(TXDataValidation::Type::List);
validation.setFormula1("选项1,选项2,选项3");
validation.setErrorMessage("请选择有效选项");

// 应用到范围
TXRange validation_range(TXCoordinate(1, 0), TXCoordinate(10, 0));
sheet.setDataValidation(validation_range, validation);
```

## ⚡ 性能优化技巧

### 1. 使用批量操作
```cpp
// ❌ 低效：逐个设置
for (int i = 0; i < 10000; ++i) {
    sheet.setNumber(TXCoordinate(i / 100, i % 100), i * 1.5);
}

// ✅ 高效：批量设置
std::vector<double> values(10000);
std::vector<TXCoordinate> coords(10000);
for (int i = 0; i < 10000; ++i) {
    values[i] = i * 1.5;
    coords[i] = TXCoordinate(i / 100, i % 100);
}
sheet.batchSetNumbers(coords, values);
```

### 2. 内存布局优化
```cpp
// 启用自动优化
sheet.enableAutoOptimization(true);

// 手动触发优化
sheet.optimizeLayout();

// 获取性能统计
auto stats = sheet.getPerformanceStats();
std::cout << "处理单元格数: " << stats.total_cells << std::endl;
std::cout << "平均操作时间: " << stats.avg_operation_time << "ms" << std::endl;
```

### 3. 内存管理配置
```cpp
// 配置统一内存管理器
TXUnifiedMemoryManager::Config config;
config.memory_limit = 2ULL * 1024 * 1024 * 1024; // 2GB 限制
config.warning_threshold_mb = 1536; // 1.5GB 警告
config.enable_monitoring = true;

// 初始化全局内存管理器
GlobalUnifiedMemoryManager::initialize(config);
```

## 🔧 错误处理

### Result 模式使用
```cpp
// 安全的操作方式
auto result = sheet.setNumber(TXCoordinate(0, 0), 123.45);
if (result.isSuccess()) {
    std::cout << "设置成功" << std::endl;
} else {
    std::cerr << "设置失败: " << result.getError().getMessage() << std::endl;
}

// 链式操作
auto final_result = sheet.setNumber(TXCoordinate(0, 0), 123.45)
    .and_then([&](auto) { return sheet.setString(TXCoordinate(0, 1), "Test"); })
    .and_then([&](auto) { return workbook->save(); });

if (!final_result.isSuccess()) {
    std::cerr << "操作失败: " << final_result.getError().getMessage() << std::endl;
}
```

## 📊 监控和调试

### 性能监控
```cpp
// 获取工作表性能统计
auto sheet_stats = sheet.getPerformanceStats();
std::cout << "总单元格数: " << sheet_stats.total_cells << std::endl;
std::cout << "批量操作次数: " << sheet_stats.batch_operations << std::endl;
std::cout << "缓存命中率: " << sheet_stats.cache_hit_ratio << std::endl;

// 获取内存统计
auto memory_stats = GlobalUnifiedMemoryManager::getInstance().getUnifiedStats();
std::cout << "总内存使用: " << memory_stats.total_memory_usage << " bytes" << std::endl;
std::cout << "内存效率: " << memory_stats.overall_efficiency << "%" << std::endl;
```

### 调试信息
```cpp
// 启用详细日志（调试版本）
#ifdef DEBUG
    sheet.enableDebugLogging(true);
    workbook->enableVerboseLogging(true);
#endif

// 获取构建信息
std::cout << TinaXlsx::getBuildInfo() << std::endl;
```

## 🎨 最佳实践

### 1. 项目结构建议
```cpp
class ExcelReportGenerator {
private:
    std::unique_ptr<TXInMemoryWorkbook> workbook_;
    
public:
    ExcelReportGenerator(const std::string& filename) 
        : workbook_(TXInMemoryWorkbook::create(filename)) {}
    
    void generateSalesReport(const std::vector<SalesData>& data) {
        auto& sheet = workbook_->createSheet("销售报表");
        
        // 设置标题
        setupHeaders(sheet);
        
        // 批量填充数据
        fillData(sheet, data);
        
        // 应用样式
        applyStyles(sheet);
    }
    
    TXResult<void> save() {
        return workbook_->save();
    }
};
```

### 2. 资源管理
```cpp
// RAII 风格的资源管理
class TinaXlsxSession {
public:
    TinaXlsxSession() {
        if (!TinaXlsx::initialize()) {
            throw std::runtime_error("TinaXlsx 初始化失败");
        }
    }
    
    ~TinaXlsxSession() {
        TinaXlsx::cleanup();
    }
    
    // 禁止拷贝
    TinaXlsxSession(const TinaXlsxSession&) = delete;
    TinaXlsxSession& operator=(const TinaXlsxSession&) = delete;
};

// 使用方式
int main() {
    try {
        TinaXlsxSession session; // 自动初始化和清理
        
        // 你的代码...
        
    } catch (const std::exception& e) {
        std::cerr << "错误: " << e.what() << std::endl;
        return -1;
    }
    
    return 0;
}
```

### 3. 性能优化检查清单
- ✅ 使用内存优先 API (`TXInMemoryWorkbook`)
- ✅ 优先使用批量操作
- ✅ 启用自动内存布局优化
- ✅ 合理配置内存管理参数
- ✅ 监控性能统计信息
- ✅ 在大数据场景下使用 SIMD 优化
- ✅ 避免频繁的单个单元格操作
- ✅ 使用字符串池减少内存占用
