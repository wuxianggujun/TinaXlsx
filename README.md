# TinaXlsx - 高性能Excel读写库

[![C++](https://img.shields.io/badge/C%2B%2B-17-blue.svg)](https://en.cppreference.com/w/cpp/17)
[![CMake](https://img.shields.io/badge/CMake-3.16%2B-green.svg)](https://cmake.org/)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-orange.svg)]()

## 📝 简介

TinaXlsx是一个基于libxlsxwriter和xlsxio的高性能C++ Excel读写库，提供了简洁易用的API来处理Excel文件的读取、写入和对比功能。

## ✨ 特性

- 🚀 **高性能**: 优化的格式缓存机制，减少重复创建开销
- 📊 **完整功能**: 支持Excel读取、写入、格式化、公式和差异对比
- 🎯 **类型安全**: 使用现代C++17特性，提供类型安全的API
- 🔧 **易于集成**: 支持CMake，可选择编译为静态库或动态库
- 📐 **Myers算法**: 内置Myers差分算法进行精确的文件对比
- 🎨 **丰富格式**: 支持字体、颜色、边框、对齐等Excel格式

## 📦 构建要求

- **编译器**: C++17兼容编译器 (GCC 8+, Clang 7+, MSVC 2019+)
- **CMake**: 3.16或更高版本
- **依赖**: libxlsxwriter (https://github.com/jmcnamara/libxlsxwriter)
- **依赖**: xlsxio (https://github.com/jmcnamara/xlsxio)

## 🛠️ 构建选项

```bash
# 构建静态库 (默认)
cmake -DTINAXLSX_BUILD_SHARED=OFF ..

# 构建动态库
cmake -DTINAXLSX_BUILD_SHARED=ON ..

# 启用性能优化 (默认开启)
cmake -DTINAXLSX_ENABLE_PERFORMANCE=ON ..

# 构建示例程序
cmake -DTINAXLSX_BUILD_EXAMPLES=ON ..

# 构建测试程序
cmake -DTINAXLSX_BUILD_TESTS=ON ..
```

## 📚 快速开始

### 写入Excel文件

```cpp
#include "TinaXlsx/TinaXlsx.hpp"

int main() {
    // 创建写入器
    auto writer = std::make_unique<TinaXlsx::Writer>("output.xlsx");
    auto worksheet = writer->createWorksheet("Sheet1");
    
    // 创建格式
    auto format = writer->createFormat();
    format->setFontName("宋体")
          .setFontSize(12)
          .setBold()
          .setAlignment(TinaXlsx::Alignment::Center);
    
    // 写入数据
    worksheet->writeString({0, 0}, "Hello", format.get());
    worksheet->writeNumber({0, 1}, 123.45, format.get());
    
    // 保存文件
    writer->save();
    return 0;
}
```

### 读取Excel文件

```cpp
#include "TinaXlsx/TinaXlsx.hpp"

int main() {
    // 创建读取器
    TinaXlsx::Reader reader("input.xlsx");
    
    // 获取工作表名称
    auto sheetNames = reader.getSheetNames();
    
    // 打开第一个工作表
    reader.openSheet(sheetNames[0]);
    
    // 读取数据
    TinaXlsx::RowData row;
    while (reader.readNextRow(row, 10)) {
        for (const auto& cell : row) {
            std::cout << TinaXlsx::Utils::Convert::cellValueToString(cell) << "\t";
        }
        std::cout << std::endl;
    }
    
    return 0;
}
```

### 文件对比

```cpp
#include "TinaXlsx/DiffTool.hpp"

int main() {
    TinaXlsx::DiffTool::CompareOptions options;
    options.similarityThreshold = 0.8;
    options.ignoreCase = false;
    
    TinaXlsx::DiffTool diffTool(options);
    
    // 对比两个Excel文件
    auto result = diffTool.compareFiles(
        "file1.xlsx", "file2.xlsx", 
        "Sheet1", "Sheet1"
    );
    
    // 导出对比结果
    diffTool.exportResult(result, "diff_result.xlsx", {}, {}, {});
    
    std::cout << "新增: " << result.addedRowCount << " 行" << std::endl;
    std::cout << "删除: " << result.deletedRowCount << " 行" << std::endl;
    std::cout << "修改: " << result.modifiedRowCount << " 行" << std::endl;
    
    return 0;
}
```

## 🔧 CMake集成

### 作为子项目

```cmake
# 添加TinaXlsx子目录
add_subdirectory(third_party/TinaXlsx)

# 链接到你的目标
target_link_libraries(your_target PRIVATE TinaXlsx)
```

### 使用find_package

```cmake
find_package(TinaXlsx REQUIRED)
target_link_libraries(your_target PRIVATE TinaXlsx::TinaXlsx)
```

## 📈 性能优化特性

### 格式缓存
```cpp
// ✅ 推荐：使用格式缓存
auto titleFormat = writer->createFormat();
titleFormat->setFontName("宋体").setFontSize(16).setBold();

for (int i = 0; i < 1000; ++i) {
    worksheet->writeString({i, 0}, "Title", titleFormat.get());
}

// ❌ 避免：每次创建新格式
for (int i = 0; i < 1000; ++i) {
    auto format = writer->createFormat();
    format->setFontName("宋体").setFontSize(16).setBold();
    worksheet->writeString({i, 0}, "Title", format.get());
}
```

### 批量写入
```cpp
// ✅ 推荐：批量设置行列属性
worksheet->setColumnWidth(0, 20);
worksheet->setColumnWidth(1, 15);
worksheet->setRowHeight(0, 25);

// 批量写入数据
std::vector<std::vector<std::string>> data = {
    {"Name", "Age", "City"},
    {"Alice", "25", "Beijing"},
    {"Bob", "30", "Shanghai"}
};

for (size_t row = 0; row < data.size(); ++row) {
    for (size_t col = 0; col < data[row].size(); ++col) {
        worksheet->writeString({row, col}, data[row][col], format.get());
    }
}
```

## 📋 API参考

### 核心类

- **Writer**: Excel文件写入器
- **Reader**: Excel文件读取器  
- **Worksheet**: 工作表操作
- **Format**: 单元格格式设置
- **DiffTool**: 文件差异对比工具

### 数据类型

- **CellValue**: 单元格值 (std::variant<std::monostate, double, int, std::string>)
- **RowData**: 行数据 (std::vector<CellValue>)
- **CellPosition**: 单元格位置 {row, col}
- **CellRange**: 单元格范围 {start, end}

### 枚举类型

```cpp
enum class Alignment { Left, Center, Right };
enum class VerticalAlignment { Top, VCenter, Bottom };
enum class BorderStyle { None, Thin, Medium, Thick };
```

## 🚀 性能基准

基于实际测试数据：

| 操作类型 | 数据量 | 处理时间 | 性能 |
|---------|--------|----------|------|
| 读取Excel | 238行 | ~12秒 | 19行/秒 |
| 写入Excel | 1000行 | ~2秒 | 500行/秒 |
| 格式创建 | 缓存vs重复 | 10x+提升 | 90%+减少 |

## 🐛 性能优化建议

基于你的性能统计（平均每行52ms），以下是优化建议：

### 1. 减少调试输出
```cpp
// 在Release模式下禁用详细日志
#ifndef NDEBUG
    qDebug() << "详细调试信息";
#endif
```

### 2. 批量操作
```cpp
// 批量设置属性而不是逐个设置
worksheet->setColumnWidths({{0, 20}, {1, 15}, {2, 30}});
```

### 3. 内存预分配
```cpp
// 预分配vector容量
std::vector<std::string> row;
row.reserve(10);  // 预分配10列
```

### 4. 异步处理
```cpp
// 对于大文件，考虑使用异步处理
std::async(std::launch::async, [&]() {
    // Excel处理逻辑
});
```

## 📄 许可证

MIT License - 详见 [LICENSE](LICENSE) 文件

## 🤝 贡献指南

1. Fork项目
2. 创建功能分支 (`git checkout -b feature/amazing-feature`)
3. 提交更改 (`git commit -m 'Add amazing feature'`)
4. 推送分支 (`git push origin feature/amazing-feature`)
5. 创建Pull Request

## 📞 支持

- 🐛 [问题反馈](../../issues)
- 💡 [功能请求](../../issues)
- 📖 [文档wiki](../../wiki) 
