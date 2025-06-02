# TinaXlsx 测试指南

本文档介绍如何运行 TinaXlsx 项目的各种测试。

## 🎯 测试结构

### 独立测试可执行文件

我们为不同的功能模块创建了独立的测试可执行文件，可以在IDE中直接运行：

| 测试名称 | 可执行文件 | 测试内容 |
|---------|-----------|----------|
| **DataFilterTests** | `DataFilterTests.exe` | 数据筛选功能（AutoFilter、数值筛选等） |
| **NumberUtilsTests** | `NumberUtilsTests.exe` | 数值工具（fast_float 性能测试） |
| **DataFeaturesTests** | `DataFeaturesTests.exe` | 数据功能（公式、数据验证等） |
| **ChartTests** | `ChartTests.exe` | 图表功能（图表创建、样式等） |
| **BasicTests** | `BasicTests.exe` | 基础功能（单元格、格式化等） |

## 🚀 运行方式

### 方式1：IDE中直接运行（推荐）

1. **编译项目**：
   ```bash
   cmake --build cmake-build-debug
   ```

2. **在IDE中运行**：
   - 在项目资源管理器中找到对应的测试可执行文件
   - 右键点击 → "运行" 或 "调试"
   - 例如：`cmake-build-debug/tests/unit/DataFilterTests.exe`

### 方式2：使用批处理脚本

```bash
# 运行特定测试
run_tests.bat DataFilter
run_tests.bat NumberUtils
run_tests.bat DataFeatures
run_tests.bat Charts
run_tests.bat Basic

# 运行所有测试
run_tests.bat All
```

### 方式3：命令行直接运行

```bash
cd cmake-build-debug

# 运行特定测试
tests/unit/DataFilterTests.exe
tests/unit/NumberUtilsTests.exe
tests/unit/DataFeaturesTests.exe
tests/unit/ChartTests.exe
tests/unit/BasicTests.exe
```

### 方式4：使用CMake目标

```bash
# 运行所有独立测试
cmake --build cmake-build-debug --target RunAllIndependentTests

# 运行快速测试
cmake --build cmake-build-debug --target RunQuickTests

# 运行特定测试
cmake --build cmake-build-debug --target RunDataFilterTest
cmake --build cmake-build-debug --target RunNumberUtilsTest
```

### 方式5：使用CTest

```bash
cd cmake-build-debug

# 运行所有测试
ctest --output-on-failure

# 运行特定测试
ctest -R DataFilter
ctest -R NumberUtils
```

## 📊 测试详情

### DataFilterTests - 数据筛选测试
- **测试文件**: `test_data_filter.cpp`
- **主要功能**:
  - AutoFilter 基础功能
  - 数值筛选（解决Excel兼容性问题）
  - 文本筛选
  - 范围筛选
  - 多条件筛选

### NumberUtilsTests - 数值工具测试
- **测试文件**: `test_number_utils.cpp`
- **主要功能**:
  - fast_float 高性能解析
  - 数值格式化
  - Excel XML兼容性
  - 性能对比测试

### DataFeaturesTests - 数据功能测试
- **测试文件**: `test_data_features.cpp`
- **主要功能**:
  - 公式计算
  - 数据验证
  - 条件格式
  - 数据排序

### ChartTests - 图表功能测试
- **测试文件**: `test_chart_*.cpp`
- **主要功能**:
  - 图表创建
  - 图表样式
  - 多系列支持
  - 图表重构

### BasicTests - 基础功能测试
- **测试文件**: `test_basic_features.cpp`, `test_cell_formatting.cpp`
- **主要功能**:
  - 单元格操作
  - 格式化
  - 样式设置
  - 基础Excel功能

## 🔧 故障排除

### 问题1：可执行文件不存在
**解决方案**：
```bash
# 重新编译项目
cmake --build cmake-build-debug --target DataFilterTests
cmake --build cmake-build-debug --target NumberUtilsTests
# ... 其他测试
```

### 问题2：测试失败
**解决方案**：
1. 检查输出目录是否存在：`cmake-build-debug/tests/unit/test_output/`
2. 确保有写入权限
3. 查看详细错误信息

### 问题3：IDE无法运行测试
**解决方案**：
1. 确保项目已正确配置CMake
2. 重新生成CMake缓存
3. 使用批处理脚本作为替代方案

## 📁 输出文件

测试运行后会在以下目录生成Excel文件：
```
cmake-build-debug/tests/unit/test_output/
├── DataFilterTest/
│   ├── data_filter_test.xlsx
│   └── advanced_filter_test.xlsx
├── DataFeaturesTest/
│   ├── data_validation_test.xlsx
│   └── formula_test.xlsx
└── ...
```

这些文件可以用Excel、WPS或LibreOffice打开验证功能。

## 🎯 推荐工作流程

1. **开发新功能时**：
   ```bash
   # 运行相关测试
   run_tests.bat DataFilter
   ```

2. **提交代码前**：
   ```bash
   # 运行快速测试
   run_tests.bat All
   ```

3. **调试特定问题**：
   - 在IDE中直接运行对应的测试可执行文件
   - 使用调试器设置断点

4. **性能测试**：
   ```bash
   # 运行性能相关测试
   run_tests.bat NumberUtils
   ```

## 💡 提示

- **IDE集成**：大多数现代IDE（Visual Studio、CLion等）会自动识别这些可执行文件
- **并行运行**：独立的测试可执行文件可以并行运行，提高测试效率
- **选择性测试**：只运行你正在开发的功能相关的测试，节省时间
- **输出验证**：生成的Excel文件可以手动验证功能正确性
