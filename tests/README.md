# TinaXlsx 测试结构说明

## 测试目录组织

本项目的测试采用分层结构，便于独立测试和全局测试：

```
tests/
├── unit/               # 单元测试
│   ├── test_cell.cpp          # TXCell 类单元测试
│   ├── test_xml_handler.cpp   # TXXmlHandler 类单元测试
│   └── test_zip_handler.cpp   # TXZipHandler 类单元测试
├── integration/        # 集成测试
│   ├── test_main.cpp          # TXWorkbook 集成测试
│   ├── debug_xml_test.cpp     # XML 处理调试测试
│   └── debug_zip_test.cpp     # ZIP 处理调试测试
├── examples/           # 示例和演示测试
│   ├── example_sales_data.cpp # 实际销售数据处理示例
│   └── example_performance.cpp # 性能测试示例
└── README.md          # 本文档
```

## 运行测试

### 运行所有测试
```bash
cd build
./Debug/TinaXlsx_test.exe
```

### 运行特定类别的测试

#### 单元测试
```bash
./Debug/TinaXlsx_test.exe --gtest_filter="*TXCellTest*:*TXXmlHandlerTest*:*TXZipHandlerTest*"
```

#### 集成测试
```bash
./Debug/TinaXlsx_test.exe --gtest_filter="*TXWorkbookTest*:*DebugXmlTest*:*DebugZipTest*"
```

#### 示例测试
```bash
./Debug/TinaXlsx_test.exe --gtest_filter="*SalesDataExampleTest*:*PerformanceExampleTest*"
```

### 运行特定测试

#### 单个组件测试
```bash
# 仅测试 TXCell
./Debug/TinaXlsx_test.exe --gtest_filter="TXCellTest.*"

# 仅测试 XML 处理
./Debug/TinaXlsx_test.exe --gtest_filter="TXXmlHandlerTest.*"

# 仅测试 ZIP 处理
./Debug/TinaXlsx_test.exe --gtest_filter="TXZipHandlerTest.*"
```

#### 销售数据示例
```bash
./Debug/TinaXlsx_test.exe --gtest_filter="SalesDataExampleTest.*"
```

#### 性能测试
```bash
./Debug/TinaXlsx_test.exe --gtest_filter="PerformanceExampleTest.*"
```

## 测试说明

### 单元测试 (unit/)
- **test_cell.cpp**: 测试 TXCell 类的所有功能，包括类型转换、赋值操作符、移动语义等
- **test_xml_handler.cpp**: 测试 XML 解析、生成、XPath 查询等功能
- **test_zip_handler.cpp**: 测试 ZIP 文件的创建、读取、压缩等功能

### 集成测试 (integration/)
- **test_main.cpp**: 测试 TXWorkbook 的完整工作流程
- **debug_xml_test.cpp**: XML 声明和编码的调试测试
- **debug_zip_test.cpp**: ZIP 压缩和写入的调试测试

### 示例测试 (examples/)
- **example_sales_data.cpp**: 
  - 演示如何创建实际的销售数据报表
  - 包含多种数据类型的处理
  - 展示多工作表功能
  - 验证数据完整性
  
- **example_performance.cpp**:
  - 大数据量处理的性能测试 (1000+ 行)
  - 批量操作性能对比
  - 多工作表性能测试
  - 文件大小和时间统计

## 测试特点

### 🎯 全面覆盖
- **100% API 覆盖**: 所有公共接口都有对应测试
- **类型安全**: 测试 std::variant 的正确使用
- **错误处理**: 测试异常情况和边界条件

### 🚀 性能验证
- **大数据处理**: 验证处理1000+行数据的能力
- **批量操作**: 测试高效的批量数据写入
- **内存管理**: 确保无内存泄漏

### 💼 实用示例
- **真实场景**: 销售数据、财务报表等实际用例
- **最佳实践**: 展示正确的 API 使用方法
- **文档化**: 每个示例都有详细注释

## 测试结果

当前测试通过率: **49/49 (100%)** ✅

### 分类统计
- 单元测试: 42/42 通过 (100%)
- 集成测试: 4/4 通过 (100%)
- 示例测试: 3/3 通过 (100%)

## 如何添加新测试

### 添加单元测试
1. 在 `unit/` 目录下创建 `test_[component].cpp`
2. 使用 Google Test 框架
3. 遵循现有的命名约定

### 添加示例测试
1. 在 `examples/` 目录下创建 `example_[scenario].cpp`
2. 重点演示实际使用场景
3. 包含性能统计和数据验证

### 测试规范
- 每个测试类使用 `::testing::Test` 基类
- 在 `SetUp()`/`TearDown()` 中处理测试环境
- 使用描述性的测试名称
- 添加必要的错误消息和断言 