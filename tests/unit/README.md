# TinaXlsx 单元测试

本目录包含TinaXlsx项目的单元测试，特别是重构后的管理器组件的测试。

## 测试结构

### 原有测试
- `test_basic_features.cpp` - 基本功能测试
- `test_column_width_row_height.cpp` - 行列尺寸测试
- `test_sheet_protection.cpp` - 工作表保护测试
- `test_enhanced_formulas.cpp` - 增强公式测试
- `test_cell_locking.cpp` - 单元格锁定测试

### 重构后的新测试
- `test_cell_manager.cpp` - 单元格管理器测试
- `test_row_column_manager.cpp` - 行列管理器测试
- `test_sheet_protection_manager.cpp` - 保护管理器测试
- `test_sheet_refactored.cpp` - 重构后工作表集成测试

## 编译和运行

### 编译测试

```bash
# 在项目根目录
mkdir build && cd build
cmake .. -DBUILD_TESTING=ON
make TinaXlsx_UnitTests
```

### 运行所有测试

```bash
# 方法1：使用CTest
ctest --output-on-failure --verbose

# 方法2：直接运行测试可执行文件
./TinaXlsx_UnitTests

# 方法3：使用自定义目标
make run_unit_tests
```

### 运行特定测试

```bash
# 运行快速测试（基本功能）
make run_quick_tests

# 运行重构组件测试
make run_refactored_tests

# 运行管理器测试
make run_manager_tests

# 运行集成测试
make run_integration_tests

# 运行特定测试类
./TinaXlsx_UnitTests --gtest_filter="TXCellManagerTest.*"

# 运行特定测试用例
./TinaXlsx_UnitTests --gtest_filter="TXCellManagerTest.BasicCellOperations"
```

### 测试输出选项

```bash
# 详细输出
./TinaXlsx_UnitTests --gtest_verbose

# 只显示失败的测试
./TinaXlsx_UnitTests --gtest_brief

# 生成XML报告
./TinaXlsx_UnitTests --gtest_output=xml:test_results.xml

# 重复运行测试（用于检测不稳定的测试）
./TinaXlsx_UnitTests --gtest_repeat=10
```

## 测试覆盖的功能

### TXCellManager 测试
- ✅ 基本单元格操作（创建、获取、设置值）
- ✅ 多种数据类型支持（字符串、数字、布尔值）
- ✅ 批量操作（批量设置值、批量获取值）
- ✅ 范围操作（使用范围、最大行列）
- ✅ 坐标变换（用于行列插入删除）
- ✅ 单元格计数和统计
- ✅ 迭代器支持
- ✅ 清空操作

### TXRowColumnManager 测试
- ✅ 行操作（插入、删除、设置高度、隐藏）
- ✅ 列操作（插入、删除、设置宽度、隐藏）
- ✅ 自动调整（自动列宽、自动行高）
- ✅ 批量操作（批量设置尺寸）
- ✅ 边界条件验证
- ✅ 默认值处理
- ✅ 清空操作

### TXSheetProtectionManager 测试
- ✅ 基本保护功能（保护、取消保护）
- ✅ 密码保护（设置密码、验证密码）
- ✅ 保护选项（严格保护、宽松保护）
- ✅ 单元格锁定（单个、范围、批量）
- ✅ 权限验证（操作权限检查）
- ✅ 可编辑性检查（单元格、范围）
- ✅ 保护状态查询（统计信息）
- ✅ 预定义保护模板

### TXSheetRefactored 集成测试
- ✅ 基本属性和单元格操作
- ✅ 行列操作集成
- ✅ 工作表保护集成
- ✅ 公式操作集成
- ✅ 合并单元格集成
- ✅ 样式操作集成
- ✅ 数字格式化集成
- ✅ 批量操作集成
- ✅ 范围操作集成
- ✅ 查询操作集成

## 测试数据和断言

### 测试数据
测试使用各种类型的数据：
- 字符串：`"Hello"`, `"Test"`, `"Very long text content"`
- 数字：`123.45`, `42.0`, `1234.567`
- 布尔值：`true`, `false`
- 坐标：`TXCoordinate(row_t(1), column_t(1))`
- 范围：`TXRange(start, end)`

### 常用断言
- `EXPECT_TRUE()` / `EXPECT_FALSE()` - 布尔值断言
- `EXPECT_EQ()` / `EXPECT_NE()` - 相等性断言
- `EXPECT_DOUBLE_EQ()` - 浮点数相等断言
- `EXPECT_GT()` / `EXPECT_LT()` - 大小比较断言
- `ASSERT_NE()` - 非空指针断言（失败时停止测试）

## 调试测试

### 调试单个测试
```bash
# 使用GDB调试
gdb ./TinaXlsx_UnitTests
(gdb) run --gtest_filter="TXCellManagerTest.BasicCellOperations"
(gdb) bt  # 查看调用栈
```

### 内存检查
```bash
# 使用Valgrind检查内存泄漏
valgrind --leak-check=full ./TinaXlsx_UnitTests

# 使用AddressSanitizer
cmake .. -DCMAKE_CXX_FLAGS="-fsanitize=address -g"
make TinaXlsx_UnitTests
./TinaXlsx_UnitTests
```

### 性能分析
```bash
# 使用perf分析性能
perf record ./TinaXlsx_UnitTests
perf report
```

## 添加新测试

### 创建新测试文件
1. 在`tests/unit/`目录下创建新的`.cpp`文件
2. 包含必要的头文件和Google Test
3. 创建测试类继承自`::testing::Test`
4. 实现`SetUp()`和`TearDown()`方法
5. 编写测试用例

### 示例测试结构
```cpp
#include <gtest/gtest.h>
#include "TinaXlsx/YourClass.hpp"

class YourClassTest : public ::testing::Test {
protected:
    void SetUp() override {
        // 初始化测试数据
    }
    
    void TearDown() override {
        // 清理测试数据
    }
    
    // 测试成员变量
};

TEST_F(YourClassTest, TestCaseName) {
    // 测试代码
    EXPECT_TRUE(condition);
}
```

### 更新CMakeLists.txt
将新的测试文件添加到`UNIT_TEST_SOURCES`列表中。

## 持续集成

这些测试设计为可以在CI/CD环境中自动运行：

```yaml
# GitHub Actions 示例
- name: Run Unit Tests
  run: |
    cd build
    ctest --output-on-failure --verbose
```

## 测试最佳实践

1. **独立性**：每个测试应该独立，不依赖其他测试的结果
2. **可重复性**：测试应该可以重复运行并得到相同结果
3. **快速性**：单元测试应该快速执行
4. **清晰性**：测试名称应该清楚地描述测试的内容
5. **完整性**：测试应该覆盖正常情况、边界情况和错误情况

## 故障排除

### 常见问题
1. **编译错误**：检查头文件路径和依赖库
2. **链接错误**：确保所有必要的源文件都被包含
3. **测试失败**：检查测试数据和预期结果
4. **内存错误**：使用内存检查工具定位问题

### 获取帮助
- 查看测试输出的详细错误信息
- 使用调试器逐步执行测试
- 检查相关的源代码实现
- 参考其他类似的测试用例
