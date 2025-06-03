# TinaXlsx CMake 测试工具函数使用指南

## 🎯 概述

为了简化测试配置并减少重复代码，我们创建了统一的CMake工具函数来管理测试。这些函数封装了常见的测试配置模式，使测试的添加和维护变得更加简单。

## 📁 文件结构

```
cmake/
└── TestUtils.cmake          # 测试工具函数定义

tests/
├── unit/
│   ├── CMakeLists.txt       # 原有配置（复杂）
│   └── CMakeLists_new.txt   # 新配置示例（简化）
└── performance/
    ├── CMakeLists.txt       # 原有配置（复杂）
    └── CMakeLists_new.txt   # 新配置示例（简化）
```

## 🚀 核心函数

### 1. `create_test_executable()`

创建完整配置的测试可执行文件。

```cmake
create_test_executable(
    TARGET_NAME MyTests                    # 目标名称
    SOURCES test_my_feature.cpp test_file_generator.cpp  # 源文件
    COMMENT "Running my feature tests"    # 注释
    PERFORMANCE_TEST                       # 可选：标记为性能测试
    DEPENDENCIES SomeOtherTarget          # 可选：依赖项
)
```

**自动配置：**
- ✅ C++17标准
- ✅ 编译选项（MSVC: /W3 /utf-8 /wd4819，GCC: -Wall）
- ✅ Windows特定定义（WIN32_LEAN_AND_MEAN等）
- ✅ 包含目录（include、tests/unit、googletest）
- ✅ 库链接（TinaXlsx、gtest_main、gtest）
- ✅ 输出目录设置
- ✅ CTest注册
- ✅ 单独运行目标（Run<TargetName>）
- ✅ 添加到全局测试列表

### 2. `add_unit_test()` 和 `add_performance_test()`

简化的宏，自动推断配置。

```cmake
# 单元测试（自动使用 test_<name>.cpp 和 test_file_generator.cpp）
add_unit_test(DataFilter)  # 创建 DataFilterTests

# 性能测试（自动使用 test_<name>_performance.cpp）
add_performance_test(Memory)  # 创建 MemoryPerformanceTests

# 自定义源文件
add_unit_test(CustomTest test_custom.cpp test_helper.cpp)
```

### 3. `create_batch_test_targets()`

自动创建批量运行目标。

```cmake
# 在主CMakeLists.txt中调用
create_batch_test_targets()
```

**创建的目标：**
- `RunAllUnitTests` - 运行所有单元测试
- `RunAllPerformanceTests` - 运行所有性能测试

### 4. `create_quick_test_target()`

创建快速测试目标（只运行关键测试）。

```cmake
create_quick_test_target(
    TESTS 
        BasicTests 
        DataFilterTests 
        SimplePerformanceTests
)
```

### 5. `print_test_summary()`

打印测试配置摘要。

```cmake
# 在主CMakeLists.txt末尾调用
print_test_summary()
```

## 📊 配置对比

### 旧方式（复杂，重复）

```cmake
# 每个测试需要15-20行配置
add_executable(DataFilterTests test_data_filter.cpp test_file_generator.cpp)

target_link_libraries(DataFilterTests
    PRIVATE
    TinaXlsx
    gtest_main
    gtest
)

target_include_directories(DataFilterTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/third_party/googletest/googletest/include
)

target_compile_features(DataFilterTests PRIVATE cxx_std_17)

if(MSVC)
    target_compile_options(DataFilterTests PRIVATE /W3 /utf-8 /wd4819)
    target_compile_definitions(DataFilterTests PRIVATE WIN32_LEAN_AND_MEAN NOMINMAX _CRT_SECURE_NO_WARNINGS)
endif()

set_target_properties(DataFilterTests PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/tests/unit
)

add_test(NAME DataFilter COMMAND DataFilterTests)
set_tests_properties(DataFilter PROPERTIES WORKING_DIRECTORY ${CMAKE_BINARY_DIR})

add_custom_target(RunDataFilterTest
    COMMAND DataFilterTests
    DEPENDS DataFilterTests
    COMMENT "Running data filter tests"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)
```

### 新方式（简洁，统一）

```cmake
# 一行搞定所有配置
create_test_executable(
    TARGET_NAME DataFilterTests
    SOURCES test_data_filter.cpp test_file_generator.cpp
    COMMENT "Running data filter tests"
)

# 或者更简单
add_unit_test(DataFilter test_data_filter.cpp test_file_generator.cpp)
```

## 🎯 使用示例

### 单元测试配置

```cmake
# tests/unit/CMakeLists.txt

# 基础功能测试
create_test_executable(
    TARGET_NAME BasicTests
    SOURCES test_basic_features.cpp test_cell_formatting.cpp test_file_generator.cpp
    COMMENT "Running basic functionality tests"
)

# 图表功能测试
create_test_executable(
    TARGET_NAME ChartTests
    SOURCES test_chart_functionality.cpp test_chart_refactoring.cpp test_file_generator.cpp
    COMMENT "Running chart functionality tests"
)

# 使用简化宏
add_unit_test(DataFilter)
add_unit_test(CellManager)
```

### 性能测试配置

```cmake
# tests/performance/CMakeLists.txt

# 字符串性能测试
create_test_executable(
    TARGET_NAME SharedStringsPerformanceTests
    SOURCES test_shared_strings_performance.cpp
    COMMENT "Running shared strings performance tests"
    PERFORMANCE_TEST
)

# 并行框架性能测试
create_test_executable(
    TARGET_NAME AdvancedParallelPerformanceTests
    SOURCES test_advanced_parallel_performance.cpp
    COMMENT "Running advanced parallel framework performance tests"
    PERFORMANCE_TEST
)

# 使用简化宏
add_performance_test(Memory)
add_performance_test(Streaming)
```

## 🔧 运行测试

### 构建时可用的目标

```bash
# 运行所有测试
cmake --build build --target run_all_tests

# 运行所有单元测试
cmake --build build --target RunAllUnitTests

# 运行所有性能测试
cmake --build build --target RunAllPerformanceTests

# 运行快速测试
cmake --build build --target RunQuickTests

# 运行单个测试
cmake --build build --target RunBasicTests
cmake --build build --target RunSharedStringsPerformanceTests
```

### 直接运行可执行文件

```bash
# 单元测试
./build/tests/unit/BasicTests
./build/tests/unit/DataFilterTests

# 性能测试
./build/tests/performance/SharedStringsPerformanceTests
./build/tests/performance/AdvancedParallelPerformanceTests
```

### 使用CTest

```bash
# 运行所有测试
ctest

# 运行特定测试
ctest -R BasicTests
ctest -R Performance

# 详细输出
ctest --output-on-failure --verbose
```

## 📈 优势总结

### 1. **代码减少**
- 从每个测试15-20行配置减少到1-5行
- 减少90%的重复代码

### 2. **统一配置**
- 所有测试使用相同的编译选项
- 统一的目录结构和命名规范
- 一致的错误处理和设置

### 3. **易于维护**
- 修改配置只需要在TestUtils.cmake中修改
- 添加新测试只需要一行代码
- 自动化的依赖管理

### 4. **功能增强**
- 自动创建批量运行目标
- 自动注册到CTest
- 支持快速测试配置
- 详细的配置摘要

### 5. **错误减少**
- 避免手动配置中的遗漏
- 统一的最佳实践
- 自动化的设置验证

## 🚀 迁移指南

1. **包含工具函数**：在主CMakeLists.txt中添加 `include(TestUtils)`
2. **替换现有配置**：使用 `create_test_executable()` 替换手动配置
3. **添加批量目标**：在主CMakeLists.txt中调用 `create_batch_test_targets()`
4. **添加摘要**：在主CMakeLists.txt末尾调用 `print_test_summary()`
5. **测试验证**：确保所有测试仍然正常工作

通过使用这些工具函数，TinaXlsx的测试配置变得更加简洁、统一和易于维护！
