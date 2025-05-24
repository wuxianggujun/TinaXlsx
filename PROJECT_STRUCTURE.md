# TinaXlsx 项目结构

## 概述
TinaXlsx 项目已经重新组织为更清晰、更符合C++项目最佳实践的目录结构。

## 目录结构

```
TinaXlsx/
├── README.md                    # 项目主要文档
├── LICENSE                      # 许可证文件
├── CMakeLists.txt              # 主CMake配置文件
├── .gitignore                  # Git忽略文件配置
│
├── include/                    # 公共头文件目录
│   └── TinaXlsx/
│       ├── TinaXlsx.hpp       # 主要包含文件
│       ├── Reader.hpp         # Excel读取器头文件
│       ├── Writer.hpp         # Excel写入器头文件
│       ├── Types.hpp          # 类型定义
│       ├── Exception.hpp      # 异常处理
│       ├── Format.hpp         # 格式处理
│       └── ...                # 其他头文件
│
├── src/                       # 源文件目录
│   ├── Reader.cpp            # Excel读取器实现
│   ├── Writer.cpp            # Excel写入器实现
│   ├── Types.cpp             # 类型实现
│   ├── Format.cpp            # 格式处理实现
│   └── ...                   # 其他源文件
│
├── examples/                  # 示例程序目录
│   ├── CMakeLists.txt        # 示例程序构建配置
│   ├── example_reader.cpp    # 读取器示例
│   ├── example_writer.cpp    # 写入器示例（可添加）
│   └── ...                   # 其他示例
│
├── tests/                     # 测试目录
│   ├── CMakeLists.txt        # 测试总配置
│   ├── unit/                 # 单元测试
│   │   ├── CMakeLists.txt    # 单元测试配置
│   │   └── ...               # 单元测试文件
│   ├── integration/          # 集成测试
│   │   ├── CMakeLists.txt    # 集成测试配置
│   │   ├── test_reader_functionality.cpp  # 读取器功能测试
│   │   └── ...               # 其他集成测试
│   └── performance/          # 性能测试
│       ├── CMakeLists.txt    # 性能测试配置
│       └── ...               # 性能测试文件
│
├── third_party/              # 第三方依赖目录
│   ├── libxlsxwriter/        # libxlsxwriter库
│   ├── minizip-ng/           # minizip-ng库
│   ├── expat/                # expat XML解析器
│   └── zlib/                 # zlib压缩库
│
├── cmake/                    # CMake模块和脚本
├── build/                    # 构建输出目录（可选）
└── cmake-build-debug/        # IDE构建目录（可选）
```

## 构建指南

### 1. 构建整个项目
```bash
mkdir build
cd build
cmake ..
cmake --build .
```

### 2. 构建示例程序
```bash
# 构建所有示例
cmake --build . --target example_reader

# 运行示例
./examples/example_reader
```

### 3. 构建和运行测试
```bash
# 构建集成测试
cmake --build . --target test_reader_functionality

# 运行集成测试
./tests/integration/test_reader_functionality

# 或使用CTest运行所有测试
ctest
```

## 项目特点

### 清晰的分离
- **源代码** (`src/`): 库的实现代码
- **头文件** (`include/`): 公共API定义
- **示例** (`examples/`): 展示如何使用库的示例代码
- **测试** (`tests/`): 各种类型的测试代码
- **第三方依赖** (`third_party/`): 外部库和依赖

### 构建系统
- 每个目录都有自己的 `CMakeLists.txt`
- 模块化的构建配置
- 支持单独构建示例和测试
- 支持不同的构建类型（Debug/Release）

### 测试组织
- **单元测试** (`tests/unit/`): 测试单个功能模块
- **集成测试** (`tests/integration/`): 测试模块间的集成
- **性能测试** (`tests/performance/`): 性能基准测试

## 最佳实践

### 添加新的示例
1. 在 `examples/` 目录创建新的 `.cpp` 文件
2. 在 `examples/CMakeLists.txt` 中添加：
   ```cmake
   add_example(new_example_name new_example.cpp)
   ```

### 添加新的测试
1. 在适当的测试子目录创建测试文件
2. 测试文件命名应以 `test_` 开头
3. CMake会自动发现并构建新的测试

### 添加新的库功能
1. 在 `include/TinaXlsx/` 添加头文件
2. 在 `src/` 添加对应的实现文件
3. 在主 `CMakeLists.txt` 中添加到源文件列表

## 优势

### 1. 项目组织更清晰
- 易于导航和理解
- 符合行业标准
- 便于团队协作

### 2. 构建更灵活
- 可以单独构建库、示例或测试
- 支持增量构建
- 便于CI/CD集成

### 3. 维护更容易
- 模块化设计
- 清晰的依赖关系
- 便于添加新功能

### 4. 适合开发流程
- 示例代码帮助用户快速上手
- 分层的测试策略确保代码质量
- 清晰的API设计

这种结构符合现代C++项目的最佳实践，使项目更加专业和易于维护。 