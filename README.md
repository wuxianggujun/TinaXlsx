# TinaXlsx 项目设计文档

## 一、项目简介

**TinaXlsx** 是一个使用 C++ 开发的 Excel 文件（XLSX 格式）解析与生成库，基于现代的 C++17 标准，提供高效、灵活的 Excel 文档处理功能。

## 二、项目结构与类关系

### 核心类与功能说明：

* **TXWorkbook**：

    * 管理整个Excel工作簿的生命周期，包括打开、保存、Sheet管理。

* **TXSheet**：

    * 表示Excel中的一个Sheet，负责管理Sheet内的数据操作。

* **TXCell**：

    * 单元格数据的抽象，提供读写单个单元格的能力。

* **TXReader / TXWriter**：

    * 分别负责XLSX文件的读取与写入。

* **TXZipHandler**：

    * 使用minizip-ng库处理XLSX文件的压缩解压。

* **TXUtils**：

    * 通用辅助函数集合。

### 类关系图示

```
TXWorkbook
   ├─ manages ─ TXSheet
   │                └─ contains ─ TXCell
   ├─ uses ─ TXReader
   └─ uses ─ TXWriter
             └─ depends on ─ TXZipHandler
                               └─ depends on ─ minizip-ng
```

## 三、项目要求与依赖库

### 项目基础要求：

* 编译器要求：支持 C++17 的编译器（如GCC 8+, MSVC 2019+）
* 构建系统：CMake 3.16及以上

### 依赖库说明：

| 库名称            | 用途           | 获取地址                                                                           |
| -------------- | ------------ | ------------------------------------------------------------------------------ |
| **minizip-ng** | ZIP文件的压缩解压操作 | [https://github.com/zlib-ng/minizip-ng](https://github.com/zlib-ng/minizip-ng) |
| **pugixml**    | XML文件解析与生成   | [https://github.com/zeux/pugixml](https://github.com/zeux/pugixml)             |
| **googletest** | 单元测试框架       | [https://github.com/google/googletest](https://github.com/google/googletest)   |

## 四、项目构建与测试

### 构建步骤：

```shell
mkdir build && cd build
cmake .. -DBUILD_TESTS=ON
cmake --build . --config Release
```

### 测试运行：

```shell
cd build
ctest --verbose
```

## 五、开发规范与要求

* 创建一个新的类时，必须同时创建对应的单元测试类，确保测试覆盖。
* 严格使用现代C++标准，推荐使用智能指针（`std::unique_ptr` 和 `std::shared_ptr`）代替传统指针。

## 六、扩展与维护

本项目采用模块化设计，新增功能只需添加对应的类即可，建议所有新类名以`TX`为前缀，以保持项目结构一致性。

---
