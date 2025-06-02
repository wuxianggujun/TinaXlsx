# TinaXlsx API 文档

本目录包含 TinaXlsx 项目的 API 文档，包括自动生成的文档和手动维护的参考文档。

## 📚 文档类型

### 🤖 自动生成文档

使用 Doxygen 从源代码注释自动生成的文档：

- **HTML文档**: [`html/index.html`](./html/index.html) - 主要的API文档
- **XML文档**: `xml/` - 用于其他工具处理的XML格式文档

### 📝 手动维护文档

- **API参考**: [`API_Reference.md`](./API_Reference.md) - 手动维护的API参考文档
- **类关系图**: 详细的类关系和功能说明

## 🔧 生成文档

### 使用CMake命令（推荐）

```bash
# 配置项目并启用文档生成
cmake -B cmake-build-debug -S . -DBUILD_DOCS=ON

# 生成API文档
cmake --build cmake-build-debug --target docs

# 清理生成的文档
cmake --build cmake-build-debug --target docs-clean

# 启动本地文档服务器（需要Python）
cmake --build cmake-build-debug --target docs-serve
```

### 一键生成和查看

```bash
# 完整流程：配置 → 生成 → 查看
cmake -B cmake-build-debug -S . -DBUILD_DOCS=ON && \
cmake --build cmake-build-debug --target docs && \
start api-docs/html/index.html
```

## 📋 环境要求

### 必需工具

- **CMake** 3.16+ - 构建系统
- **C++17编译器** - MSVC 2019+, GCC 8+, Clang 7+

### 自动包含的工具

- **Doxygen** - 项目内置，自动构建
- **所有依赖库** - 通过git子模块管理

### 可选工具

- **Python** 3.x - 用于启动本地文档服务器
- **Graphviz** - 用于生成更美观的类图（推荐安装）

### 快速安装Graphviz（可选）

#### Windows
```bash
# 从 https://graphviz.org/download/ 下载安装
# 或使用包管理器
winget install graphviz
```

#### Linux (Ubuntu/Debian)
```bash
sudo apt-get install graphviz
```

#### macOS
```bash
brew install graphviz
```

## 📖 文档内容

### 核心模块

- **TXWorkbook** - 工作簿管理
- **TXSheet** - 工作表操作
- **TXCell** - 单元格处理
- **TXStyle** - 样式管理
- **TXChart** - 图表功能
- **TXDataFilter** - 数据筛选
- **TXDataValidation** - 数据验证

### 工具类

- **TXNumberUtils** - 高性能数值处理
- **TXRange** - 范围操作
- **TXCoordinate** - 坐标转换
- **TXColor** - 颜色管理
- **TXFont** - 字体设置

### 管理器类

- **TXCellManager** - 单元格管理
- **TXStyleManager** - 样式管理
- **TXFormulaManager** - 公式管理
- **TXComponentManager** - 组件管理

## 🔍 文档导航

### 快速开始

1. **查看主页**: 打开 [`html/index.html`](./html/index.html)
2. **浏览类列表**: 点击 "Classes" 标签
3. **查看命名空间**: 点击 "Namespaces" 标签
4. **搜索功能**: 使用页面顶部的搜索框

### 常用页面

- **类层次结构**: 查看类的继承关系
- **文件列表**: 浏览源文件结构
- **函数索引**: 按字母顺序查看所有函数
- **数据结构**: 查看结构体和枚举

## 📊 文档统计

文档生成时会显示以下统计信息：
- 已文档化的类数量
- 已文档化的函数数量
- 文档覆盖率
- 警告和错误数量

## 🐛 问题报告

如果发现文档问题：

1. **内容错误**: 检查源代码注释是否正确
2. **生成失败**: 查看 Doxygen 输出的错误信息
3. **显示问题**: 检查浏览器兼容性

## 📝 贡献指南

### 改进文档注释

在源代码中使用 Doxygen 格式的注释：

```cpp
/**
 * @brief 简短描述
 * 
 * 详细描述...
 * 
 * @param param1 参数1描述
 * @param param2 参数2描述
 * @return 返回值描述
 * 
 * @example
 * @code
 * // 使用示例
 * TXWorkbook workbook;
 * @endcode
 * 
 * @see 相关函数或类
 * @since 版本信息
 */
void function(int param1, const std::string& param2);
```

### 更新手动文档

- 修改 [`API_Reference.md`](./API_Reference.md)
- 添加使用示例和最佳实践
- 更新版本变更信息

## 🔗 相关链接

- **项目主页**: [TinaXlsx](../)
- **用户文档**: [`../doc/`](../doc/)
- **开发文档**: [`../docs/`](../docs/)
- **问题跟踪**: [`../docs/KNOWN_ISSUES.md`](../docs/KNOWN_ISSUES.md)

---

**最后更新**: 2025年1月
**文档版本**: v1.1
**生成工具**: Doxygen + CMake
