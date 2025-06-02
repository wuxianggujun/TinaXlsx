# TinaXlsx API 文档索引

本文档由CMake自动生成。

## 📚 文档类型

### 自动生成文档
- **HTML文档**: `api-docs/html/index.html`
- **XML文档**: `api-docs/xml/` (用于其他工具)

### 手动维护文档
- **API参考**: [`API_Reference.md`](./API_Reference.md)
- **使用指南**: [`../doc/TinaXlsx使用指南.md`](../doc/TinaXlsx使用指南.md)

## 🔧 生成文档

```bash
# 生成API文档
cmake --build cmake-build-debug --target docs

# 清理文档
cmake --build cmake-build-debug --target docs-clean

# 启动文档服务器
cmake --build cmake-build-debug --target docs-serve
```

## 📊 项目信息

- **项目名称**: TinaXlsx
- **版本**: 2.1
- **生成时间**: $(date)
- **CMake版本**: 3.31.1
