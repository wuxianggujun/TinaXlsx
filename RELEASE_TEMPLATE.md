# TinaXlsx Release Template

## 📦 下载说明

### 快速选择
- **Windows 开发者**: 下载 `TinaXlsx-Windows-x64.zip`
- **Linux 开发者**: 下载 `TinaXlsx-Linux-x64.tar.gz`
- **macOS 开发者**: 下载 `TinaXlsx-macOS-x64.tar.gz`

### 包内容
每个发布包都包含：
- 📚 **完整的静态库文件** - 无需额外依赖
- 📄 **头文件** - 完整的 C++ API
- 📖 **文档** - README 和 LICENSE
- ⚙️ **配置文件** - CMake 和 pkg-config 支持

## 🚀 快速开始

### 方法1：CMake 集成
```cmake
# 1. 解压下载的包到项目目录
# 2. 在 CMakeLists.txt 中添加：

find_package(TinaXlsx REQUIRED)
target_link_libraries(your_target TinaXlsx::TinaXlsx)
```

### 方法2：手动链接
```cmake
# 添加头文件路径
include_directories(path/to/tinaxlsx/include)

# 链接所有库文件
if(WIN32)
    target_link_libraries(your_target 
        TinaXlsx.lib
        xlsxwriter.lib 
        xlsxio_read_STATIC.lib
        xlsxio_write_STATIC.lib
        zlibstatic.lib
        expat.lib
    )
else()
    target_link_libraries(your_target
        libTinaXlsx.a
        libxlsxwriter.a
        libxlsxio_read_STATIC.a
        libxlsxio_write_STATIC.a
        libz.a
        libexpat.a
    )
endif()
```

### 方法3：pkg-config (Linux/macOS)
```bash
# 设置 PKG_CONFIG_PATH
export PKG_CONFIG_PATH=/path/to/tinaxlsx/lib/pkgconfig:$PKG_CONFIG_PATH

# 编译时使用
g++ -o your_program your_code.cpp $(pkg-config --libs --cflags tinaxlsx)
```

## 💡 使用示例

```cpp
#include <TinaXlsx/TinaXlsx.hpp>

int main() {
    TinaXlsx::Workbook workbook("example.xlsx");
    auto worksheet = workbook.addWorksheet("Sheet1");
    
    worksheet->writeString({0, 0}, "Hello");
    worksheet->writeString({0, 1}, "World");
    worksheet->writeNumber({1, 0}, 42);
    
    workbook.save();
    return 0;
}
```

## 🔧 系统要求

- **C++17** 或更高版本
- **CMake 3.16+** (推荐)
- 支持的平台：
  - Windows 10+ (Visual Studio 2019+)
  - Linux (GCC 8+, Clang 8+)
  - macOS 10.15+ (Xcode 11+)

## 📊 性能特性

- ✅ **高性能类型转换** - 自定义字符串转换框架
- ✅ **类型安全** - 强类型枚举替代硬编码常量
- ✅ **内存优化** - 栈上缓冲区避免堆分配
- ✅ **跨平台兼容** - 统一的类型系统
- ✅ **零依赖** - 所有依赖库已静态链接

## 🐛 问题反馈

如果遇到问题，请：
1. 查看 [FAQ](https://github.com/wuxianggujun/TinaXlsx/wiki/FAQ)
2. 搜索 [已知问题](https://github.com/wuxianggujun/TinaXlsx/issues)
3. 提交新的 [Issue](https://github.com/wuxianggujun/TinaXlsx/issues/new)

## 📝 更新日志

详细的更新日志请查看 [CHANGELOG.md](https://github.com/wuxianggujun/TinaXlsx/blob/main/CHANGELOG.md)

---

**感谢使用 TinaXlsx！如果对您有帮助，请给我们一个 ⭐** 