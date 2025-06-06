cmake_minimum_required(VERSION 3.16)
project(TinaXlsx LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF) 

# 设置项目版本
set(TINAXLSX_VERSION_MAJOR 2)
set(TINAXLSX_VERSION_MINOR 1)
set(TINAXLSX_VERSION_PATCH 0)
set(TINAXLSX_VERSION "${TINAXLSX_VERSION_MAJOR}.${TINAXLSX_VERSION_MINOR}.${TINAXLSX_VERSION_PATCH}")


# 添加cmake模块路径
list(APPEND CMAKE_MODULE_PATH ${CMAKE_CURRENT_SOURCE_DIR}/cmake)

# 加载编译器配置（包含UTF-8编码设置）
include(CompilerConfig)

# 包含测试工具函数
include(TestUtils)

# ------------ 项目配置选项 ------------
option(TINAXLSX_BUILD_TESTS "Build the TinaXlsx unit tests" ON)
option(TINAXLSX_BUILD_DOCS "Build TinaXlsx API documentation with Doxygen" OFF)


# -------------- 配置 pugixml --------------
add_subdirectory(third_party/pugixml)
# pugixml 将提供 pugixml::pugixml 目标供链接使用

# -------------- 配置 zlib-ng --------------
# zlib-ng 的构建选项
set(ZLIB_COMPAT ON CACHE BOOL "" FORCE)  # 启用zlib兼容性
set(ZLIB_ENABLE_TESTS OFF CACHE BOOL "" FORCE)
set(WITH_GTEST OFF CACHE BOOL "" FORCE)

add_subdirectory(third_party/zlib-ng)
# zlib-ng 提供 zlib 目标

# -------------- 配置 minizip-ng --------------
# 确保 minizip-ng 能找到 zlib
find_package(ZLIB QUIET)
if(NOT ZLIB_FOUND)
    # 手动设置 zlib 变量指向 zlib-ng
    set(ZLIB_LIBRARY zlibstatic)
    set(ZLIB_INCLUDE_DIR ${CMAKE_CURRENT_SOURCE_DIR}/third_party/zlib-ng)
    set(ZLIB_LIBRARIES ${ZLIB_LIBRARY})
    set(ZLIB_FOUND TRUE)
endif()

# minizip-ng 的构建选项
set(MZ_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_BZIP2 OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_LZMA OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_ZSTD OFF CACHE BOOL "" FORCE)
set(MZ_COMPAT OFF CACHE BOOL "" FORCE)
set(MZ_FETCH_LIBS OFF CACHE BOOL "" FORCE)
set(MZ_FORCE_FETCH_LIBS OFF CACHE BOOL "" FORCE)
# 启用 zlib 支持
set(MZ_ZLIB ON CACHE BOOL "" FORCE)

add_subdirectory(third_party/minizip-ng)
# minizip 提供 minizip 目标

# -------------- 配置 fast_float --------------
# fast_float 是 header-only 库
add_subdirectory(third_party/fast_float)

# -------------- 配置 xsimd --------------
# xsimd 是 header-only SIMD 库，禁用其测试
set(BUILD_TESTS OFF CACHE BOOL "" FORCE)
add_subdirectory(third_party/xsimd)

# -------------- 配置 fmt --------------
# fmt 是高性能格式化库
set(FMT_INSTALL OFF CACHE BOOL "" FORCE)
set(FMT_TEST OFF CACHE BOOL "" FORCE)
add_subdirectory(third_party/fmt)

# -----------------------------------------

# TinaXlsx 核心源文件 - 极简内存优先架构
set(CORE_SOURCES
    # 🎯 核心内存优先架构 (必需)
    src/TXBatchSIMDProcessor.cpp      # SIMD批量处理
    src/TXInMemorySheet.cpp           # 内存优先工作表
    src/TXZeroCopySerializer.cpp      # 零拷贝序列化
    
    # 📦 基础支撑组件 (必需)
    src/TXCoordinate.cpp              # 坐标系统
    src/TXTypes.cpp                   # 基础类型
    src/TXError.cpp                   # 错误处理
    src/TXGlobalStringPool.cpp        # 字符串池
    src/TXUnifiedMemoryManager.cpp    # 统一内存管理
    src/TXSlabAllocator.cpp           # Slab分配器
    src/TXChunkAllocator.cpp          # Chunk分配器
    src/TXSmartMemoryManager.cpp      # 智能内存管理器
    
    # 🔧 工具库 (最小化)
    src/TXNumberUtils.cpp             # 数字工具
    src/TXRange.cpp                   # 范围操作
    src/TinaXlsx.cpp                  # 主入口
    
    # 🎨 可选样式支持 (简化)
    src/TXColor.cpp                   # 颜色处理
    src/TXStyle.cpp                   # 基础样式
)


# 创建 TinaXlsx 静态库
add_library(${PROJECT_NAME} STATIC ${CORE_SOURCES})

# 编译器特定选项已在CompilerConfig.cmake中配置

# 链接核心依赖库
target_link_libraries(${PROJECT_NAME}
        PUBLIC
        pugixml::pugixml
        minizip-ng
        fast_float
        xsimd
        fmt::fmt
)

target_include_directories(${PROJECT_NAME} PUBLIC
        ${PROJECT_SOURCE_DIR}/include
)

# ------------ GoogleTest 配置 ------------
if(TINAXLSX_BUILD_TESTS)
    enable_testing()

    # 配置 GoogleTest
    set(BUILD_GMOCK OFF CACHE BOOL "" FORCE)
    set(INSTALL_GTEST OFF CACHE BOOL "" FORCE)
    add_subdirectory(third_party/googletest)
    
    # 包含 CTest 模块
    include(CTest)
    
    # 创建一个运行所有测试的目标
    add_custom_target(run_all_tests
        COMMAND ${CMAKE_CTEST_COMMAND} --output-on-failure --verbose
        COMMENT "Running all tests"
        WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
    )
    
    # 包含测试配置
    add_subdirectory(tests)

endif()

# ------------ 文档生成配置 ------------
if(TINAXLSX_BUILD_DOCS)
    message(STATUS "启用文档生成")
    include(DoxygenConfig)
else()
    message(STATUS "文档生成已禁用，使用 -DTINAXLSX_BUILD_DOCS=ON 启用")
endif()
