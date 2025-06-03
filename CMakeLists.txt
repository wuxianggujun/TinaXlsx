cmake_minimum_required(VERSION 3.16)
project(TinaXlsx VERSION 2.1 LANGUAGES CXX)

# 添加cmake模块路径
list(APPEND CMAKE_MODULE_PATH ${CMAKE_CURRENT_SOURCE_DIR}/cmake)

# 加载编译器配置（包含UTF-8编码设置）
include(CompilerConfig)

# ------------ 项目配置选项 ------------
option(BUILD_TESTS "Build the unit tests" ON)
option(BUILD_DOCS "Build API documentation with Doxygen" OFF)


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

# -----------------------------------------

# 搜索 TinaXlsx 项目源文件
file(GLOB_RECURSE CORE_SOURCES CONFIGURE_DEPENDS ${PROJECT_SOURCE_DIR}/src/*.cpp)


# 创建 TinaXlsx 静态库
add_library(${PROJECT_NAME} STATIC ${CORE_SOURCES})

# 编译器特定选项已在CompilerConfig.cmake中配置

# 链接 pugixml、minizip 和 fast_float 库到项目
target_link_libraries(${PROJECT_NAME}
        PUBLIC
        pugixml::pugixml
        minizip-ng
        fast_float
)

target_include_directories(${PROJECT_NAME} PUBLIC
        ${PROJECT_SOURCE_DIR}/include
)

# 定义版本宏
target_compile_definitions(${PROJECT_NAME} PUBLIC
    TINAXLSX_VERSION_MAJOR=2
    TINAXLSX_VERSION_MINOR=0
    TINAXLSX_VERSION_PATCH=0
)

# ------------ GoogleTest 配置 ------------
if(BUILD_TESTS)
    enable_testing()

    # 配置 GoogleTest
    set(BUILD_GMOCK OFF CACHE BOOL "" FORCE)
    set(INSTALL_GTEST OFF CACHE BOOL "" FORCE)
    add_subdirectory(third_party/googletest)
    
    # 包含 CTest 模块
    include(CTest)
    
    # 添加单元测试子目录
    add_subdirectory(tests/unit)

    # 添加性能测试子目录
    add_subdirectory(tests/performance)

    # 创建一个运行所有测试的目标
    add_custom_target(run_all_tests
        COMMAND ${CMAKE_CTEST_COMMAND} --output-on-failure --verbose
        COMMENT "Running all tests"
        WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
    )

    # 创建一个专门运行性能测试的目标
    add_custom_target(run_performance_tests
        COMMAND SimplePerformanceTests --gtest_color=no
        COMMENT "Running simple performance tests"
        WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
        DEPENDS SimplePerformanceTests
    )

    # 创建一个运行极致性能测试的目标
    add_custom_target(run_extreme_performance_tests
        COMMAND ExtremePerformanceTests --gtest_color=no
        COMMENT "Running extreme performance tests"
        WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
        DEPENDS ExtremePerformanceTests
    )

    # 🚀 创建批量保存测试可执行文件（GoogleTest版本）
    add_executable(BatchSaveTests tests/batch_save_test.cpp)
    target_link_libraries(BatchSaveTests ${PROJECT_NAME} gtest gtest_main)

    # 创建运行批量保存测试的目标
    add_custom_target(run_batch_save_tests
        COMMAND BatchSaveTests --gtest_color=yes
        COMMENT "Running batch save tests (GoogleTest)"
        WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
        DEPENDS BatchSaveTests
    )

endif()

# ------------ 文档生成配置 ------------
if(BUILD_DOCS)
    message(STATUS "启用文档生成")
    include(DoxygenConfig)
else()
    message(STATUS "文档生成已禁用，使用 -DBUILD_DOCS=ON 启用")
endif()
