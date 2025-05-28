cmake_minimum_required(VERSION 3.16)
project(TinaXlsx VERSION 2.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# ------------ 项目配置选项 ------------
option(BUILD_TESTS "Build the unit tests" ON)


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

# -----------------------------------------

# 搜索 TinaXlsx 项目源文件
file(GLOB_RECURSE CORE_SOURCES CONFIGURE_DEPENDS ${PROJECT_SOURCE_DIR}/src/*.cpp)


# 创建 TinaXlsx 静态库
add_library(${PROJECT_NAME} STATIC ${CORE_SOURCES})

# 设置编译器特定选项
if(MSVC)
    target_compile_options(${PROJECT_NAME} PRIVATE 
        /W4 
        /utf-8  # 设置源和执行字符集为UTF-8
        /wd4819  # 禁用C4819编码警告
    )
else()
    target_compile_options(${PROJECT_NAME} PRIVATE -Wall -Wextra -Wpedantic)
endif()

# 链接 pugixml 和 minizip 库到项目
target_link_libraries(${PROJECT_NAME}
        PUBLIC
        pugixml::pugixml
        minizip-ng
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

    set(BUILD_GMOCK OFF CACHE BOOL "" FORCE)
    set(INSTALL_GTEST OFF CACHE BOOL "" FORCE)
    add_subdirectory(third_party/googletest)

    file(GLOB_RECURSE TEST_SOURCES CONFIGURE_DEPENDS ${PROJECT_SOURCE_DIR}/tests/*.cpp)

    add_executable(${PROJECT_NAME}_test ${TEST_SOURCES})

    target_link_libraries(${PROJECT_NAME}_test
            PRIVATE
            ${PROJECT_NAME}
            gtest_main
    )

    # 为测试添加更宽松的编译选项
    if(MSVC)
        target_compile_options(${PROJECT_NAME}_test PRIVATE 
            /W3 
            /utf-8  # 设置源和执行字符集为UTF-8
            /wd4819  # 禁用C4819编码警告
        )
    else()
        target_compile_options(${PROJECT_NAME}_test PRIVATE -Wall -Wno-unused-parameter)
    endif()

    include(GoogleTest)
    gtest_discover_tests(${PROJECT_NAME}_test)
endif()

