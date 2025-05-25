cmake_minimum_required(VERSION 3.16)
project(TinaXlsx VERSION 1.0 LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# ------------ 项目配置选项 ------------
option(BUILD_TESTS "Build the unit tests" ON)


# -------------- 配置 pugixml --------------
add_subdirectory(third_party/pugixml)
# pugixml 将提供 pugixml::pugixml 目标供链接使用

# -------------- 配置 minizip-ng --------------
# minizip-ng 的构建选项
set(MZ_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_BZIP2 OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_LZMA OFF CACHE BOOL "" FORCE)
set(MZ_BUILD_ZSTD OFF CACHE BOOL "" FORCE)
set(MZ_COMPAT OFF CACHE BOOL "" FORCE)
set(MZ_FETCH_LIBS OFF CACHE BOOL "" FORCE)
set(MZ_FORCE_FETCH_LIBS OFF CACHE BOOL "" FORCE)

add_subdirectory(third_party/minizip-ng)
# minizip 提供 minizip 目标

# -----------------------------------------

# 搜索 TinaXlsx 项目源文件
file(GLOB_RECURSE SOURCES CONFIGURE_DEPENDS ${PROJECT_SOURCE_DIR}/src/*.cpp)

# 创建 TinaXlsx 静态库
add_library(${PROJECT_NAME} STATIC ${SOURCES})

# 链接 pugixml 和 minizip 库到项目
target_link_libraries(${PROJECT_NAME}
        PUBLIC
        pugixml::pugixml
        minizip-ng
)

target_include_directories(${PROJECT_NAME} PUBLIC
        ${PROJECT_SOURCE_DIR}/include
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

    include(GoogleTest)
    gtest_discover_tests(${PROJECT_NAME}_test)
endif()
