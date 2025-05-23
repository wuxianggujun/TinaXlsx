cmake_minimum_required(VERSION 3.16)

project(TinaXlsx 
    VERSION 1.0.0
    DESCRIPTION "高性能Excel读写库，基于libxlsxwriter封装"
    LANGUAGES CXX)

# 设置C++标准
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 构建选项
option(TINAXLSX_BUILD_SHARED "构建动态库" OFF)
option(TINAXLSX_BUILD_EXAMPLES "构建示例程序" OFF)
option(TINAXLSX_BUILD_TESTS "构建测试程序" OFF)
option(TINAXLSX_ENABLE_PERFORMANCE "启用性能优化" ON)

# 配置第三方库
# 禁用第三方库的示例和测试
set(BUILD_EXAMPLES OFF CACHE BOOL "Build libxlsxwriter examples" FORCE)
set(BUILD_TESTS OFF CACHE BOOL "Build libxlsxwriter tests" FORCE)
set(BUILD_TOOLS OFF CACHE BOOL "Build xlsxio tools" FORCE)

# 构建zlib
message(STATUS "正在配置zlib...")
set(ZLIB_BUILD_EXAMPLES OFF CACHE BOOL "Build zlib examples" FORCE)
add_subdirectory(third_party/zlib)

# 设置ZLIB路径，让libxlsxwriter能找到
get_target_property(ZLIB_INCLUDE_DIRS zlib INTERFACE_INCLUDE_DIRECTORIES)
set(ZLIB_INCLUDE_DIR "${CMAKE_CURRENT_SOURCE_DIR}/third_party/zlib;${CMAKE_CURRENT_BINARY_DIR}/third_party/zlib" CACHE STRING "ZLIB include dir" FORCE)
set(ZLIB_LIBRARY zlib CACHE STRING "ZLIB library" FORCE)
set(ZLIB_LIBRARIES zlib CACHE STRING "ZLIB libraries" FORCE)
set(ZLIB_FOUND TRUE CACHE BOOL "ZLIB found" FORCE)

# 构建expat
message(STATUS "正在配置expat...")
set(EXPAT_BUILD_EXAMPLES OFF CACHE BOOL "Build expat examples" FORCE)
set(EXPAT_BUILD_TESTS OFF CACHE BOOL "Build expat tests" FORCE)
set(EXPAT_BUILD_TOOLS OFF CACHE BOOL "Build expat tools" FORCE)
set(EXPAT_BUILD_DOCS OFF CACHE BOOL "Build expat docs" FORCE)
add_subdirectory(third_party/expat/expat)

# 构建minizip-ng
message(STATUS "正在配置minizip-ng...")
set(MZ_BUILD_TESTS OFF CACHE BOOL "Build minizip-ng tests" FORCE)
set(MZ_BUILD_UNIT_TESTS OFF CACHE BOOL "Build minizip-ng unit tests" FORCE)
set(MZ_COMPAT ON CACHE BOOL "Enable minizip compatibility layer" FORCE)
add_subdirectory(third_party/minizip-ng)

# 构建libxlsxwriter
message(STATUS "正在配置libxlsxwriter...")
add_subdirectory(third_party/libxlsxwriter)

# 构建xlsxio
message(STATUS "正在配置xlsxio...")
set(BUILD_STATIC ON CACHE BOOL "Build xlsxio static libraries" FORCE)
set(BUILD_SHARED OFF CACHE BOOL "Build xlsxio shared libraries" FORCE)
set(WITH_LIBZIP OFF CACHE BOOL "Use libzip instead of minizip" FORCE)
set(WITH_MINIZIP_NG OFF CACHE BOOL "Use minizip-ng" FORCE)
# 设置依赖路径和库
set(EXPAT_INCLUDE_DIR "${CMAKE_CURRENT_SOURCE_DIR}/third_party/expat/expat/lib" CACHE STRING "EXPAT include dir" FORCE)
set(EXPAT_LIBRARY expat CACHE STRING "EXPAT library" FORCE)
set(EXPAT_LIBRARIES expat CACHE STRING "EXPAT libraries" FORCE)
set(EXPAT_FOUND TRUE CACHE BOOL "EXPAT found" FORCE)
# 设置minizip路径，使用minizip-ng的兼容层
set(MINIZIP_INCLUDE_DIRS "${CMAKE_CURRENT_SOURCE_DIR}/third_party/minizip-ng/compat" CACHE STRING "MINIZIP include dirs" FORCE)
set(MINIZIP_LIBRARIES minizip CACHE STRING "MINIZIP libraries" FORCE)
set(MINIZIP_FOUND TRUE CACHE BOOL "MINIZIP found" FORCE)
add_subdirectory(third_party/xlsxio)

# 查找依赖
find_package(PkgConfig QUIET)

# 第三方库已通过add_subdirectory构建，直接使用目标名
# libxlsxwriter目标名为：xlsxwriter
# xlsxio目标名为：xlsxio_read_STATIC, xlsxio_write_STATIC

message(STATUS "第三方库已配置完成")
message(STATUS "libxlsxwriter: 通过add_subdirectory构建")
message(STATUS "xlsxio: 通过add_subdirectory构建")

# 收集源文件
file(GLOB_RECURSE TINAXLSX_SOURCES
    "src/*.cpp"
)

file(GLOB_RECURSE TINAXLSX_HEADERS
    "include/*.hpp"
)

# 创建库
if(TINAXLSX_BUILD_SHARED)
    add_library(TinaXlsx SHARED ${TINAXLSX_SOURCES} ${TINAXLSX_HEADERS})
    message(STATUS "构建TinaXlsx动态库")
else()
    add_library(TinaXlsx STATIC ${TINAXLSX_SOURCES} ${TINAXLSX_HEADERS})
    message(STATUS "构建TinaXlsx静态库")
endif()

# 设置库属性
set_target_properties(TinaXlsx PROPERTIES
    VERSION ${PROJECT_VERSION}
    SOVERSION ${PROJECT_VERSION_MAJOR}
    CXX_STANDARD 17
    CXX_STANDARD_REQUIRED ON
    POSITION_INDEPENDENT_CODE ON
)

# 包含目录
target_include_directories(TinaXlsx
    PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
        $<INSTALL_INTERFACE:include>
    PRIVATE
        ${CMAKE_CURRENT_SOURCE_DIR}/src
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/libxlsxwriter/include
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/xlsxio/include
)

# 链接库
target_link_libraries(TinaXlsx
    PRIVATE
        xlsxwriter
        xlsxio_read_STATIC
        xlsxio_write_STATIC
)

# Windows特定设置
if(WIN32)
    target_link_libraries(TinaXlsx PRIVATE
        ole32
        oleaut32
        ws2_32
    )
    
    # 导出符号（动态库）
    if(TINAXLSX_BUILD_SHARED)
        target_compile_definitions(TinaXlsx
            PRIVATE TINAXLSX_EXPORTS
            INTERFACE TINAXLSX_SHARED
        )
    endif()
endif()

if(MSVC)
    target_compile_options(TinaXlsx PRIVATE
            /W3                    # 警告级别
            /bigobj               # 大对象文件
            /utf-8                # UTF-8编码
    )

    # 性能优化，仅在 Release 模式下启用
    if(TINAXLSX_ENABLE_PERFORMANCE AND CMAKE_BUILD_TYPE STREQUAL "Release")
        message(STATUS "TinaXlsx: 启用性能优化选项 (Release 模式)")
        target_compile_options(TinaXlsx PRIVATE
                /O2               # 最大速度优化
                /Oi               # 启用内联函数
                /GL               # 全程序优化
        )

        if(TINAXLSX_BUILD_SHARED)
            target_link_options(TinaXlsx PRIVATE /LTCG)  # 链接时代码生成
        endif()
    else()
        message(STATUS "TinaXlsx: 禁用性能优化选项 (Debug 模式或其他)")
        target_compile_options(TinaXlsx PRIVATE
                /Od               # 禁用优化，确保与 Debug 模式兼容
        )
    endif()

    # 禁用一些警告
    target_compile_definitions(TinaXlsx PRIVATE
            _CRT_SECURE_NO_WARNINGS
            _SCL_SECURE_NO_WARNINGS
            NOMINMAX
    )
endif()


# 调试信息
if(CMAKE_BUILD_TYPE STREQUAL "Debug")
    target_compile_definitions(TinaXlsx PRIVATE
        DEBUG
        TINAXLSX_DEBUG
    )
else()
    target_compile_definitions(TinaXlsx PRIVATE
        NDEBUG
        TINAXLSX_RELEASE
    )
endif()

# 安装配置
include(GNUInstallDirs)

# 安装库文件
install(TARGETS TinaXlsx
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
    ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
)

# 安装头文件
install(DIRECTORY include/
    DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
    FILES_MATCHING PATTERN "*.hpp"
)

# 导出配置 (暂时注释掉，避免依赖问题)
# install(EXPORT TinaXlsxTargets
#     FILE TinaXlsxTargets.cmake
#     NAMESPACE TinaXlsx::
#     DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/TinaXlsx
# )

# 创建配置文件
include(CMakePackageConfigHelpers)

configure_package_config_file(
    "${CMAKE_CURRENT_SOURCE_DIR}/cmake/TinaXlsxConfig.cmake.in"
    "${CMAKE_CURRENT_BINARY_DIR}/TinaXlsxConfig.cmake"
    INSTALL_DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/TinaXlsx
)

write_basic_package_version_file(
    "${CMAKE_CURRENT_BINARY_DIR}/TinaXlsxConfigVersion.cmake"
    VERSION ${PROJECT_VERSION}
    COMPATIBILITY SameMajorVersion
)

install(FILES
    "${CMAKE_CURRENT_BINARY_DIR}/TinaXlsxConfig.cmake"
    "${CMAKE_CURRENT_BINARY_DIR}/TinaXlsxConfigVersion.cmake"
    DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/TinaXlsx
)

# 创建pkgconfig文件
configure_file(
    "${CMAKE_CURRENT_SOURCE_DIR}/cmake/TinaXlsx.pc.in"
    "${CMAKE_CURRENT_BINARY_DIR}/TinaXlsx.pc"
    @ONLY
)

install(FILES
    "${CMAKE_CURRENT_BINARY_DIR}/TinaXlsx.pc"
    DESTINATION ${CMAKE_INSTALL_LIBDIR}/pkgconfig
)

# 输出配置信息
message(STATUS "==================== TinaXlsx 配置信息 ====================")
message(STATUS "版本: ${PROJECT_VERSION}")
message(STATUS "构建类型: ${CMAKE_BUILD_TYPE}")
message(STATUS "库类型: ${TINAXLSX_BUILD_SHARED}")
message(STATUS "性能优化: ${TINAXLSX_ENABLE_PERFORMANCE}")
message(STATUS "安装前缀: ${CMAKE_INSTALL_PREFIX}")
message(STATUS "=======================================================") 
