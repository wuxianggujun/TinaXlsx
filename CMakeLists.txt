cmake_minimum_required(VERSION 3.16)

project(TinaXlsx 
    VERSION 1.0.0
    DESCRIPTION "高性能Excel读写库，基于libxlsxwriter封装"
    LANGUAGES C CXX)

# 设置C++标准
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# 构建选项
option(TINAXLSX_BUILD_SHARED "构建动态库" OFF)
option(TINAXLSX_BUILD_EXAMPLES "构建示例程序" OFF)
option(TINAXLSX_BUILD_TESTS "构建测试程序" ON)
option(TINAXLSX_ENABLE_PERFORMANCE "启用性能优化" ON)

# =============================================================================
# 配置第三方库 - 不修改任何第三方库的默认配置
# =============================================================================

# 禁用第三方库不需要的构建内容
set(BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
set(BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(BUILD_TOOLS OFF CACHE BOOL "" FORCE)
set(BUILD_DOCS OFF CACHE BOOL "" FORCE)

# 1. 构建 zlib
message(STATUS "配置 zlib...")
add_subdirectory(third_party/zlib)

# 2. 构建 expat (XML解析库，xlsxio需要)
message(STATUS "配置 expat...")
# 设置 expat 构建为静态库，避免 DLL 依赖问题
set(EXPAT_BUILD_SHARED OFF CACHE BOOL "Build expat shared library" FORCE)
set(EXPAT_SHARED_LIBS OFF CACHE BOOL "Build expat shared library" FORCE)
add_subdirectory(third_party/expat/expat)

# 3. 为 minizip 创建一个简单的包装器
# 使用 libxlsxwriter 内置的 minizip 源文件，创建独立库目标
message(STATUS "配置 minizip...")
set(MINIZIP_SOURCES
    third_party/libxlsxwriter/third_party/minizip/ioapi.c
    third_party/libxlsxwriter/third_party/minizip/zip.c
    third_party/libxlsxwriter/third_party/minizip/unzip.c
    third_party/libxlsxwriter/third_party/minizip/mztools.c
)
if(WIN32)
    list(APPEND MINIZIP_SOURCES third_party/libxlsxwriter/third_party/minizip/iowin32.c)
endif()

add_library(minizip STATIC ${MINIZIP_SOURCES})
target_include_directories(minizip 
    PUBLIC 
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/third_party/libxlsxwriter/third_party/minizip>
        $<INSTALL_INTERFACE:include/minizip>
)
target_link_libraries(minizip PUBLIC ZLIB::ZLIB)
target_compile_definitions(minizip PRIVATE NOCRYPT NOUNCRYPT)
if(WIN32)
    target_compile_definitions(minizip PRIVATE USE_FILE32API)
endif()

# 4. 构建 libxlsxwriter，使用内置 minizip（避免依赖查找问题）
message(STATUS "配置 libxlsxwriter...")
# 让 libxlsxwriter 使用内置的 minizip，避免复杂的依赖查找
set(USE_SYSTEM_MINIZIP OFF CACHE BOOL "" FORCE)

# 在 libxlsxwriter 配置前设置 ZLIB 变量，模拟 find_package 的结果
set(ZLIB_FOUND TRUE CACHE BOOL "ZLIB found" FORCE)
set(ZLIB_INCLUDE_DIRS "${CMAKE_CURRENT_SOURCE_DIR}/third_party/zlib;${CMAKE_CURRENT_BINARY_DIR}/third_party/zlib" CACHE STRING "ZLIB include dirs" FORCE)
set(ZLIB_INCLUDE_DIR "${CMAKE_CURRENT_SOURCE_DIR}/third_party/zlib" CACHE PATH "ZLIB include dir" FORCE)
set(ZLIB_LIBRARIES zlibstatic CACHE STRING "ZLIB libraries" FORCE)
set(ZLIB_LIBRARY zlibstatic CACHE STRING "ZLIB library" FORCE)
set(ZLIB_VERSION_STRING "1.3.1" CACHE STRING "ZLIB version" FORCE)

# 创建 ZLIB::ZLIB 目标别名，让 libxlsxwriter 能找到
add_library(ZLIB::ZLIB ALIAS zlibstatic)

add_subdirectory(third_party/libxlsxwriter)

# 5. 构建 xlsxio
message(STATUS "配置 xlsxio...")

# 设置 xlsxio 构建选项 - 只构建静态库，避免文件名冲突
set(BUILD_STATIC ON CACHE BOOL "Build xlsxio static libraries" FORCE)
set(BUILD_SHARED OFF CACHE BOOL "Build xlsxio shared libraries" FORCE)
set(BUILD_TOOLS OFF CACHE BOOL "Build xlsxio tools" FORCE)
set(BUILD_EXAMPLES OFF CACHE BOOL "Build xlsxio examples" FORCE)

# 为 xlsxio 设置 Minizip 变量
set(MINIZIP_FOUND TRUE CACHE BOOL "Minizip found" FORCE)
set(MINIZIP_INCLUDE_DIRS "${CMAKE_CURRENT_SOURCE_DIR}/third_party/libxlsxwriter/third_party" CACHE STRING "Minizip include dirs" FORCE)
set(MINIZIP_LIBRARIES minizip CACHE STRING "Minizip libraries" FORCE)

# 为 xlsxio 设置 EXPAT 变量
set(EXPAT_FOUND TRUE CACHE BOOL "EXPAT found" FORCE)
set(EXPAT_INCLUDE_DIR "${CMAKE_CURRENT_SOURCE_DIR}/third_party/expat/expat/lib" CACHE PATH "EXPAT include dir" FORCE)
set(EXPAT_LIBRARY expat CACHE STRING "EXPAT library" FORCE)
set(EXPAT_LIBRARIES expat CACHE STRING "EXPAT libraries" FORCE)

# 禁用 xlsxio 的导出安装，避免依赖问题
set(CMAKE_SKIP_INSTALL_RULES ON)

add_subdirectory(third_party/xlsxio)

# 恢复安装规则
set(CMAKE_SKIP_INSTALL_RULES OFF)

# 6. Google Test (仅在启用测试时)
if(TINAXLSX_BUILD_TESTS)
    message(STATUS "配置 Google Test...")
    set(INSTALL_GTEST OFF CACHE BOOL "" FORCE)
    set(BUILD_GMOCK ON CACHE BOOL "" FORCE)
    set(gtest_force_shared_crt ON CACHE BOOL "" FORCE)
    add_subdirectory(third_party/googletest)
    enable_testing()
endif()

# =============================================================================
# 主库配置
# =============================================================================

# 收集源文件
file(GLOB_RECURSE TINAXLSX_SOURCES "src/*.cpp")
file(GLOB_RECURSE TINAXLSX_HEADERS "include/*.hpp")

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
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/xlsxio/include
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/libxlsxwriter/include
)

# 链接库 - 使用PRIVATE确保依赖不会传播，并通过库文件直接链接
target_link_libraries(TinaXlsx
    PRIVATE
        $<TARGET_FILE:xlsxwriter>
        $<TARGET_FILE:xlsxio_read_STATIC>
        $<TARGET_FILE:xlsxio_write_STATIC>
        $<TARGET_FILE:zlibstatic>
        $<TARGET_FILE:expat>
        $<TARGET_FILE:minizip>
)

# Windows特定设置
if(WIN32)
    target_link_libraries(TinaXlsx PRIVATE ole32 oleaut32 ws2_32)
    
    if(TINAXLSX_BUILD_SHARED)
        target_compile_definitions(TinaXlsx
            PRIVATE TINAXLSX_EXPORTS
            INTERFACE TINAXLSX_SHARED
        )
    endif()
endif()

# 编译器特定设置
if(MSVC)
    target_compile_options(TinaXlsx PRIVATE /W3 /bigobj /utf-8)
    
    if(TINAXLSX_ENABLE_PERFORMANCE AND CMAKE_BUILD_TYPE STREQUAL "Release")
        target_compile_options(TinaXlsx PRIVATE /O2 /Oi /GL)
        if(TINAXLSX_BUILD_SHARED)
            target_link_options(TinaXlsx PRIVATE /LTCG)
        endif()
    endif()
    
    target_compile_definitions(TinaXlsx PRIVATE
        _CRT_SECURE_NO_WARNINGS
        _SCL_SECURE_NO_WARNINGS
        NOMINMAX
    )
endif()

# 调试/发布配置
if(CMAKE_BUILD_TYPE STREQUAL "Debug")
    target_compile_definitions(TinaXlsx PRIVATE DEBUG TINAXLSX_DEBUG)
else()
    target_compile_definitions(TinaXlsx PRIVATE NDEBUG TINAXLSX_RELEASE)
endif()

# xlsxio 静态库链接定义
target_compile_definitions(TinaXlsx PRIVATE 
    BUILD_XLSXIO_STATIC
    STATIC
)

# =============================================================================
# 测试配置
# =============================================================================

if(TINAXLSX_BUILD_TESTS)
    add_subdirectory(tests)
endif()

# =============================================================================
# 安装配置
# =============================================================================

include(GNUInstallDirs)
include(CMakePackageConfigHelpers)

# 创建别名目标（供内部和外部使用）
add_library(TinaXlsx::TinaXlsx ALIAS TinaXlsx)

# 安装目标 - 只安装TinaXlsx，依赖库已静态链接
install(TARGETS TinaXlsx
    EXPORT TinaXlsxTargets
    LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
    ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
    RUNTIME DESTINATION ${CMAKE_INSTALL_BINDIR}
    INCLUDES DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
)

# 注意：第三方库不需要单独安装，因为它们已经静态链接到TinaXlsx中

# 安装头文件
install(DIRECTORY include/
    DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
    FILES_MATCHING PATTERN "*.hpp"
)

# 安装文档
install(FILES README.md LICENSE
    DESTINATION ${CMAKE_INSTALL_DOCDIR}
)

# 导出目标配置
install(EXPORT TinaXlsxTargets
    FILE TinaXlsxTargets.cmake
    NAMESPACE TinaXlsx::
    DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/TinaXlsx
)

# 创建配置文件
configure_package_config_file(
    "${CMAKE_CURRENT_SOURCE_DIR}/cmake/TinaXlsxConfig.cmake.in"
    "${CMAKE_CURRENT_BINARY_DIR}/TinaXlsxConfig.cmake"
    INSTALL_DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/TinaXlsx
)

# 创建版本文件
write_basic_package_version_file(
    "${CMAKE_CURRENT_BINARY_DIR}/TinaXlsxConfigVersion.cmake"
    VERSION ${PROJECT_VERSION}
    COMPATIBILITY SameMajorVersion
)

# 安装配置文件
install(FILES
    "${CMAKE_CURRENT_BINARY_DIR}/TinaXlsxConfig.cmake"
    "${CMAKE_CURRENT_BINARY_DIR}/TinaXlsxConfigVersion.cmake"
    DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/TinaXlsx
)

# =============================================================================
# 打包配置 (CPack)
# =============================================================================

set(CPACK_PACKAGE_NAME "TinaXlsx")
set(CPACK_PACKAGE_VENDOR "TinaXlsx Project")
set(CPACK_PACKAGE_DESCRIPTION_SUMMARY "高性能Excel读写库，基于libxlsxwriter封装")
set(CPACK_PACKAGE_VERSION ${PROJECT_VERSION})
set(CPACK_PACKAGE_VERSION_MAJOR ${PROJECT_VERSION_MAJOR})
set(CPACK_PACKAGE_VERSION_MINOR ${PROJECT_VERSION_MINOR})
set(CPACK_PACKAGE_VERSION_PATCH ${PROJECT_VERSION_PATCH})
set(CPACK_PACKAGE_INSTALL_DIRECTORY "TinaXlsx")
set(CPACK_RESOURCE_FILE_LICENSE "${CMAKE_CURRENT_SOURCE_DIR}/LICENSE")
set(CPACK_RESOURCE_FILE_README "${CMAKE_CURRENT_SOURCE_DIR}/README.md")

# 平台特定的包配置
if(WIN32)
    set(CPACK_GENERATOR "ZIP;NSIS")
    set(CPACK_NSIS_DISPLAY_NAME "TinaXlsx ${PROJECT_VERSION}")
    set(CPACK_NSIS_PACKAGE_NAME "TinaXlsx")
    set(CPACK_NSIS_URL_INFO_ABOUT "https://github.com/wuxianggujun/TinaXlsx")
    set(CPACK_NSIS_MODIFY_PATH ON)
elseif(APPLE)
    set(CPACK_GENERATOR "TGZ;DragNDrop")
    set(CPACK_DMG_FORMAT "UDBZ")
    set(CPACK_DMG_VOLUME_NAME "TinaXlsx")
else()
    set(CPACK_GENERATOR "TGZ;DEB;RPM")
    set(CPACK_DEBIAN_PACKAGE_MAINTAINER "TinaXlsx Project")
    set(CPACK_DEBIAN_PACKAGE_SECTION "libs")
    set(CPACK_DEBIAN_PACKAGE_DEPENDS "libc6, libstdc++6")
    set(CPACK_RPM_PACKAGE_GROUP "Development/Libraries")
    set(CPACK_RPM_PACKAGE_LICENSE "MIT")
endif()

include(CPack)

# =============================================================================
# 开发环境配置
# =============================================================================

# 导出编译命令（用于IDE和代码分析工具）
set(CMAKE_EXPORT_COMPILE_COMMANDS ON)

# 创建符号链接到build目录（方便IDE）
if(CMAKE_BUILD_TYPE STREQUAL "Debug" AND UNIX)
    add_custom_target(dev-setup
        COMMAND ${CMAKE_COMMAND} -E create_symlink 
            ${CMAKE_CURRENT_SOURCE_DIR}/include 
            ${CMAKE_CURRENT_BINARY_DIR}/include
        COMMENT "Creating development symlinks"
    )
endif()

# =============================================================================
# 输出信息
# =============================================================================

message(STATUS "==================== TinaXlsx 配置信息 ====================")
message(STATUS "版本: ${PROJECT_VERSION}")
message(STATUS "构建类型: ${CMAKE_BUILD_TYPE}")
message(STATUS "库类型: ${TINAXLSX_BUILD_SHARED}")
message(STATUS "性能优化: ${TINAXLSX_ENABLE_PERFORMANCE}")
message(STATUS "构建测试: ${TINAXLSX_BUILD_TESTS}")
message(STATUS "安装前缀: ${CMAKE_INSTALL_PREFIX}")
message(STATUS "编译器: ${CMAKE_CXX_COMPILER_ID} ${CMAKE_CXX_COMPILER_VERSION}")
message(STATUS "C++标准: ${CMAKE_CXX_STANDARD}")
if(TINAXLSX_BUILD_SHARED)
    message(STATUS "构建类型: 动态库")
else()
    message(STATUS "构建类型: 静态库")
endif()
message(STATUS "========================================================") 
