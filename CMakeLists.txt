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

# 2. 构建 expat (XML解析库，替代xlsxio需要)
message(STATUS "配置 expat...")
# 设置 expat 构建为静态库，避免 DLL 依赖问题
set(EXPAT_BUILD_SHARED OFF CACHE BOOL "Build expat shared library" FORCE)
set(EXPAT_SHARED_LIBS OFF CACHE BOOL "Build expat shared library" FORCE)
add_subdirectory(third_party/expat/expat)

# 3. 配置 minizip-ng (替代xlsxio的ZIP处理)
message(STATUS "配置 minizip-ng...")
set(MZ_COMPAT ON CACHE BOOL "Enable compatibility layer" FORCE)
set(MZ_ZLIB ON CACHE BOOL "Enable ZLIB compression" FORCE)
set(MZ_BZIP2 OFF CACHE BOOL "Enable BZIP2 compression" FORCE)
set(MZ_LZMA OFF CACHE BOOL "Enable LZMA compression" FORCE)
set(MZ_ZSTD OFF CACHE BOOL "Enable ZSTD compression" FORCE)
set(MZ_LIBCOMP OFF CACHE BOOL "Enable Apple compression" FORCE)
set(MZ_OPENSSL OFF CACHE BOOL "Enable OpenSSL" FORCE)
set(MZ_PKCRYPT ON CACHE BOOL "Enable PKWARE encryption" FORCE)
set(MZ_WZAES ON CACHE BOOL "Enable WinZIP AES encryption" FORCE)
set(MZ_BUILD_TESTS OFF CACHE BOOL "Build tests" FORCE)
set(MZ_BUILD_UNIT_TESTS OFF CACHE BOOL "Build unit tests" FORCE)

# 确保minizip-ng使用我们的静态zlib
set(ZLIB_LIBRARY zlibstatic CACHE STRING "ZLIB library" FORCE)
set(ZLIB_INCLUDE_DIR ${CMAKE_CURRENT_SOURCE_DIR}/third_party/zlib CACHE STRING "ZLIB include dir" FORCE)

add_subdirectory(third_party/minizip-ng)

# 4. 构建 libxlsxwriter (只用于写入功能)
message(STATUS "配置 libxlsxwriter...")
# 让 libxlsxwriter 使用内置的 minizip，避免复杂的依赖查找
set(USE_SYSTEM_MINIZIP OFF CACHE BOOL "" FORCE)

add_subdirectory(third_party/libxlsxwriter)

# 5. Google Test (仅在启用测试时)
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

# 包含目录 - 移除xlsxio，添加minizip-ng和expat
target_include_directories(TinaXlsx
    PUBLIC
        $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
        $<INSTALL_INTERFACE:include>
    PRIVATE
        ${CMAKE_CURRENT_SOURCE_DIR}/src
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/libxlsxwriter/include
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/minizip-ng
        ${CMAKE_CURRENT_SOURCE_DIR}/third_party/expat/expat/lib
)

# 链接库 - 移除xlsxio依赖，添加minizip-ng和expat，保留libxlsxwriter用于写入
target_link_libraries(TinaXlsx PRIVATE 
    zlibstatic
    expat  # expat的目标名称
    minizip  # minizip-ng的库名是minizip
    xlsxwriter  # libxlsxwriter用于写入功能
)

# 确保minizip_legacy被正确构建并链接（libxlsxwriter的依赖）
add_dependencies(TinaXlsx xlsxwriter)
if(TARGET minizip_legacy)
    add_dependencies(TinaXlsx minizip_legacy)
    target_link_libraries(TinaXlsx PRIVATE minizip_legacy)
endif()

# Windows特定设置
if(WIN32)
    target_link_libraries(TinaXlsx PRIVATE ole32 oleaut32 ws2_32 bcrypt)
    
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

# =============================================================================
# 示例程序配置
# =============================================================================

# 添加示例程序
add_executable(example_reader example_reader.cpp)
target_link_libraries(example_reader PRIVATE TinaXlsx)
target_include_directories(example_reader PRIVATE 
    ${CMAKE_CURRENT_SOURCE_DIR}/include
)

# 确保示例程序也有正确的依赖关系
add_dependencies(example_reader TinaXlsx)
if(TARGET minizip_legacy)
    add_dependencies(example_reader minizip_legacy)
endif()

# 添加功能完善测试程序
add_executable(test_reader_enhanced test_reader_functionality.cpp)
target_link_libraries(test_reader_enhanced PRIVATE TinaXlsx)
target_include_directories(test_reader_enhanced PRIVATE 
    ${CMAKE_CURRENT_SOURCE_DIR}/include
)
add_dependencies(test_reader_enhanced TinaXlsx)
if(TARGET minizip_legacy)
    add_dependencies(test_reader_enhanced minizip_legacy)
endif()

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

# 创建别名目标（供内部使用）
add_library(TinaXlsx::TinaXlsx ALIAS TinaXlsx)

# 基本安装 - 只安装库和头文件，不导出目标配置（避免第三方依赖问题）
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

# 安装文档
install(FILES README.md LICENSE
    DESTINATION ${CMAKE_INSTALL_DOCDIR}
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
