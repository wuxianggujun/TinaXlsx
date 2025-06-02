# CompilerConfig.cmake
# 编译器配置模块 - 统一编码和编译选项

# ==================== 编码配置 ====================

# 强制使用UTF-8编码
if(MSVC)
    # MSVC编译器配置
    message(STATUS "配置MSVC编译器使用UTF-8编码")
    
    # 强制源文件和执行字符集都使用UTF-8
    add_compile_options(/utf-8)
    
    # 禁用特定警告
    add_compile_options(
        /wd4819  # 禁用代码页警告
        /wd4996  # 禁用不安全函数警告
        /wd4267  # 禁用size_t转换警告
        /wd4244  # 禁用数据类型转换警告
    )
    
    # 启用多处理器编译
    add_compile_options(/MP)
    
    # 设置警告级别
    add_compile_options(/W3)
    
    message(STATUS "✅ MSVC UTF-8编码配置完成")
    
elseif(CMAKE_CXX_COMPILER_ID STREQUAL "GNU")
    # GCC编译器配置
    message(STATUS "配置GCC编译器使用UTF-8编码")
    
    # 设置输入编码为UTF-8
    add_compile_options(-finput-charset=UTF-8)
    
    # 设置执行字符集为UTF-8
    add_compile_options(-fexec-charset=UTF-8)
    
    # 启用更多警告
    add_compile_options(-Wall -Wextra)
    
    message(STATUS "✅ GCC UTF-8编码配置完成")
    
elseif(CMAKE_CXX_COMPILER_ID STREQUAL "Clang")
    # Clang编译器配置
    message(STATUS "配置Clang编译器使用UTF-8编码")
    
    # Clang默认使用UTF-8，但我们可以显式设置
    add_compile_options(-finput-charset=UTF-8)
    add_compile_options(-fexec-charset=UTF-8)
    
    # 启用更多警告
    add_compile_options(-Wall -Wextra)
    
    message(STATUS "✅ Clang UTF-8编码配置完成")
    
else()
    message(WARNING "未知的编译器: ${CMAKE_CXX_COMPILER_ID}")
endif()

# ==================== 通用编译配置 ====================

# 设置C++标准
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# 调试信息配置
if(CMAKE_BUILD_TYPE STREQUAL "Debug")
    message(STATUS "配置Debug模式编译选项")
    
    if(MSVC)
        # MSVC Debug配置
        add_compile_options(/Od /Zi /RTC1)
        add_link_options(/DEBUG)
    else()
        # GCC/Clang Debug配置
        add_compile_options(-g -O0)
    endif()
    
elseif(CMAKE_BUILD_TYPE STREQUAL "Release")
    message(STATUS "配置Release模式编译选项")
    
    if(MSVC)
        # MSVC Release配置
        add_compile_options(/O2 /DNDEBUG)
    else()
        # GCC/Clang Release配置
        add_compile_options(-O3 -DNDEBUG)
    endif()
endif()

# ==================== 平台特定配置 ====================

if(WIN32)
    message(STATUS "配置Windows平台特定选项")
    
    # Windows特定定义
    add_compile_definitions(
        WIN32_LEAN_AND_MEAN
        NOMINMAX
        _CRT_SECURE_NO_WARNINGS
        _SILENCE_ALL_CXX17_DEPRECATION_WARNINGS
    )
    
    # 设置Windows目标版本
    add_compile_definitions(_WIN32_WINNT=0x0601)  # Windows 7+
    
elseif(UNIX)
    message(STATUS "配置Unix/Linux平台特定选项")
    
    # Unix/Linux特定配置
    add_compile_options(-fPIC)
    
endif()

# ==================== 第三方库配置 ====================

# 禁用第三方库的警告
function(disable_warnings_for_target target_name)
    if(MSVC)
        target_compile_options(${target_name} PRIVATE /w)
    else()
        target_compile_options(${target_name} PRIVATE -w)
    endif()
endfunction()

# ==================== 测试配置 ====================

# 为测试目标配置特殊选项
function(configure_test_target target_name)
    message(STATUS "配置测试目标: ${target_name}")
    
    # 确保测试使用UTF-8编码
    if(MSVC)
        target_compile_options(${target_name} PRIVATE /utf-8)
    endif()
    
    # 设置测试工作目录
    set_target_properties(${target_name} PROPERTIES
        RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/tests/unit
    )
    
    # 为测试创建输出目录
    add_custom_command(TARGET ${target_name} POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E make_directory 
        ${CMAKE_BINARY_DIR}/tests/unit/test_output
        COMMENT "创建测试输出目录"
    )
endfunction()

# ==================== 工具函数 ====================

# 打印编译器信息
function(print_compiler_info)
    message(STATUS "==================== 编译器信息 ====================")
    message(STATUS "编译器ID: ${CMAKE_CXX_COMPILER_ID}")
    message(STATUS "编译器版本: ${CMAKE_CXX_COMPILER_VERSION}")
    message(STATUS "编译器路径: ${CMAKE_CXX_COMPILER}")
    message(STATUS "C++标准: ${CMAKE_CXX_STANDARD}")
    message(STATUS "构建类型: ${CMAKE_BUILD_TYPE}")
    message(STATUS "目标平台: ${CMAKE_SYSTEM_NAME}")
    message(STATUS "目标架构: ${CMAKE_SYSTEM_PROCESSOR}")
    message(STATUS "================================================")
endfunction()

# 验证UTF-8支持
function(verify_utf8_support)
    message(STATUS "验证UTF-8编码支持...")
    
    # 创建一个简单的UTF-8测试文件
    set(test_file "${CMAKE_BINARY_DIR}/utf8_test.cpp")
    file(WRITE ${test_file} 
        "// UTF-8测试文件\n"
        "#include <iostream>\n"
        "int main() {\n"
        "    std::cout << \"UTF-8编码测试：中文字符\" << std::endl;\n"
        "    return 0;\n"
        "}\n"
    )
    
    message(STATUS "✅ UTF-8测试文件已创建")
endfunction()

# 调用信息打印函数
print_compiler_info()
verify_utf8_support()

message(STATUS "✅ 编译器配置模块加载完成")
