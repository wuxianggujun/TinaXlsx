#
# @file TestUtils.cmake
# @brief TinaXlsx 测试工具函数 - 统一测试配置管理
#

# 配置单个测试目标的通用函数
function(configure_test_target target_name)
    # 设置C++17标准
    target_compile_features(${target_name} PRIVATE cxx_std_17)
    
    # 设置编译选项
    if(MSVC)
        target_compile_options(${target_name} PRIVATE 
            /W3 
            /utf-8
            /wd4819  # 忽略字符编码警告
        )
        
        # Windows特定编译定义
        target_compile_definitions(${target_name} PRIVATE
            WIN32_LEAN_AND_MEAN
            NOMINMAX
            _CRT_SECURE_NO_WARNINGS
        )
    else()
        target_compile_options(${target_name} PRIVATE 
            -Wall 
            -Wno-unused-parameter
        )
    endif()
    
    # 设置包含目录
    target_include_directories(${target_name} PRIVATE
        ${CMAKE_SOURCE_DIR}/include
        ${CMAKE_SOURCE_DIR}/tests/unit  # 用于test_file_generator.hpp
        ${CMAKE_SOURCE_DIR}/third_party/googletest/googletest/include
    )
endfunction()

# 创建标准测试可执行文件的通用函数
function(create_test_executable)
    # 解析参数
    set(options PERFORMANCE_TEST)  # 是否为性能测试
    set(oneValueArgs TARGET_NAME COMMENT)  # 单值参数
    set(multiValueArgs SOURCES DEPENDENCIES)  # 多值参数
    
    cmake_parse_arguments(TEST "${options}" "${oneValueArgs}" "${multiValueArgs}" ${ARGN})
    
    # 验证必需参数
    if(NOT TEST_TARGET_NAME)
        message(FATAL_ERROR "create_test_executable: TARGET_NAME is required")
    endif()
    
    if(NOT TEST_SOURCES)
        message(FATAL_ERROR "create_test_executable: SOURCES is required")
    endif()
    
    # 创建可执行文件
    add_executable(${TEST_TARGET_NAME} ${TEST_SOURCES})
    
    # 链接库
    target_link_libraries(${TEST_TARGET_NAME}
        PRIVATE
        TinaXlsx
        gtest_main
        gtest
    )
    
    # 应用通用配置
    configure_test_target(${TEST_TARGET_NAME})
    
    # 设置输出目录
    if(TEST_PERFORMANCE_TEST)
        set_target_properties(${TEST_TARGET_NAME} PROPERTIES
            RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
        )
        set(test_working_dir ${CMAKE_BINARY_DIR}/tests/performance)
    else()
        set_target_properties(${TEST_TARGET_NAME} PROPERTIES
            RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/tests/unit
        )
        set(test_working_dir ${CMAKE_BINARY_DIR})
    endif()
    
    # 注册到CTest
    add_test(NAME ${TEST_TARGET_NAME} COMMAND ${TEST_TARGET_NAME})
    set_tests_properties(${TEST_TARGET_NAME} PROPERTIES 
        WORKING_DIRECTORY ${test_working_dir}
    )
    
    # 创建单独的运行目标
    if(NOT TEST_COMMENT)
        set(TEST_COMMENT "Running ${TEST_TARGET_NAME}")
    endif()
    
    add_custom_target(Run${TEST_TARGET_NAME}
        COMMAND ${TEST_TARGET_NAME}
        DEPENDS ${TEST_TARGET_NAME}
        COMMENT "${TEST_COMMENT}"
        WORKING_DIRECTORY ${test_working_dir}
    )
    
    # 如果有依赖，添加到全局列表
    if(TEST_DEPENDENCIES)
        foreach(dep ${TEST_DEPENDENCIES})
            add_dependencies(${TEST_TARGET_NAME} ${dep})
        endforeach()
    endif()
    
    # 将测试添加到全局列表（用于批量运行）
    if(TEST_PERFORMANCE_TEST)
        set_property(GLOBAL APPEND PROPERTY PERFORMANCE_TESTS ${TEST_TARGET_NAME})
    else()
        set_property(GLOBAL APPEND PROPERTY UNIT_TESTS ${TEST_TARGET_NAME})
    endif()
    
    message(STATUS "Created test: ${TEST_TARGET_NAME} (${TEST_COMMENT})")
endfunction()

# 创建批量运行目标的函数
function(create_batch_test_targets)
    # 获取所有单元测试
    get_property(unit_tests GLOBAL PROPERTY UNIT_TESTS)
    if(unit_tests)
        # 创建运行所有单元测试的目标
        add_custom_target(RunAllUnitTests
            COMMAND echo "Running all unit tests..."
            COMMENT "Running all unit test executables"
            WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
        )
        
        foreach(test ${unit_tests})
            add_custom_command(TARGET RunAllUnitTests POST_BUILD
                COMMAND ${test}
                WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
                COMMENT "Running ${test}"
            )
            add_dependencies(RunAllUnitTests ${test})
        endforeach()
        
        message(STATUS "Created batch target: RunAllUnitTests (${list_length} tests)")
    endif()
    
    # 获取所有性能测试
    get_property(performance_tests GLOBAL PROPERTY PERFORMANCE_TESTS)
    if(performance_tests)
        # 创建运行所有性能测试的目标
        add_custom_target(RunAllPerformanceTests
            COMMAND echo "Running all performance tests..."
            COMMENT "Running all performance test executables"
            WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
        )
        
        foreach(test ${performance_tests})
            add_custom_command(TARGET RunAllPerformanceTests POST_BUILD
                COMMAND ${test}
                WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
                COMMENT "Running ${test}"
            )
            add_dependencies(RunAllPerformanceTests ${test})
        endforeach()
        
        message(STATUS "Created batch target: RunAllPerformanceTests")
    endif()
endfunction()

# 简化的测试创建宏（最常用的情况）
macro(add_unit_test test_name)
    set(sources ${ARGN})
    if(NOT sources)
        set(sources "test_${test_name}.cpp" "test_file_generator.cpp")
    endif()
    
    create_test_executable(
        TARGET_NAME ${test_name}Tests
        SOURCES ${sources}
        COMMENT "Running ${test_name} tests"
    )
endmacro()

macro(add_performance_test test_name)
    set(sources ${ARGN})
    if(NOT sources)
        set(sources "test_${test_name}_performance.cpp")
    endif()
    
    create_test_executable(
        TARGET_NAME ${test_name}PerformanceTests
        SOURCES ${sources}
        COMMENT "Running ${test_name} performance tests"
        PERFORMANCE_TEST
    )
endmacro()


# 打印测试配置摘要
function(print_test_summary)
    get_property(unit_tests GLOBAL PROPERTY UNIT_TESTS)
    get_property(performance_tests GLOBAL PROPERTY PERFORMANCE_TESTS)
    
    message(STATUS "")
    message(STATUS "=== TinaXlsx Test Configuration Summary ===")
    
    if(unit_tests)
        list(LENGTH unit_tests unit_count)
        message(STATUS "Unit Tests (${unit_count}):")
        foreach(test ${unit_tests})
            message(STATUS "  - ${test}")
        endforeach()
    endif()
    
    if(performance_tests)
        list(LENGTH performance_tests perf_count)
        message(STATUS "Performance Tests (${perf_count}):")
        foreach(test ${performance_tests})
            message(STATUS "  - ${test}")
        endforeach()
    endif()
    
    message(STATUS "")
    message(STATUS "Available targets:")
    message(STATUS "  - run_all_tests          : Run all tests via CTest")
    message(STATUS "  - RunAllUnitTests        : Run all unit tests")
    message(STATUS "  - RunAllPerformanceTests : Run all performance tests")
    message(STATUS "  - RunQuickTests          : Run essential tests only")
    message(STATUS "  - Run<TestName>          : Run individual test")
    message(STATUS "==========================================")
    message(STATUS "")
endfunction()
