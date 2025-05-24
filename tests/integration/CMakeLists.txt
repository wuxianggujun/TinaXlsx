# ==========================================
# TinaXlsx Integration Tests
# ==========================================

message(STATUS "配置 TinaXlsx 集成测试...")

# 添加集成测试程序
function(add_integration_test test_name source_file)
    add_executable(${test_name} ${source_file})
    target_link_libraries(${test_name} PRIVATE TinaXlsx gtest gtest_main)
    target_include_directories(${test_name} PRIVATE ${CMAKE_SOURCE_DIR}/include)
    
    # 设置输出目录
    set_target_properties(${test_name} PROPERTIES
        RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/tests/integration
        RUNTIME_OUTPUT_DIRECTORY_DEBUG ${CMAKE_BINARY_DIR}/tests/integration
        RUNTIME_OUTPUT_DIRECTORY_RELEASE ${CMAKE_BINARY_DIR}/tests/integration
    )
    
    # 添加到CTest（如果使用CTest）
    add_test(NAME ${test_name} COMMAND ${test_name})
    
    message(STATUS "  添加集成测试: ${test_name}")
endfunction()

# 添加我们的主要集成测试
if(EXISTS "${CMAKE_CURRENT_SOURCE_DIR}/test_reader_comprehensive.cpp")
    add_integration_test(test_reader_comprehensive test_reader_comprehensive.cpp)
    
    # 设置测试标签，便于分类运行
    set_tests_properties(test_reader_comprehensive PROPERTIES
        LABELS "reader;integration;comprehensive"
        TIMEOUT 300  # 5分钟超时
    )
endif()

# 如果还有旧的测试文件，保持兼容性
if(EXISTS "${CMAKE_CURRENT_SOURCE_DIR}/test_reader_functionality.cpp")
    add_integration_test(test_reader_functionality test_reader_functionality.cpp)
    set_tests_properties(test_reader_functionality PROPERTIES
        LABELS "reader;integration;legacy"
    )
endif()

# 收集所有其他集成测试源文件（如果有的话）
file(GLOB_RECURSE OTHER_INTEGRATION_TESTS
    "test_*.cpp"
)

# 为每个测试文件创建单独的可执行文件
foreach(test_file ${OTHER_INTEGRATION_TESTS})
    get_filename_component(test_name ${test_file} NAME_WE)
    # 跳过已经处理的测试
    if(NOT test_name STREQUAL "test_reader_comprehensive" AND 
       NOT test_name STREQUAL "test_reader_functionality")
        add_integration_test(${test_name} ${test_file})
        set_tests_properties(${test_name} PROPERTIES
            LABELS "integration;other"
        )
    endif()
endforeach()

# 添加测试组，便于批量运行
if(TARGET test_reader_comprehensive)
    # 创建一个运行所有Reader测试的自定义目标
    add_custom_target(run_reader_tests
        COMMAND ${CMAKE_CTEST_COMMAND} -L reader -V
        DEPENDS test_reader_comprehensive
        COMMENT "运行所有Reader相关测试"
    )
    
    # 创建性能测试目标
    add_custom_target(run_performance_tests
        COMMAND $<TARGET_FILE:test_reader_comprehensive> --gtest_also_run_disabled_tests --gtest_filter="*Performance*"
        DEPENDS test_reader_comprehensive
        COMMENT "运行性能测试（包括默认禁用的测试）"
    )
endif()

message(STATUS "集成测试配置完成") 