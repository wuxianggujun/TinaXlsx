# ==========================================
# TinaXlsx Integration Tests
# ==========================================

message(STATUS "配置 TinaXlsx 集成测试...")

# 添加集成测试程序
function(add_integration_test test_name source_file)
    add_executable(${test_name} ${source_file})
    target_link_libraries(${test_name} PRIVATE TinaXlsx)
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

# 添加我们的集成测试
if(EXISTS "${CMAKE_CURRENT_SOURCE_DIR}/test_reader_functionality.cpp")
    add_integration_test(test_reader_functionality test_reader_functionality.cpp)
endif()

# 收集所有其他集成测试源文件（如果有的话）
file(GLOB_RECURSE OTHER_INTEGRATION_TESTS
    "test_*.cpp"
)

# 为每个测试文件创建单独的可执行文件
foreach(test_file ${OTHER_INTEGRATION_TESTS})
    get_filename_component(test_name ${test_file} NAME_WE)
    if(NOT test_name STREQUAL "test_reader_functionality")
        add_integration_test(${test_name} ${test_file})
    endif()
endforeach()

message(STATUS "集成测试配置完成") 