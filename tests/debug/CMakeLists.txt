# Debug测试配置

# 创建Vector调试测试
add_executable(VectorDebugTest test_vector_debug.cpp)

# 链接必要的库
target_link_libraries(VectorDebugTest
    PRIVATE
    TinaXlsx
    gtest
    gtest_main
)

# 设置包含目录
target_include_directories(VectorDebugTest PRIVATE
    ${PROJECT_SOURCE_DIR}/include
)

# 设置编译选项
target_compile_features(VectorDebugTest PRIVATE cxx_std_17)

# 添加到测试
add_test(NAME VectorDebugTest COMMAND VectorDebugTest)

# 设置测试属性
set_tests_properties(VectorDebugTest PROPERTIES
    TIMEOUT 60
    LABELS "debug;vector;crash"
)

# 创建独立运行目标
add_custom_target(run_vector_debug
    COMMAND VectorDebugTest --gtest_output=xml:vector_debug_results.xml
    DEPENDS VectorDebugTest
    COMMENT "运行Vector调试测试"
    WORKING_DIRECTORY ${CMAKE_CURRENT_BINARY_DIR}
)

# 输出信息
message(STATUS "配置Vector调试测试: VectorDebugTest")
message(STATUS "独立运行命令: make run_vector_debug 或直接运行 VectorDebugTest.exe")
