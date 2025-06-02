# TinaXlsx 性能测试 CMakeLists.txt

# 创建性能测试可执行文件
add_executable(ExtremePerformanceTests
    test_performance_extreme.cpp
)

# 创建简化性能测试可执行文件
add_executable(SimplePerformanceTests
    test_performance_simple.cpp
)

# 创建优化性能测试可执行文件
add_executable(OptimizedPerformanceTests
    test_optimized_performance.cpp
)

# 创建批量操作性能测试可执行文件
add_executable(BatchOperationsTests
    test_batch_operations.cpp
)

# 创建Pugixml流式性能测试可执行文件
add_executable(PugiStreamPerformanceTests
    test_pugi_stream_performance.cpp
)

# 创建SharedStrings性能测试可执行文件
add_executable(SharedStringsPerformanceTests
    test_shared_strings_performance.cpp
)

# 链接必要的库
target_link_libraries(ExtremePerformanceTests
    PRIVATE
    TinaXlsx
    gtest
    gtest_main
)

target_link_libraries(SimplePerformanceTests
    PRIVATE
    TinaXlsx
    gtest
    gtest_main
)

target_link_libraries(OptimizedPerformanceTests
    PRIVATE
    TinaXlsx
    gtest
    gtest_main
)

target_link_libraries(BatchOperationsTests
    PRIVATE
    TinaXlsx
    gtest
    gtest_main
)

target_link_libraries(PugiStreamPerformanceTests
    PRIVATE
    TinaXlsx
    gtest
    gtest_main
)

target_link_libraries(SharedStringsPerformanceTests
    PRIVATE
    TinaXlsx
    gtest
    gtest_main
)

# 设置包含目录
target_include_directories(ExtremePerformanceTests
    PRIVATE
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/third_party/googletest/googletest/include
)

target_include_directories(SimplePerformanceTests
    PRIVATE
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/third_party/googletest/googletest/include
)

target_include_directories(OptimizedPerformanceTests
    PRIVATE
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/third_party/googletest/googletest/include
)

target_include_directories(BatchOperationsTests
    PRIVATE
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/third_party/googletest/googletest/include
)

target_include_directories(PugiStreamPerformanceTests
    PRIVATE
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/third_party/googletest/googletest/include
)

target_include_directories(SharedStringsPerformanceTests
    PRIVATE
    ${CMAKE_SOURCE_DIR}/include
    ${CMAKE_SOURCE_DIR}/third_party/googletest/googletest/include
)

# 设置编译选项
target_compile_features(ExtremePerformanceTests PRIVATE cxx_std_17)
target_compile_features(SimplePerformanceTests PRIVATE cxx_std_17)
target_compile_features(OptimizedPerformanceTests PRIVATE cxx_std_17)
target_compile_features(BatchOperationsTests PRIVATE cxx_std_17)
target_compile_features(PugiStreamPerformanceTests PRIVATE cxx_std_17)
target_compile_features(SharedStringsPerformanceTests PRIVATE cxx_std_17)

# Windows特定设置
if(WIN32)
    target_compile_definitions(ExtremePerformanceTests PRIVATE
        WIN32_LEAN_AND_MEAN
        NOMINMAX
        _CRT_SECURE_NO_WARNINGS
    )
    target_compile_definitions(SimplePerformanceTests PRIVATE
        WIN32_LEAN_AND_MEAN
        NOMINMAX
        _CRT_SECURE_NO_WARNINGS
    )
    target_compile_definitions(OptimizedPerformanceTests PRIVATE
        WIN32_LEAN_AND_MEAN
        NOMINMAX
        _CRT_SECURE_NO_WARNINGS
    )
    target_compile_definitions(BatchOperationsTests PRIVATE
        WIN32_LEAN_AND_MEAN
        NOMINMAX
        _CRT_SECURE_NO_WARNINGS
    )
    target_compile_definitions(PugiStreamPerformanceTests PRIVATE
        WIN32_LEAN_AND_MEAN
        NOMINMAX
        _CRT_SECURE_NO_WARNINGS
    )
    target_compile_definitions(SharedStringsPerformanceTests PRIVATE
        WIN32_LEAN_AND_MEAN
        NOMINMAX
        _CRT_SECURE_NO_WARNINGS
    )
endif()

# 设置输出目录
set_target_properties(ExtremePerformanceTests PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
)
set_target_properties(SimplePerformanceTests PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
)

set_target_properties(OptimizedPerformanceTests PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
)

set_target_properties(BatchOperationsTests PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
)

set_target_properties(PugiStreamPerformanceTests PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
)

set_target_properties(SharedStringsPerformanceTests PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
)

# 添加测试
add_test(NAME ExtremePerformanceTests COMMAND ExtremePerformanceTests)
add_test(NAME SimplePerformanceTests COMMAND SimplePerformanceTests)
add_test(NAME OptimizedPerformanceTests COMMAND OptimizedPerformanceTests)
add_test(NAME BatchOperationsTests COMMAND BatchOperationsTests)
add_test(NAME PugiStreamPerformanceTests COMMAND PugiStreamPerformanceTests)
add_test(NAME SharedStringsPerformanceTests COMMAND SharedStringsPerformanceTests)

# 设置测试工作目录
set_tests_properties(ExtremePerformanceTests PROPERTIES
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
)
set_tests_properties(SimplePerformanceTests PROPERTIES
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
)

set_tests_properties(OptimizedPerformanceTests PROPERTIES
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
)

set_tests_properties(BatchOperationsTests PROPERTIES
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
)

set_tests_properties(PugiStreamPerformanceTests PROPERTIES
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
)

set_tests_properties(SharedStringsPerformanceTests PROPERTIES
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
)
