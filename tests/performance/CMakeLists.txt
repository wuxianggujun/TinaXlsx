# TinaXlsx 性能测试 CMakeLists.txt - 使用新的工具函数简化配置

# ==================== 使用新工具函数创建性能测试 ====================

# 1. 极致性能测试
create_test_executable(
    TARGET_NAME ExtremePerformanceTests
    SOURCES test_performance_extreme.cpp
    COMMENT "Running extreme performance tests"
    PERFORMANCE_TEST
)

# 2. 简化性能测试
create_test_executable(
    TARGET_NAME SimplePerformanceTests
    SOURCES test_performance_simple.cpp
    COMMENT "Running simple performance tests"
    PERFORMANCE_TEST
)

# 3. 优化性能测试
create_test_executable(
    TARGET_NAME OptimizedPerformanceTests
    SOURCES test_optimized_performance.cpp
    COMMENT "Running optimized performance tests"
    PERFORMANCE_TEST
)

# 4. 批量操作性能测试
create_test_executable(
    TARGET_NAME BatchOperationsTests
    SOURCES test_batch_operations.cpp
    COMMENT "Running batch operations performance tests"
    PERFORMANCE_TEST
)

# 5. Pugixml流式性能测试
create_test_executable(
    TARGET_NAME PugiStreamPerformanceTests
    SOURCES test_pugi_stream_performance.cpp
    COMMENT "Running pugixml streaming performance tests"
    PERFORMANCE_TEST
)

# 6. SharedStrings性能测试
create_test_executable(
    TARGET_NAME SharedStringsPerformanceTests
    SOURCES test_shared_strings_performance.cpp
    COMMENT "Running shared strings performance tests"
    PERFORMANCE_TEST
)

# 7. 高级并行框架性能测试
create_test_executable(
    TARGET_NAME AdvancedParallelPerformanceTests
    SOURCES test_advanced_parallel_performance.cpp ../unit/test_file_generator.cpp
    COMMENT "Running advanced parallel framework performance tests"
    PERFORMANCE_TEST
)

# ==================== 配置快速性能测试 ====================

# 配置快速性能测试目标（只运行关键性能测试）
create_quick_test_target(
    TARGET_NAME RunQuickPerformanceTests
    TESTS
        SimplePerformanceTests
        SharedStringsPerformanceTests
)
