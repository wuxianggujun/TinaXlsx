#
# @file tests/CMakeLists.txt
# @brief TinaXlsx 测试配置 - 使用 TestUtils.cmake 工具函数
#

message(STATUS "Configuring TinaXlsx Tests...")

# 确保包含测试工具函数
include(${CMAKE_SOURCE_DIR}/cmake/TestUtils.cmake)

# 启用测试
enable_testing()

# =============================================================================
# 功能测试 (Functional Tests)
# =============================================================================

# 内存优先架构API测试
add_unit_test(MemoryFirstAPI 
    functional/test_memory_first_api.cpp
)


add_unit_test(Variant
    unit/test_variant.cpp
)

# 🚀 高性能日志库测试
add_unit_test(HighPerformanceLogger
    unit/test_high_performance_logger.cpp
)


# =============================================================================
# 性能测试 (Performance Tests)
# =============================================================================

# 极致性能测试
add_performance_test(ExtremePerformance 
    performance/test_extreme_performance.cpp
)


# =============================================================================
# 集成测试 (Integration Tests)
# =============================================================================

# 完整工作流集成测试
add_unit_test(WorkflowIntegration
    integration/test_workflow_integration.cpp
)

# =============================================================================
# 调试测试 (Debug Tests)
# =============================================================================

# 包含调试测试子目录
add_subdirectory(debug)

# =============================================================================
# 批量运行目标
# =============================================================================

# 创建批量测试目标
create_batch_test_targets()

# =============================================================================
# 自定义测试目标
# =============================================================================

# 性能基准测试 - 单独目标
add_custom_target(RunPerformanceBenchmark
    COMMAND echo "🚀 运行性能基准测试..."
    COMMAND ExtremePerformancePerformanceTests
    COMMAND TwoMsChallengePerformanceTests
    COMMAND SIMDPerformancePerformanceTests
    DEPENDS 
        ExtremePerformancePerformanceTests 
        TwoMsChallengePerformanceTests
        SIMDPerformancePerformanceTests
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
    COMMENT "运行完整性能基准测试套件"
)

# 验证核心功能 - 快速验证
add_custom_target(ValidateCore
    COMMAND echo "✅ 验证核心功能..."
    COMMAND MemoryFirstAPITests
    COMMAND BatchSIMDProcessorTests
    COMMAND ZeroCopySerializerTests
    DEPENDS 
        MemoryFirstAPITests
        BatchSIMDProcessorTests  
        ZeroCopySerializerTests
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
    COMMENT "验证核心组件功能"
)

# 2ms挑战专项测试
add_custom_target(Challenge2Ms
    COMMAND echo "🎯 2ms挑战开始..."
    COMMAND TwoMsChallengePerformanceTests --gtest_filter="*TwoMillisecond*"
    DEPENDS TwoMsChallengePerformanceTests
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}/tests/performance
    COMMENT "🚀 2ms性能挑战测试"
)

# =============================================================================
# 测试配置摘要
# =============================================================================

# 打印测试配置摘要
print_test_summary()

message(STATUS "")
message(STATUS "🚀 TinaXlsx 测试配置完成！")
message(STATUS "")
message(STATUS "可用的测试目标:")
message(STATUS "  基础目标:")
message(STATUS "    - make test              : 运行所有CTest测试")
message(STATUS "    - make RunAllUnitTests   : 运行所有单元测试")
message(STATUS "    - make RunAllPerformanceTests : 运行所有性能测试")
message(STATUS "    - make RunQuickTests     : 运行快速测试")
message(STATUS "")
message(STATUS "  专项目标:")
message(STATUS "    - make ValidateCore      : 验证核心功能")
message(STATUS "    - make Challenge2Ms      : 🎯 2ms挑战")
message(STATUS "    - make RunPerformanceBenchmark : 🚀 性能基准测试")
message(STATUS "")
message(STATUS "  单独测试:")
message(STATUS "    - make RunMemoryFirstAPITests : API功能测试")
message(STATUS "    - make RunExtremePerformancePerformanceTests : 极致性能测试")
message(STATUS "")
message(STATUS "==========================================") 
