# 单元测试 CMakeLists.txt - 使用新的工具函数简化配置

# 确保包含必要的模块
include(GoogleTest)

# ==================== 使用新工具函数创建测试 ====================

# 1. 数据筛选测试
create_test_executable(
    TARGET_NAME DataFilterTests
    SOURCES test_data_filter.cpp test_file_generator.cpp
    COMMENT "Running data filter tests"
)

# 2. 数值工具测试
create_test_executable(
    TARGET_NAME NumberUtilsTests
    SOURCES test_number_utils.cpp test_file_generator.cpp
    COMMENT "Running number utils tests"
)

# 3. 数据功能测试
create_test_executable(
    TARGET_NAME DataFeaturesTests
    SOURCES test_data_features.cpp test_file_generator.cpp
    COMMENT "Running data features tests"
)

# 4. 图表功能测试
create_test_executable(
    TARGET_NAME ChartTests
    SOURCES
        test_chart_functionality.cpp
        test_chart_refactoring.cpp
        test_chart_phase2.cpp
        test_file_generator.cpp
    COMMENT "Running chart tests"
)

# 5. 基础功能测试
create_test_executable(
    TARGET_NAME BasicTests
    SOURCES
        test_basic_features.cpp
        test_cell_formatting.cpp
        test_file_generator.cpp
    COMMENT "Running basic tests"
)

# 6. 高级功能测试
create_test_executable(
    TARGET_NAME AdvancedFeaturesTests
    SOURCES test_pivot_table.cpp test_file_generator.cpp
    COMMENT "Running advanced features tests"
)

# 7. 统一透视表测试
create_test_executable(
    TARGET_NAME UnifiedPivotTableTests
    SOURCES test_pivot_table_unified.cpp test_file_generator.cpp
    COMMENT "Running unified pivot table tests (all pivot table functionality)"
)

# ==================== 创建统一的测试可执行文件（可选） ====================

# 定义测试源文件
set(UNIT_TEST_SOURCES
    # 原有测试
    test_basic_features.cpp
    test_column_width_row_height.cpp
    test_sheet_protection.cpp
    test_enhanced_formulas.cpp
    test_cell_locking.cpp

    # 重构后的新测试
    test_cell_manager.cpp
    test_row_column_manager.cpp
    test_sheet_protection_manager.cpp
    test_sheet_refactored_integration.cpp

    # 测试文件生成器
    test_file_generator.cpp

    # 调试测试
    test_column_width_debug.cpp
    test_column_width_decimal.cpp
    test_protection_debug.cpp

    # 图表功能测试
    test_chart_functionality.cpp
    test_chart_refactoring.cpp
    test_chart_phase2.cpp

    # 数据功能测试
    test_data_features.cpp
    test_data_filter.cpp
)

add_executable(TinaXlsx_UnitTests ${UNIT_TEST_SOURCES})
target_link_libraries(TinaXlsx_UnitTests PRIVATE TinaXlsx gtest_main)
configure_test_target(TinaXlsx_UnitTests)

# 添加测试到CTest - 这会自动发现所有测试用例
gtest_discover_tests(TinaXlsx_UnitTests
    WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
    PROPERTIES VS_DEBUGGER_WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}"
)

# ==================== 配置快速测试 ====================

# 配置快速单元测试目标（只运行关键测试）
create_quick_test_target(
    TARGET_NAME RunQuickUnitTests
    TESTS
        DataFilterTests
        NumberUtilsTests
        BasicTests
)
