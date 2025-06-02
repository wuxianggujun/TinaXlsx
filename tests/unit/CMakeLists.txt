# 单元测试 CMakeLists.txt

# 确保包含必要的模块
include(GoogleTest)

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

    # 内存优化测试
    test_compact_cell.cpp
)

# 方法1：创建统一的测试可执行文件（推荐）
add_executable(TinaXlsx_UnitTests ${UNIT_TEST_SOURCES})

target_link_libraries(TinaXlsx_UnitTests
    PRIVATE
    TinaXlsx
    gtest_main
)

# 设置C++标准
set_property(TARGET TinaXlsx_UnitTests PROPERTY CXX_STANDARD 17)

# 为测试添加编译选项
if(MSVC)
    target_compile_options(TinaXlsx_UnitTests PRIVATE 
        /W3 
        /utf-8
        /wd4819
    )
else()
    target_compile_options(TinaXlsx_UnitTests PRIVATE -Wall -Wno-unused-parameter)
endif()

# 添加测试到CTest - 这会自动发现所有测试用例
gtest_discover_tests(TinaXlsx_UnitTests
    WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
    PROPERTIES VS_DEBUGGER_WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}"
)

# 方法2：为每个测试文件创建单独的可执行文件（可选）
option(BUILD_INDIVIDUAL_TESTS "Build individual test executables" OFF)

if(BUILD_INDIVIDUAL_TESTS)
    foreach(test_source ${UNIT_TEST_SOURCES})
        get_filename_component(test_name ${test_source} NAME_WE)
        add_executable(${test_name} ${test_source})
        target_link_libraries(${test_name}
            PRIVATE
            TinaXlsx
            gtest_main
        )
        set_property(TARGET ${test_name} PROPERTY CXX_STANDARD 17)
        
        # 编译选项
        if(MSVC)
            target_compile_options(${test_name} PRIVATE /W3 /utf-8 /wd4819)
        else()
            target_compile_options(${test_name} PRIVATE -Wall -Wno-unused-parameter)
        endif()
        
        gtest_discover_tests(${test_name})
    endforeach()
endif()

# ==================== 创建独立的测试可执行文件 ====================

# 1. 数据筛选测试（独立可执行文件）
add_executable(DataFilterTests
    test_data_filter.cpp
    test_file_generator.cpp
)

target_link_libraries(DataFilterTests
    PRIVATE
    ${PROJECT_NAME}
    gtest_main
    gtest
)

target_include_directories(DataFilterTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
)

# 配置测试目标的UTF-8编码
configure_test_target(DataFilterTests)

# 2. 数值工具测试（独立可执行文件）
add_executable(NumberUtilsTests
    test_number_utils.cpp
    test_file_generator.cpp
)

target_link_libraries(NumberUtilsTests
    PRIVATE
    ${PROJECT_NAME}
    gtest_main
    gtest
)

target_include_directories(NumberUtilsTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
)

# 配置测试目标的UTF-8编码
configure_test_target(NumberUtilsTests)

# 3. 数据功能测试（独立可执行文件）
add_executable(DataFeaturesTests
    test_data_features.cpp
    test_file_generator.cpp
)

target_link_libraries(DataFeaturesTests
    PRIVATE
    ${PROJECT_NAME}
    gtest_main
    gtest
)

target_include_directories(DataFeaturesTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
)

# 配置测试目标的UTF-8编码
configure_test_target(DataFeaturesTests)

# 4. 图表功能测试（独立可执行文件）
add_executable(ChartTests
    test_chart_functionality.cpp
    test_chart_refactoring.cpp
    test_chart_phase2.cpp
    test_file_generator.cpp
)

target_link_libraries(ChartTests
    PRIVATE
    ${PROJECT_NAME}
    gtest_main
    gtest
)

target_include_directories(ChartTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
)

# 配置测试目标的UTF-8编码
configure_test_target(ChartTests)

# 5. 基础功能测试（独立可执行文件）
add_executable(BasicTests
    test_basic_features.cpp
    test_cell_formatting.cpp
    test_file_generator.cpp
)

target_link_libraries(BasicTests
    PRIVATE
    ${PROJECT_NAME}
    gtest_main
    gtest
)

target_include_directories(BasicTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
)

# 配置测试目标的UTF-8编码
configure_test_target(BasicTests)

# 6. 高级功能测试（第四阶段）
add_executable(AdvancedFeaturesTests
    test_pivot_table.cpp
    test_file_generator.cpp
)

target_link_libraries(AdvancedFeaturesTests
    PRIVATE
    ${PROJECT_NAME}
    gtest_main
    gtest
)

target_include_directories(AdvancedFeaturesTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
)

# 配置测试目标的UTF-8编码
configure_test_target(AdvancedFeaturesTests)

# 7. 统一透视表测试（包含所有透视表功能）
add_executable(UnifiedPivotTableTests
    test_pivot_table_unified.cpp
    test_file_generator.cpp
)

target_link_libraries(UnifiedPivotTableTests
    PRIVATE
    ${PROJECT_NAME}
    gtest_main
    gtest
)

target_include_directories(UnifiedPivotTableTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
)

# 配置测试目标的UTF-8编码
configure_test_target(UnifiedPivotTableTests)

# 8. 紧凑单元格测试（内存优化）
add_executable(CompactCellTests
    test_compact_cell.cpp
    test_file_generator.cpp
)

target_link_libraries(CompactCellTests
    PRIVATE
    ${PROJECT_NAME}
    gtest_main
    gtest
)

target_include_directories(CompactCellTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
)

# 配置测试目标的UTF-8编码
configure_test_target(CompactCellTests)

# 9. 单元格替换测试（渐进式迁移）
add_executable(CellReplacementTests
    test_cell_replacement.cpp
    test_file_generator.cpp
)

target_link_libraries(CellReplacementTests
    PRIVATE
    ${PROJECT_NAME}
    gtest_main
    gtest
)

target_include_directories(CellReplacementTests PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}
)

# 配置测试目标的UTF-8编码
configure_test_target(CellReplacementTests)

# ==================== 注册测试到CTest ====================

# 注册独立测试到CTest
add_test(NAME DataFilter COMMAND DataFilterTests)
add_test(NAME NumberUtils COMMAND NumberUtilsTests)
add_test(NAME DataFeatures COMMAND DataFeaturesTests)
add_test(NAME Charts COMMAND ChartTests)
add_test(NAME Basic COMMAND BasicTests)

# 高级功能测试
add_test(NAME AdvancedFeatures COMMAND AdvancedFeaturesTests)
set_tests_properties(AdvancedFeatures PROPERTIES WORKING_DIRECTORY ${CMAKE_BINARY_DIR})

# 统一透视表测试
add_test(NAME UnifiedPivotTable COMMAND UnifiedPivotTableTests)
set_tests_properties(UnifiedPivotTable PROPERTIES WORKING_DIRECTORY ${CMAKE_BINARY_DIR})

# 紧凑单元格测试
add_test(NAME CompactCell COMMAND CompactCellTests)
set_tests_properties(CompactCell PROPERTIES WORKING_DIRECTORY ${CMAKE_BINARY_DIR})

# 单元格替换测试
add_test(NAME CellReplacement COMMAND CellReplacementTests)
set_tests_properties(CellReplacement PROPERTIES WORKING_DIRECTORY ${CMAKE_BINARY_DIR})

# 设置测试工作目录
set_tests_properties(DataFilter PROPERTIES WORKING_DIRECTORY ${CMAKE_BINARY_DIR})
set_tests_properties(NumberUtils PROPERTIES WORKING_DIRECTORY ${CMAKE_BINARY_DIR})
set_tests_properties(DataFeatures PROPERTIES WORKING_DIRECTORY ${CMAKE_BINARY_DIR})
set_tests_properties(Charts PROPERTIES WORKING_DIRECTORY ${CMAKE_BINARY_DIR})
set_tests_properties(Basic PROPERTIES WORKING_DIRECTORY ${CMAKE_BINARY_DIR})

# ==================== IDE 友好的目标 ====================

# 创建一个运行所有独立测试的目标
add_custom_target(RunAllIndependentTests
    COMMAND echo "Running all independent tests..."
    COMMAND DataFilterTests
    COMMAND NumberUtilsTests
    COMMAND DataFeaturesTests
    COMMAND ChartTests
    COMMAND BasicTests
    DEPENDS DataFilterTests NumberUtilsTests DataFeaturesTests ChartTests BasicTests
    COMMENT "Running all independent test executables"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

# 创建快速测试目标（只运行关键测试）
add_custom_target(RunQuickTests
    COMMAND echo "Running quick tests..."
    COMMAND DataFilterTests
    COMMAND NumberUtilsTests
    DEPENDS DataFilterTests NumberUtilsTests
    COMMENT "Running quick tests (DataFilter + NumberUtils)"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

# 为每个测试创建单独的运行目标
add_custom_target(RunDataFilterTest
    COMMAND DataFilterTests
    DEPENDS DataFilterTests
    COMMENT "Running data filter tests"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

add_custom_target(RunNumberUtilsTest
    COMMAND NumberUtilsTests
    DEPENDS NumberUtilsTests
    COMMENT "Running number utils tests"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

add_custom_target(RunDataFeaturesTest
    COMMAND DataFeaturesTests
    DEPENDS DataFeaturesTests
    COMMENT "Running data features tests"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

add_custom_target(RunChartTest
    COMMAND ChartTests
    DEPENDS ChartTests
    COMMENT "Running chart tests"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

add_custom_target(RunBasicTest
    COMMAND BasicTests
    DEPENDS BasicTests
    COMMENT "Running basic tests"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

# 高级功能测试目标
add_custom_target(RunAdvancedFeaturesTest
    COMMAND AdvancedFeaturesTests
    DEPENDS AdvancedFeaturesTests
    COMMENT "Running advanced features tests"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

# 统一透视表测试目标
add_custom_target(RunUnifiedPivotTableTest
    COMMAND UnifiedPivotTableTests
    DEPENDS UnifiedPivotTableTests
    COMMENT "Running unified pivot table tests (all pivot table functionality)"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

# 紧凑单元格测试目标
add_custom_target(RunCompactCellTest
    COMMAND CompactCellTests
    DEPENDS CompactCellTests
    COMMENT "Running compact cell tests (memory optimization)"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

# 单元格替换测试目标
add_custom_target(RunCellReplacementTest
    COMMAND CellReplacementTests
    DEPENDS CellReplacementTests
    COMMENT "Running cell replacement tests (progressive migration)"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)