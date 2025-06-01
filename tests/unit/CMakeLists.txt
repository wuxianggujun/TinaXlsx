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

# 添加自定义目标来运行所有单元测试
add_custom_target(run_unit_tests
    COMMAND ${CMAKE_CTEST_COMMAND} --output-on-failure --verbose
    DEPENDS TinaXlsx_UnitTests
    COMMENT "Running all unit tests"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

# 添加一个快速测试目标（只运行基本测试）
add_custom_target(run_quick_tests
    COMMAND TinaXlsx_UnitTests --gtest_filter="BasicFeaturesTest.*"
    DEPENDS TinaXlsx_UnitTests
    COMMENT "Running quick tests"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

# 添加重构测试目标（只运行重构后的测试）
add_custom_target(run_refactored_tests
    COMMAND TinaXlsx_UnitTests --gtest_filter="*Manager*:*Refactored*"
    DEPENDS TinaXlsx_UnitTests
    COMMENT "Running refactored component tests"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

# 添加管理器测试目标
add_custom_target(run_manager_tests
    COMMAND TinaXlsx_UnitTests --gtest_filter="TX*ManagerTest.*"
    DEPENDS TinaXlsx_UnitTests
    COMMENT "Running manager component tests"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)

# 添加集成测试目标
add_custom_target(run_integration_tests
    COMMAND TinaXlsx_UnitTests --gtest_filter="TXSheetRefactoredIntegrationTest.*"
    DEPENDS TinaXlsx_UnitTests
    COMMENT "Running integration tests for refactored sheet"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
)