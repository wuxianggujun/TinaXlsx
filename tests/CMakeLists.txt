# 测试配置
cmake_minimum_required(VERSION 3.16)

# 包含Google Test
include(GoogleTest)

# 通用测试设置
set(TEST_INCLUDE_DIRS
    ${CMAKE_CURRENT_SOURCE_DIR}/../include
    ${CMAKE_CURRENT_SOURCE_DIR}/../src
    ${CMAKE_CURRENT_SOURCE_DIR}/../third_party/libxlsxwriter/include
    ${CMAKE_CURRENT_SOURCE_DIR}/../third_party/xlsxio/include
)

set(TEST_LINK_LIBRARIES
    TinaXlsx
    gtest
    gtest_main
    gmock
    gmock_main
)

# 单元测试
add_subdirectory(unit)

# 集成测试
add_subdirectory(integration)

# 性能测试
add_subdirectory(performance)

# 创建一个总的测试目标
add_custom_target(run_all_tests
    COMMAND ${CMAKE_CTEST_COMMAND} --verbose
    COMMENT "运行所有测试"
    WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
) 