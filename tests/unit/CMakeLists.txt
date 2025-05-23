# 单元测试配置

# 收集所有单元测试源文件
file(GLOB_RECURSE UNIT_TEST_SOURCES
    "*.cpp"
    "*.hpp"
)

# 创建单元测试可执行文件
add_executable(unit_tests ${UNIT_TEST_SOURCES})

# 设置包含目录
target_include_directories(unit_tests PRIVATE ${TEST_INCLUDE_DIRS})

# 链接库
target_link_libraries(unit_tests PRIVATE ${TEST_LINK_LIBRARIES})

# 设置C++标准
set_target_properties(unit_tests PROPERTIES
    CXX_STANDARD 17
    CXX_STANDARD_REQUIRED ON
)

# 添加到CTest
gtest_discover_tests(unit_tests
    WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
    PROPERTIES VS_DEBUGGER_WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}"
) 