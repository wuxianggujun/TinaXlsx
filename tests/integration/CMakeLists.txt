# 集成测试配置

# 收集所有集成测试源文件
file(GLOB_RECURSE INTEGRATION_TEST_SOURCES
    "*.cpp"
    "*.hpp"
)

# 只有存在测试文件时才创建目标
if(INTEGRATION_TEST_SOURCES)
    # 创建集成测试可执行文件
    add_executable(integration_tests ${INTEGRATION_TEST_SOURCES})

    # 设置包含目录
    target_include_directories(integration_tests PRIVATE ${TEST_INCLUDE_DIRS})

    # 链接库
    target_link_libraries(integration_tests PRIVATE ${TEST_LINK_LIBRARIES})

    # 设置C++标准
    set_target_properties(integration_tests PROPERTIES
        CXX_STANDARD 17
        CXX_STANDARD_REQUIRED ON
    )

    # 添加到CTest
    gtest_discover_tests(integration_tests
        WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
        PROPERTIES VS_DEBUGGER_WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}"
    )
else()
    message(STATUS "没有找到集成测试文件，跳过集成测试构建")
endif() 