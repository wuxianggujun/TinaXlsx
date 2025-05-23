# 性能测试配置

# 收集所有性能测试源文件
file(GLOB_RECURSE PERFORMANCE_TEST_SOURCES
    "*.cpp"
    "*.hpp"
)

# 只有存在测试文件时才创建目标
if(PERFORMANCE_TEST_SOURCES)
    # 创建性能测试可执行文件
    add_executable(performance_tests ${PERFORMANCE_TEST_SOURCES})

    # 设置包含目录
    target_include_directories(performance_tests PRIVATE ${TEST_INCLUDE_DIRS})

    # 链接库
    target_link_libraries(performance_tests PRIVATE ${TEST_LINK_LIBRARIES})

    # 设置C++标准
    set_target_properties(performance_tests PROPERTIES
        CXX_STANDARD 17
        CXX_STANDARD_REQUIRED ON
    )
    
    # 性能测试优化选项（仅在Release模式下启用）
    if(CMAKE_BUILD_TYPE STREQUAL "Release")
        if(MSVC)
            target_compile_options(performance_tests PRIVATE /O2 /Oi)
        else()
            target_compile_options(performance_tests PRIVATE -O2)
        endif()
    endif()

    # 添加到CTest
    gtest_discover_tests(performance_tests
        WORKING_DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}
        PROPERTIES VS_DEBUGGER_WORKING_DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}"
    )
else()
    message(STATUS "没有找到性能测试文件，跳过性能测试构建")
endif() 