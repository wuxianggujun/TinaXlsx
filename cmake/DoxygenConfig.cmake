# DoxygenConfig.cmake
# Doxygen文档生成配置模块

# ==================== Doxygen配置 ====================

# 首先尝试使用项目内的Doxygen
set(DOXYGEN_SOURCE_DIR "${CMAKE_SOURCE_DIR}/third_party/doxygen")
set(DOXYGEN_BUILD_DIR "${CMAKE_BINARY_DIR}/doxygen-build")
set(DOXYGEN_EXECUTABLE "${DOXYGEN_BUILD_DIR}/bin/doxygen")

# 检查项目内Doxygen是否存在
if(EXISTS "${DOXYGEN_SOURCE_DIR}/CMakeLists.txt")
    message(STATUS "🔧 使用项目内的Doxygen源码")

    # 配置Doxygen构建选项
    set(build_doc OFF CACHE BOOL "Build documentation" FORCE)
    set(build_wizard OFF CACHE BOOL "Build wizard" FORCE)
    set(build_xmlparser OFF CACHE BOOL "Build XML parser" FORCE)
    set(build_search OFF CACHE BOOL "Build search" FORCE)
    set(build_parse OFF CACHE BOOL "Build parse" FORCE)
    set(use_sqlite3 OFF CACHE BOOL "Use sqlite3" FORCE)
    set(use_libclang OFF CACHE BOOL "Use libclang" FORCE)

    # 添加Doxygen子项目
    add_subdirectory(${DOXYGEN_SOURCE_DIR} ${DOXYGEN_BUILD_DIR} EXCLUDE_FROM_ALL)

    # 设置Doxygen可执行文件路径
    if(WIN32)
        set(DOXYGEN_EXECUTABLE "${DOXYGEN_BUILD_DIR}/bin/doxygen.exe")
    else()
        set(DOXYGEN_EXECUTABLE "${DOXYGEN_BUILD_DIR}/bin/doxygen")
    endif()

    set(DOXYGEN_FOUND TRUE)
    message(STATUS "✅ 项目内Doxygen配置完成: ${DOXYGEN_EXECUTABLE}")

else()
    # 回退到系统Doxygen
    find_package(Doxygen QUIET)

    if(DOXYGEN_FOUND)
        message(STATUS "✅ 找到系统Doxygen: ${DOXYGEN_EXECUTABLE}")
    else()
        message(STATUS "⚠️  未找到Doxygen，请检查third_party/doxygen目录")
    endif()
endif()

if(DOXYGEN_FOUND)
    
    # 设置Doxygen配置变量
    set(DOXYGEN_PROJECT_NAME "TinaXlsx")
    set(DOXYGEN_PROJECT_VERSION ${PROJECT_VERSION})
    set(DOXYGEN_PROJECT_BRIEF "高性能C++ Excel文件读写库")
    
    # 输入和输出目录
    set(DOXYGEN_INPUT_DIR "${CMAKE_SOURCE_DIR}/include ${CMAKE_SOURCE_DIR}/src")
    set(DOXYGEN_OUTPUT_DIR "${CMAKE_SOURCE_DIR}/api-docs")
    set(DOXYGEN_HTML_OUTPUT "html")
    set(DOXYGEN_LATEX_OUTPUT "latex")
    
    # 文档生成选项
    set(DOXYGEN_GENERATE_HTML YES)
    set(DOXYGEN_GENERATE_LATEX NO)
    set(DOXYGEN_GENERATE_XML YES)
    set(DOXYGEN_GENERATE_MAN NO)
    
    # 源码分析选项
    set(DOXYGEN_EXTRACT_ALL YES)
    set(DOXYGEN_EXTRACT_PRIVATE YES)
    set(DOXYGEN_EXTRACT_STATIC YES)
    set(DOXYGEN_EXTRACT_LOCAL_CLASSES YES)
    set(DOXYGEN_EXTRACT_ANON_NSPACES YES)
    
    # 包含图表
    set(DOXYGEN_HAVE_DOT NO)
    set(DOXYGEN_CLASS_DIAGRAMS YES)
    set(DOXYGEN_COLLABORATION_GRAPH YES)
    set(DOXYGEN_INCLUDE_GRAPH YES)
    set(DOXYGEN_INCLUDED_BY_GRAPH YES)
    set(DOXYGEN_CALL_GRAPH YES)
    set(DOXYGEN_CALLER_GRAPH YES)
    
    # 文件模式
    set(DOXYGEN_FILE_PATTERNS "*.hpp *.h *.cpp *.c *.md")
    set(DOXYGEN_RECURSIVE YES)
    set(DOXYGEN_EXCLUDE_PATTERNS "*/third_party/* */cmake-build-*/* */build/*")
    
    # 输出格式
    set(DOXYGEN_HTML_THEME "default")
    set(DOXYGEN_HTML_COLORSTYLE_HUE 220)
    set(DOXYGEN_HTML_COLORSTYLE_SAT 100)
    set(DOXYGEN_HTML_COLORSTYLE_GAMMA 80)
    
    # 中文支持
    set(DOXYGEN_OUTPUT_LANGUAGE "Chinese")
    set(DOXYGEN_INPUT_ENCODING "UTF-8")
    
    # 创建Doxyfile配置文件
    configure_file(
        ${CMAKE_CURRENT_LIST_DIR}/Doxyfile.in
        ${CMAKE_BINARY_DIR}/Doxyfile
        @ONLY
    )
    
    # 创建文档生成目标
    if(EXISTS "${DOXYGEN_SOURCE_DIR}/CMakeLists.txt")
        # 使用项目内Doxygen，需要先构建
        add_custom_target(docs
            COMMAND ${DOXYGEN_EXECUTABLE} ${CMAKE_BINARY_DIR}/Doxyfile
            WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
            DEPENDS doxygen  # 确保Doxygen先被构建
            COMMENT "🔧 生成API文档（使用项目内Doxygen）..."
            VERBATIM
        )
    else()
        # 使用系统Doxygen
        add_custom_target(docs
            COMMAND ${DOXYGEN_EXECUTABLE} ${CMAKE_BINARY_DIR}/Doxyfile
            WORKING_DIRECTORY ${CMAKE_SOURCE_DIR}
            COMMENT "🔧 生成API文档（使用系统Doxygen）..."
            VERBATIM
        )
    endif()
    
    # 创建清理文档目标
    add_custom_target(docs-clean
        COMMAND ${CMAKE_COMMAND} -E remove_directory ${DOXYGEN_OUTPUT_DIR}/html
        COMMAND ${CMAKE_COMMAND} -E remove_directory ${DOXYGEN_OUTPUT_DIR}/xml
        COMMENT "🧹 清理生成的文档..."
    )
    
    # 创建文档服务器目标（可选）
    find_program(PYTHON_EXECUTABLE python3 python)
    if(PYTHON_EXECUTABLE)
        add_custom_target(docs-serve
            COMMAND ${PYTHON_EXECUTABLE} -m http.server 8080
            WORKING_DIRECTORY ${DOXYGEN_OUTPUT_DIR}/html
            COMMENT "🌐 启动文档服务器 http://localhost:8080"
            DEPENDS docs
        )
    endif()
    
    # 安装文档（可选）
    if(CMAKE_BUILD_TYPE STREQUAL "Release")
        install(DIRECTORY ${DOXYGEN_OUTPUT_DIR}/html/
                DESTINATION share/doc/${PROJECT_NAME}
                COMPONENT documentation
                OPTIONAL)
    endif()
    
    message(STATUS "✅ Doxygen文档生成已配置")
    message(STATUS "   生成文档: cmake --build . --target docs")
    message(STATUS "   清理文档: cmake --build . --target docs-clean")
    if(PYTHON_EXECUTABLE)
        message(STATUS "   启动服务: cmake --build . --target docs-serve")
    endif()
    
else()
    message(STATUS "⚠️  未找到Doxygen，跳过文档生成配置")
    message(STATUS "   安装Doxygen: https://www.doxygen.nl/download.html")
    
    # 创建空的文档目标，避免构建错误
    add_custom_target(docs
        COMMAND ${CMAKE_COMMAND} -E echo "Doxygen未安装，无法生成文档"
        COMMENT "⚠️  Doxygen未安装"
    )
endif()

# ==================== 文档工具函数 ====================

# 为目标添加文档注释验证
function(add_documentation_check target_name)
    if(DOXYGEN_FOUND)
        # 创建文档检查目标
        add_custom_target(${target_name}-docs-check
            COMMAND ${DOXYGEN_EXECUTABLE} -g ${CMAKE_BINARY_DIR}/${target_name}_check.cfg
            COMMAND ${CMAKE_COMMAND} -E echo "检查 ${target_name} 的文档注释..."
            COMMENT "📝 检查${target_name}的文档注释"
        )
    endif()
endfunction()

# 生成简单的API索引
function(generate_api_index)
    set(API_INDEX_FILE "${CMAKE_SOURCE_DIR}/api-docs/API_INDEX.md")
    
    file(WRITE ${API_INDEX_FILE}
        "# TinaXlsx API 文档索引\n\n"
        "本文档由CMake自动生成。\n\n"
        "## 📚 文档类型\n\n"
        "### 自动生成文档\n"
        "- **HTML文档**: `api-docs/html/index.html`\n"
        "- **XML文档**: `api-docs/xml/` (用于其他工具)\n\n"
        "### 手动维护文档\n"
        "- **API参考**: [`API_Reference.md`](./API_Reference.md)\n"
        "- **使用指南**: [`../doc/TinaXlsx使用指南.md`](../doc/TinaXlsx使用指南.md)\n\n"
        "## 🔧 生成文档\n\n"
        "```bash\n"
        "# 生成API文档\n"
        "cmake --build cmake-build-debug --target docs\n\n"
        "# 清理文档\n"
        "cmake --build cmake-build-debug --target docs-clean\n\n"
        "# 启动文档服务器\n"
        "cmake --build cmake-build-debug --target docs-serve\n"
        "```\n\n"
        "## 📊 项目信息\n\n"
        "- **项目名称**: ${PROJECT_NAME}\n"
        "- **版本**: ${PROJECT_VERSION}\n"
        "- **生成时间**: $(date)\n"
        "- **CMake版本**: ${CMAKE_VERSION}\n"
    )
    
    message(STATUS "📝 生成API索引: ${API_INDEX_FILE}")
endfunction()

# 调用索引生成
generate_api_index()

message(STATUS "✅ Doxygen配置模块加载完成")
