@echo off
REM TinaXlsx 简化性能测试运行脚本

echo ========================================
echo TinaXlsx 简化性能测试
echo ========================================
echo.

REM 检查可执行文件是否存在
if not exist "..\..\cmake-build-debug\tests\performance\SimplePerformanceTests.exe" (
    echo 错误: 找不到简化性能测试可执行文件
    echo 请先编译项目: cmake --build cmake-build-debug --target SimplePerformanceTests
    pause
    exit /b 1
)

REM 创建输出目录
if not exist "test_output" mkdir test_output
if not exist "test_output\performance" mkdir test_output\performance

echo 开始执行简化性能测试...
echo 这个测试专注于核心功能，避免复杂依赖
echo.

REM 记录开始时间
echo 测试开始时间: %date% %time%
echo.

REM 运行简化性能测试
..\..\cmake-build-debug\tests\performance\SimplePerformanceTests.exe --gtest_color=no

REM 检查测试结果
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo 简化性能测试完成!
    echo ========================================
    echo.
    echo 测试结果文件位置:
    echo - test_output\performance\*.xlsx
    echo.
    echo 性能分析建议:
    echo 1. 查看控制台输出的性能数据
    echo 2. 检查内存使用趋势
    echo 3. 对比不同数据类型的性能差异
    echo 4. 识别性能瓶颈并优化
    echo.
    echo 下一步:
    echo - 如果简化测试通过，可以尝试运行极致性能测试
    echo - 使用 run_performance_tests.bat 运行完整测试
    echo.
) else (
    echo.
    echo ========================================
    echo 简化性能测试失败! 错误代码: %ERRORLEVEL%
    echo ========================================
    echo.
    echo 请检查:
    echo 1. 编译是否成功
    echo 2. 依赖库是否正确链接
    echo 3. 磁盘空间是否充足
    echo 4. 是否有权限创建文件
    echo.
)

echo 测试结束时间: %date% %time%
echo.
pause
