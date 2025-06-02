@echo off
REM TinaXlsx 性能测试运行脚本

echo ========================================
echo TinaXlsx 极致性能测试
echo ========================================
echo.

REM 检查可执行文件是否存在
if not exist "..\..\cmake-build-debug\tests\performance\ExtremePerformanceTests.exe" (
    echo 错误: 找不到性能测试可执行文件
    echo 请先编译项目: cmake --build cmake-build-debug --target ExtremePerformanceTests
    pause
    exit /b 1
)

REM 创建输出目录
if not exist "test_output" mkdir test_output
if not exist "test_output\performance" mkdir test_output\performance

echo 开始执行性能测试...
echo 注意: 这可能需要几分钟时间，请耐心等待
echo.

REM 记录开始时间
echo 测试开始时间: %date% %time%
echo.

REM 运行性能测试
..\..\cmake-build-debug\tests\performance\ExtremePerformanceTests.exe --gtest_color=no

REM 检查测试结果
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo 性能测试完成!
    echo ========================================
    echo.
    echo 测试结果文件位置:
    echo - test_output\performance\*.xlsx
    echo.
    echo 建议的后续操作:
    echo 1. 查看生成的Excel文件验证功能正确性
    echo 2. 分析控制台输出的性能数据
    echo 3. 根据性能报告优化代码
    echo.
) else (
    echo.
    echo ========================================
    echo 性能测试失败! 错误代码: %ERRORLEVEL%
    echo ========================================
    echo.
    echo 请检查:
    echo 1. 编译是否成功
    echo 2. 依赖库是否正确链接
    echo 3. 磁盘空间是否充足
    echo.
)

echo 测试结束时间: %date% %time%
echo.
pause
