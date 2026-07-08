@echo off
chcp 65001 >nul
title Word 表格工具
echo ========================================
echo   Word 表格批量工具
echo ========================================
echo.
echo   1 - 替换工具
echo   2 - 格式刷工具
echo.
set /p choice=请选择 (1/2): 
echo.
if "%choice%"=="1" (
    echo 启动替换工具...
    python word_replacer_tool.py
) else if "%choice%"=="2" (
    echo 启动格式刷工具...
    python format_brush.py
) else (
    echo 无效选择！
)
echo.
echo ========================================
echo   操作完成！
echo ========================================
echo.
pause
