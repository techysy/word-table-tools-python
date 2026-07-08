@echo off
chcp 65001 >nul
title Word 表格批量替换工具
echo ========================================
echo   Word 表格批量替换工具
echo ========================================
echo.
python word_replacer_tool.py
echo.
echo ========================================
echo   操作完成！日志文件: logs\replace_%date:~0,4%%date:~5,2%%date:~8,2%.log
echo ========================================
echo.
pause
