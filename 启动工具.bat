@echo off
chcp 65001 >nul
title Word Table Tools
echo 1 - Replacer
echo 2 - Format Brush
echo.
set /p choice=Select (1/2): 
if "%choice%"=="1" python word_replacer_tool.py
if "%choice%"=="2" python format_brush.py
echo.
echo Done!
echo.
pause
