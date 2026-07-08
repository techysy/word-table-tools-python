@echo off
chcp 65001 >nul
echo 1 - Word Replacer
echo 2 - Format Brush
echo.
set /p choice=Select (1/2): 
if "%choice%"=="1" python word_replacer_tool.py
if "%choice%"=="2" python format_brush.py
pause