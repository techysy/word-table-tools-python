@echo off
chcp 65001 >nul
title Word Table Replacer
python word_replacer_tool.py
echo.
echo Done! Log: logs\replace_%date:~0,4%%date:~5,2%%date:~8,2%.log
echo.
pause
