@echo off
chcp 65001 >nul
cd /d "%~dp0"
py "관로번호_검사프로그램.py"
if errorlevel 1 pause
