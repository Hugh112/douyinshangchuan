@echo off
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
py "app_main.py"
if errorlevel 1 python "app_main.py"
pause
