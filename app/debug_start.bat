@echo off
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
set DOUYIN_KEEP_CONSOLE=1

echo ==============================
echo Douyin Publisher Debug Start
echo ==============================
echo.
echo Current directory:
cd
echo.
echo Python version:
py --version
echo.
echo Python path:
py -c "import sys; print(sys.executable)"
echo.
echo Testing dependencies:
py -c "import pandas, openpyxl, xlrd, greenlet; from playwright.sync_api import sync_playwright; print('dependencies ok')"
if errorlevel 1 (
    echo.
    echo Dependency check failed. Please install dependencies first.
    pause
    exit /b 1
)
echo.
echo Starting app_main.py:
py -u "app_main.py"
echo.
echo Program exited. Exit code: %errorlevel%
pause
