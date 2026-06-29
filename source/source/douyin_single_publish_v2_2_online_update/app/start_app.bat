@echo off
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
where pyw >nul 2>nul
if not errorlevel 1 (start "" pyw "app_main.pyw"&exit /b 0)
where pythonw >nul 2>nul
if not errorlevel 1 (start "" pythonw "app_main.pyw"&exit /b 0)
if exist "%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe" (start "" "%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe" "app_main.pyw"&exit /b 0)
if exist "%LOCALAPPDATA%\Programs\Python\Python313\pythonw.exe" (start "" "%LOCALAPPDATA%\Programs\Python\Python313\pythonw.exe" "app_main.pyw"&exit /b 0)
echo Python not found.
pause
