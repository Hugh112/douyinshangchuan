@echo off
cd /d "%~dp0"
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8
py -m pip install --upgrade pip setuptools wheel
py -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
set PLAYWRIGHT_DOWNLOAD_HOST=https://npmmirror.com/mirrors/playwright
py -m playwright install chromium
pause
