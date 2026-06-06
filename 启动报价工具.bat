@echo off
chcp 65001 >nul
cd /d "%~dp0"
for /f "delims=" %%P in ('where pyw.exe 2^>nul') do (
    start "" "%%P" "%~dp0main.py"
    exit /b 0
)
for /f "delims=" %%P in ('where pythonw.exe 2^>nul') do (
    start "" "%%P" "%~dp0main.py"
    exit /b 0
)
start "" py.exe "%~dp0main.py"
exit /b 0
