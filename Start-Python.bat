@echo off
setlocal
cd /d "%~dp0"
where py.exe >nul 2>nul
if %errorlevel%==0 (
    py -3 NetworkManager.py
) else (
    python NetworkManager.py
)
if errorlevel 1 pause
