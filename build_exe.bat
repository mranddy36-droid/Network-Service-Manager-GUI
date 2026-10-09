@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>&1
if not errorlevel 1 (
    py -m PyInstaller --noconfirm --onefile --noconsole NetworkManager.py
    goto :build_result
)

where python >nul 2>&1
if not errorlevel 1 (
    python -m PyInstaller --noconfirm --onefile --noconsole NetworkManager.py
    goto :build_result
)

echo Python не найден. Установите Python и добавьте его в PATH.
pause
exit /b 1

:build_result
if errorlevel 1 (
    echo Сборка не выполнена. Проверьте, установлен ли PyInstaller: py -m pip install pyinstaller
    pause
    exit /b 1
)

echo Готово: %~dp0dist\NetworkManager.exe
pause
endlocal
