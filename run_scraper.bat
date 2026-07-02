@echo off
setlocal
pushd "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo Python environment not found.
    echo Please run install_requirements.bat first.
    pause
    popd
    exit /b 1
)

if exist ".venv\Scripts\pythonw.exe" (
    set "PYEXE=.venv\Scripts\pythonw.exe"
) else (
    set "PYEXE=.venv\Scripts\python.exe"
)

%PYEXE% pyqt6_menu.py
set "ERR=%ERRORLEVEL%"

popd
exit /b %ERR%
