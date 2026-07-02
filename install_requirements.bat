@echo off
setlocal
pushd "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
    if errorlevel 1 goto :fail
) else (
    echo Reusing existing virtual environment.
)

".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :fail

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :fail

echo.
echo Installation finished.
popd
exit /b 0

:fail
echo.
echo Installation failed.
popd
exit /b 1
