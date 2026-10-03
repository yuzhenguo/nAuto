@echo off
cd /d "%~dp0"

echo ============================================
echo   Naver Auto Review - Startup
echo ============================================
echo.

echo [1/3] Clearing Python cache...
if exist __pycache__ rmdir /s /q __pycache__ >nul 2>&1
if exist ..\__pycache__ rmdir /s /q ..\__pycache__ >nul 2>&1
echo  Done.

echo [2/3] Finding Python...
set PYTHON_EXE=
for %%V in (313 312 311 310 39 38) do (
    if "%PYTHON_EXE%"=="" if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
        set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
    )
)
for %%V in (313 312 311 310 39 38) do (
    if "%PYTHON_EXE%"=="" if exist "C:\Python%%V\python.exe" (
        set "PYTHON_EXE=C:\Python%%V\python.exe"
    )
)
if "%PYTHON_EXE%"=="" set "PYTHON_EXE=python"
echo   Python: %PYTHON_EXE%

echo [3/3] Checking required packages...
"%PYTHON_EXE%" -c "import openpyxl, cv2, numpy, appium" >nul 2>&1
if %errorlevel% neq 0 (
    echo   Installing required packages...
    "%PYTHON_EXE%" -m pip install openpyxl opencv-python numpy Pillow appium-python-client
)
echo.

echo Starting Naver Auto Review Program...
"%PYTHON_EXE%" main_review.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to start.
    pause
)
