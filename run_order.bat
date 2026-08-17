@echo off
cd /d "%~dp0"

echo ============================================
echo   Naver Auto Order - Startup
echo ============================================
echo.

echo [1/5] Killing previous Appium node processes...
taskkill /F /IM node.exe /T >nul 2>&1
echo  Done.

echo [2/5] Killing processes on known Appium ports...
for %%P in (7723 7724 7725 7726 7727 7728 7729 7730 7731 7732 7733 7734 7735 7736 7737 7738 7739 7740 7741 7742) do (
    for /f "tokens=5" %%i in ('netstat -ano 2^>nul ^| findstr " :%%P "') do (
        if not "%%i"=="0" taskkill /F /PID %%i >nul 2>&1
    )
)
echo  Done.

echo [3/5] Clearing Python cache...
if exist __pycache__ rmdir /s /q __pycache__ >nul 2>&1
echo  Done.

echo [4/5] Finding Python...
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
if "%PYTHON_EXE%"=="" if exist "C:\Program Files\Python310\python.exe" set "PYTHON_EXE=C:\Program Files\Python310\python.exe"
if "%PYTHON_EXE%"=="" if exist "C:\Program Files\Python311\python.exe" set "PYTHON_EXE=C:\Program Files\Python311\python.exe"
if "%PYTHON_EXE%"=="" if exist "C:\Program Files\Python312\python.exe" set "PYTHON_EXE=C:\Program Files\Python312\python.exe"

if "%PYTHON_EXE%"=="" set "PYTHON_EXE=python"

echo   Python: %PYTHON_EXE%
echo  Done.

echo [5/5] Checking required packages...
"%PYTHON_EXE%" -c "import openpyxl, cv2, PIL, pytesseract, paddleocr" >nul 2>&1
if %errorlevel% neq 0 (
    echo   Installing required packages...
    "%PYTHON_EXE%" -m pip install openpyxl opencv-python Pillow appium-python-client pytesseract paddleocr
    echo   Done.
) else (
    echo   All packages are ready.
)
echo.

echo Starting Naver Auto Order Program...
echo.
"%PYTHON_EXE%" main_order.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed to start.
    echo Run: "%PYTHON_EXE%" -m pip install openpyxl opencv-python Pillow appium-python-client
    echo.
    pause
)
