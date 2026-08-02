@echo off
cd /d "%~dp0"

echo ============================================
echo  Naver Auto Address Registration - Startup
echo ============================================

echo [1/4] Killing ALL node.exe (Appium) processes...
taskkill /F /IM node.exe /T >nul 2>&1
echo  Done.

echo [2/4] Killing processes on known Appium ports...
for %%P in (5720 5721 5722 5723 5724 5725 5730 5733 5743 5753 5763 9200 9201 9202 9203 9204 9205 9210) do (
    for /f "tokens=5" %%i in ('netstat -ano 2^>nul ^| findstr " :%%P "') do (
        if not "%%i"=="0" taskkill /F /PID %%i >nul 2>&1
    )
)
echo  Done.

echo [3/4] Clearing Python cache...
if exist __pycache__ rmdir /s /q __pycache__ >nul 2>&1
echo  Done.

echo [4/5] Checking required Python libraries...
python -c "import easyocr, cv2, appium, PIL, ddddocr" >nul 2>&1
if %errorlevel% neq 0 (
    echo   [INFO] Some required libraries are missing. Installing: easyocr, opencv-python, appium-python-client, Pillow, ddddocr...
    pip install easyocr opencv-python appium-python-client Pillow ddddocr
) else (
    echo   All required libraries are already installed.
)
echo.

echo [5/5] Starting program...
echo.
python main.py
if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Failed. Run install.bat first.
    pause
)
