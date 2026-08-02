@echo off
cd /d "%~dp0"
echo ===================================================
echo   Naver Address Auto-Registration - Install
echo ===================================================
echo.
echo [1/3] Upgrading pip...
python -m pip install --upgrade pip
echo.
echo [2/3] Installing required packages...
pip install -r requirements.txt
echo.
echo [3/3] Checking Appium installation...
where appium >nul 2>&1
if %errorlevel% neq 0 (
    echo    Appium is not installed.
    echo    Please install it using: npm install -g appium
    echo    And install UiAutomator2 driver: appium driver install uiautomator2
) else (
    echo    Appium is installed.
)
echo.
echo Installation completed!
pause
