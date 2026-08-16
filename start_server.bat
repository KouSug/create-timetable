@echo off
echo ======================================================
echo   TimeTable App & Cloudflare Tunnel Startup Script
echo ======================================================
echo.

echo [1/3] Checking Python...
py --version > nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH!
    echo Please install Python 3.10-3.12 and check 'Add Python to PATH'.
    pause
    exit /b
)

echo [2/3] Installing/Checking required libraries...
py -m pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install requirements. Please check internet connection.
    pause
    exit /b
)

echo [3/3] Starting TimeTable App (Streamlit)...
start /b py -m streamlit run app_latest.py --server.port 8501 --server.headless true

timeout /t 5 > nul

echo ======================================================
echo   Cloudflare Tunnel is starting.
echo   Look for the URL like "https://*.trycloudflare.com" below!
echo.
echo   To stop the server, just close this command window.
echo ======================================================
echo.

.\cloudflared.exe tunnel --url http://127.0.0.1:8501
pause
