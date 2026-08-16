@echo off
cd /d %~dp0
echo ======================================================
echo   TimeTable App ^& ngrok Startup Script
echo ======================================================
echo.

:: 1. Load settings from ngrok_settings.txt
if not exist ngrok_settings.txt (
    echo [ERROR] ngrok_settings.txt not found!
    echo Please make sure the file is in the same folder.
    pause
    exit /b
)

for /f "usebackq tokens=1,2 delims==" %%i in ("ngrok_settings.txt") do (
    if "%%i"=="NGROK_AUTHTOKEN" set AUTHTOKEN=%%j
    if "%%i"=="NGROK_DOMAIN" set DOMAIN=%%j
)

:: Validate variables
if "%AUTHTOKEN%"=="" (
    echo [ERROR] NGROK_AUTHTOKEN is empty in ngrok_settings.txt!
    pause
    exit /b
)
if "%AUTHTOKEN%"=="YOUR_TOKEN_HERE" (
    echo [ERROR] Please replace YOUR_TOKEN_HERE with your real ngrok authtoken in ngrok_settings.txt!
    pause
    exit /b
)
if "%DOMAIN%"=="" (
    echo [ERROR] NGROK_DOMAIN is empty in ngrok_settings.txt!
    pause
    exit /b
)
if "%DOMAIN%"=="YOUR_DOMAIN_HERE.ngrok-free.app" (
    echo [ERROR] Please replace YOUR_DOMAIN_HERE.ngrok-free.app with your real ngrok domain in ngrok_settings.txt!
    pause
    exit /b
)

:: 2. Check Python
echo [1/4] Checking Python...
py --version > nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed or not in PATH!
    pause
    exit /b
)

:: 3. Check requirements
echo [2/4] Checking required libraries...
py -m pip install -r requirements.txt > nul 2>&1

:: 4. Start Streamlit server in background
echo [3/4] Starting TimeTable App (Streamlit)...
start /b py -m streamlit run app_latest.py --server.port 8501 --server.headless true

timeout /t 5 > nul

:: 5. Configure ngrok token
echo [4/4] Configuring ngrok authtoken...
.\ngrok.exe config add-authtoken %AUTHTOKEN% > nul 2>&1

:: 6. Launch ngrok tunnel
echo.
echo ======================================================
echo   ngrok tunnel is starting.
echo   Your app will be accessible at:
echo   https://%DOMAIN%
echo.
echo   To stop the server, just close this window.
echo ======================================================
echo.

.\ngrok.exe http --domain=%DOMAIN% 8501
pause
