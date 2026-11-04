@echo off
title SSBPAC Traceability App & Tunnel
cd /d "%~dp0"

echo ======================================================================
echo  Starting SSBPAC Traceability Suite & Permanent Static Tunnel
echo ======================================================================
echo.

set PYTHON_CMD=python
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set PYTHON_CMD="%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
)

:: Close any old lingering process on port 8080
for /f "tokens=5" %%a in ('netstat -aon ^| findstr /R /C:":8080 .*LISTENING"') do (
    taskkill /F /PID %%a >nul 2>&1
)

echo [1/3] Starting Python web server on port 8080...
start "SSBPAC Python Backend (Do not close)" /min %PYTHON_CMD% server.py 8080

:: Wait 2 seconds for server to initialize
timeout /t 2 /nobreak >nul

echo [2/3] Opening browser...
start http://localhost:8080

echo [3/3] Connecting permanent public tunnel...
echo ======================================================================
echo  SSBPAC TRACEABILITY APP IS RUNNING & CONNECTED!
echo ======================================================================
echo.
echo  Access URLs:
echo    On THIS Laptop:     http://localhost:8080
echo    On OTHER Laptops:   https://unblock-panning-rockstar.ngrok-free.dev
echo.
echo  [IMPORTANT] Keep this window OPEN while working!
echo  (Closing this window will stop the public link).
echo ======================================================================
echo.

"%~dp0ngrok.exe" http --url=https://unblock-panning-rockstar.ngrok-free.dev --authtoken=3LdJcbAxmrk0ULmuzqbrH0KbNOA_6DDzzNKN5jJW3CyGgh66V 8080

pause

