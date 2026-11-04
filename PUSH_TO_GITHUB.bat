@echo off
title Push to GitHub - SSBPAC Traceability
cd /d "%~dp0"

echo ======================================================================
echo  Pushing Traceability Suite to GitHub:
echo  https://github.com/9353967513/ssbpac-traceability
echo ======================================================================
echo.

git remote set-url origin https://github.com/9353967513/ssbpac-traceability.git
git add server.py generator_core.py requirements.txt Procfile .gitignore README.md
git commit -m "Deploy SSBPAC Traceability Suite" >nul 2>&1

echo Uploading files via Git...
echo (If a GitHub sign-in window opens, click "Sign in with your browser")
echo.

git push -u origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ======================================================================
    echo  [SUCCESS] All files uploaded to GitHub!
    echo  Check at: https://github.com/9353967513/ssbpac-traceability
    echo ======================================================================
    echo.
    echo Next step: Go to https://dashboard.render.com to deploy it!
) else (
    echo.
    echo ======================================================================
    echo  [FAILED] Git could not push. 
    echo  You can also drag and drop the files directly on GitHub:
    echo  https://github.com/9353967513/ssbpac-traceability
    echo ======================================================================
)

echo.
pause

