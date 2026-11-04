@echo off
title Push to GitHub - SSBPAC Traceability
cd /d "%~dp0"

echo ======================================================================
echo  Pushing Traceability Suite to GitHub:
echo  https://github.com/Vivek-bhyrappa/ssbpac-traceability
echo ======================================================================
echo.

git remote set-url origin https://github.com/Vivek-bhyrappa/ssbpac-traceability.git
git add server.py generator_core.py requirements.txt Procfile .gitignore README.md test_suite.py .github/
git commit -m "Update SSBPAC Traceability Suite" >nul 2>&1

echo Uploading files via Git...
echo (If a GitHub sign-in window opens, click "Sign in with your browser")
echo.

git push -u origin main

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ======================================================================
    echo  [SUCCESS] All files uploaded to GitHub!
    echo  Check at: https://github.com/Vivek-bhyrappa/ssbpac-traceability
    echo ======================================================================
) else (
    echo.
    echo ======================================================================
    echo  If push was rejected because the remote has commits, pulling with rebase...
    echo ======================================================================
    git pull origin main --rebase
    git push -u origin main
)

echo.
pause
