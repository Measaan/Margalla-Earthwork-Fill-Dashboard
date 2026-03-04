@echo off
title Margalla Link Road - Daily Fill Updater
color 0A

echo.
echo  ============================================================
echo   MARGALLA ENCLAVE LINK ROAD  -  DAILY FILL UPDATER
echo   FWO / BK Consultants
echo  ============================================================
echo.

:: ── Check Python ─────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python not installed!
    echo  Download from: https://www.python.org
    echo  During install tick "Add Python to PATH"
    pause & exit
)

:: ── Check openpyxl ───────────────────────────────────
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo  Installing openpyxl...
    pip install openpyxl --quiet
)

:: ── Run the updater ───────────────────────────────────
echo  Running updater...
echo.
python "%~dp0Daily_Update.py"
if errorlevel 1 (
    echo.
    echo  ERROR: Updater failed. See message above.
    pause & exit
)

:: ── Push to GitHub ────────────────────────────────────
echo.
echo  ============================================================
echo   PUSHING TO GITHUB...
echo  ============================================================
echo.

:: Check Git is installed
git --version >nul 2>&1
if errorlevel 1 (
    echo  Git not found. Skipping GitHub upload.
    echo  Install Git from: https://git-scm.com/download/win
    echo  Then run this BAT file again.
    pause & exit
)

cd /d "%~dp0"

:: Init repo if not already done
if not exist ".git" (
    echo  Setting up Git repository for first time...
    git init
    git branch -M main
    git remote add origin https://github.com/Measaan/Margalla-Earthwork-Fill-Dashboard.git
    echo.
    echo  *** FIRST TIME SETUP ***
    echo  Setup complete. Run this BAT file again.
    echo  Then run again.
    pause & exit
)

:: Copy dashboard as index.html for GitHub Pages
copy /Y "%~dp0Margalla_Fill_Dashboard.html" "%~dp0index.html" >nul

:: Stage, commit, push
git add index.html Road_Data.xlsx
git commit -m "Daily update %date% %time%"
git push origin main

if errorlevel 1 (
    echo.
    echo  Push failed. Check your internet connection or GitHub credentials.
    echo  The local files are already saved correctly.
) else (
    echo.
    echo  ============================================================
    echo   DONE! Dashboard live at:
    echo   https://Measaan.github.io/Margalla-Earthwork-Fill-Dashboard/
    echo  ============================================================
)

echo.
pause
