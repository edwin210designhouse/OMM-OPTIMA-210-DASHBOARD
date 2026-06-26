@echo off
echo ================================================
echo   OMM Optima -- One-Time Setup and First Push
echo ================================================
echo.

echo Step 1: Creating C:\Temp\omm-optima folder...
mkdir "C:\Temp\omm-optima" 2>nul
mkdir "C:\Temp\omm-optima-git" 2>nul

echo Step 2: Copying dashboard to C:\Temp...
copy /y "C:\Users\Edwin\OneDrive - Studio Snaidero Chicago\OMM_Optima Website\OMM_Optima_Dashboard.html" "C:\Temp\omm-optima\OMM_Optima_Dashboard.html"
copy /y "C:\Users\Edwin\OneDrive - Studio Snaidero Chicago\OMM_Optima Website\OMM_Optima_Dashboard.html" "C:\Temp\omm-optima\OMM_Optima_Template.html"

echo.
echo Step 3: Setting up git in C:\Temp (away from OneDrive)...
set GIT_DIR=C:\Temp\omm-optima-git
set GIT_WORK_TREE=C:\Temp\omm-optima

git init
git config user.email "flores.edwin9271@gmail.com"
git config user.name "Edwin"
git branch -M main
git remote remove origin 2>nul
git remote add origin https://github.com/edwin210designhouse/OMM-OPTIMA-210-DASHBOARD.git

echo.
echo Step 4: Committing and pushing...
git add -A
git commit -m "Setup: clean dashboard with live data"
git push -u origin main --force

echo.
echo ================================================
echo   Done! Site will be live in ~60 seconds.
echo   After this, just double-click update_dashboard.py
echo   https://edwin210designhouse.github.io/OMM-OPTIMA-210-DASHBOARD/
echo ================================================
echo.
pause
