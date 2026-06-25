@echo off
echo ================================================
echo   OMM Optima Dashboard -- Git Re-setup
echo ================================================
echo.

cd /d "C:\Users\Edwin\OneDrive - Studio Snaidero Chicago\OMM_Optima Website"

echo Initializing git repository...
git init
git branch -M main

echo.
echo Setting remote to GitHub...
git remote add origin https://github.com/edwin210designhouse/OMM-OPTIMA-210-DASHBOARD.git

echo.
echo Staging all files...
git add -A

echo.
echo Committing...
git commit -m "Re-setup: new dashboard layout with live data"

echo.
echo Pushing to GitHub (you may see a login popup)...
git push -u origin main --force

echo.
echo ================================================
echo   Done! Site will be live in ~60 seconds.
echo   https://edwin210designhouse.github.io/OMM-OPTIMA-210-DASHBOARD/
echo ================================================
echo.
pause
