@echo off
echo ================================================
echo   OMM Optima -- FULL RESET AND PUSH
echo ================================================
echo.

cd /d "C:\Users\Edwin\OneDrive - Studio Snaidero Chicago\OMM_Optima Website"

echo Step 1: Force-removing .git folder...
takeown /f ".git" /r /d y >nul 2>&1
icacls ".git" /grant Everyone:F /t >nul 2>&1
rd /s /q ".git" 2>nul
if exist ".git" (
    echo   WARNING: Could not remove .git - trying anyway...
) else (
    echo   Removed cleanly.
)

echo.
echo Step 2: Fresh git init...
git init
git config user.email "flores.edwin9271@gmail.com"
git config user.name "Edwin"
git branch -M main

echo.
echo Step 3: Setting GitHub remote...
git remote add origin https://github.com/edwin210designhouse/OMM-OPTIMA-210-DASHBOARD.git

echo.
echo Step 4: Staging all files...
git add -A

echo.
echo Step 5: Committing...
git commit -m "Fix: data now renders correctly"

echo.
echo Step 6: Force pushing to GitHub...
git push -u origin main --force

echo.
if %errorlevel%==0 (
    echo ================================================
    echo   SUCCESS! Site live in ~60 seconds.
    echo   https://edwin210designhouse.github.io/OMM-OPTIMA-210-DASHBOARD/
    echo ================================================
) else (
    echo ================================================
    echo   PUSH FAILED.
    echo   Try the manual upload instead:
    echo   1. Go to github.com/edwin210designhouse/OMM-OPTIMA-210-DASHBOARD
    echo   2. Click "Add file" then "Upload files"
    echo   3. Drag in OMM_Optima_Dashboard.html
    echo   4. Click "Commit changes"
    echo ================================================
)
echo.
pause
