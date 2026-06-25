@echo off
cd /d "C:\Users\Edwin\OneDrive - Studio Snaidero Chicago\OMM_Optima Website"
echo Pushing to GitHub...
git add -A
git commit -m "Update dashboard"
git push
echo.
echo Done! Site updates in ~60 seconds.
pause
