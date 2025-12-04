@echo off
echo Starting Angular development server...
echo The application will open in your external browser at http://localhost:4200
echo.

REM Start ng serve in a new window and open browser after delay
start "Angular Dev Server" cmd /k "ng serve & timeout /t 8 /nobreak >nul & start http://localhost:4200"

