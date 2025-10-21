@echo off
echo ========================================
echo   HiMarine Invoicing - Deploy to Firebase
echo ========================================
echo.
echo Building production version...
echo.

call npm run build

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Build failed!
    echo Please fix the errors and try again.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build successful! Deploying to Firebase...
echo ========================================
echo.

call firebase deploy --only hosting

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Deployment failed!
    echo Please check your Firebase configuration.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Deployment Complete!
echo ========================================
echo.
echo Your website is now live at:
echo https://himarine-invoicing.web.app
echo.
pause

