# PowerShell script to start Angular dev server with external browser
# This prevents Cursor from intercepting the browser opening

Write-Host "Starting Angular development server..." -ForegroundColor Green
Write-Host "The application will open in your external browser at http://localhost:4200" -ForegroundColor Yellow
Write-Host ""

# Start ng serve in background and open browser after delay
Start-Job -ScriptBlock { 
    param($pwd)
    Set-Location $pwd
    ng serve
} -ArgumentList (Get-Location).Path | Out-Null

# Wait for server to start, then open browser externally
Start-Sleep -Seconds 8
Start-Process "http://localhost:4200"

# Keep script running and show ng serve output
Get-Job | Receive-Job -Wait
