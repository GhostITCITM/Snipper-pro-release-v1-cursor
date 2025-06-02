# Close Excel
Write-Host "Closing Excel..." -ForegroundColor Yellow
Stop-Process -Name EXCEL -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# Build
Write-Host "Building Snipper Pro..." -ForegroundColor Green
.\build-snipper-pro.ps1

# Install
Write-Host "Installing Snipper Pro..." -ForegroundColor Green
.\install-snipper-pro-complete.ps1

# Register
Write-Host "Registering COM component..." -ForegroundColor Green
.\REGISTER_NOW.bat

Write-Host "Done! Start Excel to test." -ForegroundColor Cyan 