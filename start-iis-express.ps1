# Start IIS Express with HTTPS
$port = 8443
$path = Join-Path $PWD "dist\app"

# Check if IIS Express is installed
$iisExpressPath = "${env:ProgramFiles(x86)}\IIS Express\iisexpress.exe"
if (-not (Test-Path $iisExpressPath)) {
    $iisExpressPath = "${env:ProgramFiles}\IIS Express\iisexpress.exe"
}

if (-not (Test-Path $iisExpressPath)) {
    Write-Host "IIS Express not found. Please install it from:" -ForegroundColor Red
    Write-Host "https://www.microsoft.com/en-us/download/details.aspx?id=48264"
    exit 1
}

Write-Host "Starting IIS Express on https://localhost:$port" -ForegroundColor Green
Write-Host "Serving files from: $path" -ForegroundColor Green
Write-Host ""
Write-Host "IMPORTANT: After server starts:" -ForegroundColor Yellow
Write-Host "1. Open https://localhost:8443 in your browser" -ForegroundColor Yellow
Write-Host "2. Accept the security warning" -ForegroundColor Yellow
Write-Host "3. Upload manifest-https.xml to M365 Admin Center" -ForegroundColor Yellow
Write-Host ""
Write-Host "Press Ctrl+C to stop the server" -ForegroundColor Cyan

# Start IIS Express
& $iisExpressPath /path:$path /port:$port /systray:false 