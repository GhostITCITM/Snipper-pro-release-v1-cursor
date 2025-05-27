# Launch SnipperClone via sideloaded starter workbook

$docPath = "C:\Users\piete\Desktop\snipper pro v1\sideload\SnipperCloneStarter.xlsx"

Write-Host "🚀 Launching SnipperClone..." -ForegroundColor Green

if (Test-Path $docPath) {
    Start-Process $docPath
    Write-Host "✅ Started Excel with SnipperClone sideloaded"
    Write-Host "Look for 'SnipperClone' tab in the Excel ribbon"
} else {
    Write-Error "Starter workbook not found: $docPath"
    Write-Host "Run .\sideload-setup.ps1 first"
}