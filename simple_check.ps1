Write-Host "Checking Snipper Pro Registration..." -ForegroundColor Green

# Check if DLL exists
$dllPath = Join-Path $PSScriptRoot "SnipperCloneCleanFinal\bin\Release\SnipperCloneCleanFinal.dll"
if (Test-Path $dllPath) {
    $fileInfo = Get-Item $dllPath
    Write-Host "✓ DLL found: $([math]::Round($fileInfo.Length/1KB, 2)) KB" -ForegroundColor Green
} else {
    Write-Host "✗ DLL not found" -ForegroundColor Red
    exit 1
}

# Check Excel add-in registry
$regPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\SnipperPro"
if (Test-Path $regPath) {
    $reg = Get-ItemProperty $regPath
    Write-Host "✓ Excel add-in registered" -ForegroundColor Green
    Write-Host "  LoadBehavior: $($reg.LoadBehavior)" -ForegroundColor Cyan
} else {
    Write-Host "✗ Excel add-in not registered" -ForegroundColor Red
}

Write-Host ""
Write-Host "Status: Registration check complete" -ForegroundColor Yellow 
