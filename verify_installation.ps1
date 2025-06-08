Write-Host "=== Snipper Pro Installation Verification ===" -ForegroundColor Green

# Check DLL file
$dllPath = ".\SnipperCloneCleanFinal\bin\Release\SnipperCloneCleanFinal.dll"
if (Test-Path $dllPath) {
    $fileInfo = Get-Item $dllPath
    Write-Host "✓ DLL File: EXISTS" -ForegroundColor Green
    Write-Host "  Path: $dllPath" -ForegroundColor White
    Write-Host "  Size: $([math]::Round($fileInfo.Length/1KB, 2)) KB" -ForegroundColor White
    Write-Host "  Modified: $($fileInfo.LastWriteTime)" -ForegroundColor White
} else {
    Write-Host "✗ DLL File: MISSING" -ForegroundColor Red
    Write-Host "  Expected: $dllPath" -ForegroundColor White
}

# Check Registry Keys
Write-Host "`n--- Registry Check ---" -ForegroundColor Yellow

try {
    $regPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\SnipperPro"
    $regEntry = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
    
    if ($regEntry) {
        Write-Host "✓ Excel Registry: FOUND" -ForegroundColor Green
        Write-Host "  LoadBehavior: $($regEntry.LoadBehavior)" -ForegroundColor White
        Write-Host "  FriendlyName: $($regEntry.FriendlyName)" -ForegroundColor White
        Write-Host "  Description: $($regEntry.Description)" -ForegroundColor White
    } else {
        Write-Host "✗ Excel Registry: NOT FOUND" -ForegroundColor Red
    }
} catch {
    Write-Host "✗ Registry Check Failed: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`n--- Next Steps ---" -ForegroundColor Cyan
Write-Host "1. Open Microsoft Excel" -ForegroundColor White
Write-Host "2. Look for 'SNIPPER PRO' tab in the ribbon" -ForegroundColor White
Write-Host "3. If not visible: File → Options → Add-ins → COM Add-ins → Check 'Snipper Pro v1'" -ForegroundColor White

Write-Host "`n=== Verification Complete ===" -ForegroundColor Green 