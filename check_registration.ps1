Write-Host "Checking Snipper Pro Registration Status..." -ForegroundColor Green
Write-Host "=========================================" -ForegroundColor Green

# Check if the COM assembly is registered
try {
    $progId = "SnipperPro.Connect"
    Write-Host "Checking ProgID registration for: $progId" -ForegroundColor Yellow
    
    $regKey = Get-ItemProperty -Path "HKEY_CLASSES_ROOT\$progId" -ErrorAction SilentlyContinue
    if ($regKey) {
        Write-Host "✓ ProgID is registered" -ForegroundColor Green
    } else {
        Write-Host "✗ ProgID not found in registry" -ForegroundColor Red
    }
    
    # Check Excel add-ins registry
    $excelAddInsKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\$progId"
    Write-Host "Checking Excel add-in registration: $excelAddInsKey" -ForegroundColor Yellow
    
    $excelReg = Get-ItemProperty -Path "Registry::$excelAddInsKey" -ErrorAction SilentlyContinue
    if ($excelReg) {
        Write-Host "✓ Excel add-in is registered" -ForegroundColor Green
        Write-Host "  LoadBehavior: $($excelReg.LoadBehavior)" -ForegroundColor Cyan
        if ($excelReg.Description) {
            Write-Host "  Description: $($excelReg.Description)" -ForegroundColor Cyan
        }
    } else {
        Write-Host "✗ Excel add-in not found in registry" -ForegroundColor Red
    }
    
    # Check if DLL exists
    $dllPath = Join-Path $PSScriptRoot "SnipperCloneCleanFinal\bin\Release\SnipperCloneCleanFinal.dll"
    Write-Host "Checking DLL file: $dllPath" -ForegroundColor Yellow
    
    if (Test-Path $dllPath) {
        $fileInfo = Get-Item $dllPath
        Write-Host "✓ DLL exists (Size: $([math]::Round($fileInfo.Length/1KB, 2)) KB, Modified: $($fileInfo.LastWriteTime))" -ForegroundColor Green
    } else {
        Write-Host "✗ DLL file not found" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error checking registration: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "Next Steps:" -ForegroundColor Yellow
Write-Host "1. Open Excel" -ForegroundColor White
Write-Host "2. Check if Snipper Pro tab appears in the ribbon" -ForegroundColor White
Write-Host "3. If not visible, go to File > Options > Add-ins > Manage COM Add-ins" -ForegroundColor White
Write-Host "4. Look for Snipper Pro in the list and ensure it is checked" -ForegroundColor White
Write-Host ""
Write-Host "DataSnipper-style Column Snip Features:" -ForegroundColor Cyan
Write-Host "• Table Snip mode with visual +/- column controls" -ForegroundColor White
Write-Host "• Click + icons to add column dividers" -ForegroundColor White
Write-Host "• Click - icons to remove column dividers" -ForegroundColor White
Write-Host "• Drag dividers to adjust column boundaries" -ForegroundColor White
Write-Host "• Double-click to extract data to Excel with proper column structure" -ForegroundColor White

Read-Host "Press Enter to continue..." 