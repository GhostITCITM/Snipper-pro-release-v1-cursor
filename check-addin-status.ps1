# Check-Addin-Status.ps1 - Diagnostic script for add-in issues

Write-Host "🔍 Snipper Pro Add-in Diagnostic Report" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

# Check registry LoadBehavior
Write-Host "1. Registry Status:" -ForegroundColor Yellow
$regEntry = Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\SnipperPro.Connect" -ErrorAction SilentlyContinue
if ($regEntry) {
    $loadBehavior = $regEntry.LoadBehavior
    Write-Host "   LoadBehavior: $loadBehavior" -ForegroundColor White
    switch ($loadBehavior) {
        0 { Write-Host "   ❌ Status: Unloaded" -ForegroundColor Red }
        1 { Write-Host "   ⚠️  Status: Load on demand" -ForegroundColor Yellow }
        2 { Write-Host "   ❌ Status: Disabled due to error" -ForegroundColor Red }
        3 { Write-Host "   ✅ Status: Should load at startup" -ForegroundColor Green }
    }
} else {
    Write-Host "   ❌ Registry entry not found" -ForegroundColor Red
}

# Check installation files
Write-Host ""
Write-Host "2. Installation Files:" -ForegroundColor Yellow
$installPath = "C:\Users\piete\AppData\Local\SnipperPro\SnipperCloneCleanFinal.dll"
if (Test-Path $installPath) {
    $fileInfo = Get-Item $installPath
    Write-Host "   ✅ DLL exists: $($fileInfo.Length) bytes" -ForegroundColor Green
    Write-Host "   📅 Modified: $($fileInfo.LastWriteTime)" -ForegroundColor White
} else {
    Write-Host "   ❌ DLL not found at expected location" -ForegroundColor Red
}

# Check COM registration
Write-Host ""
Write-Host "3. COM Registration:" -ForegroundColor Yellow
try {
    $comObject = New-Object -ComObject "SnipperPro.Connect"
    Write-Host "   ✅ COM instantiation successful" -ForegroundColor Green
} catch {
    Write-Host "   ❌ COM instantiation failed: $($_.Exception.Message)" -ForegroundColor Red
}

# Check recent errors
Write-Host ""
Write-Host "4. Recent Errors (last 30 minutes):" -ForegroundColor Yellow
$errors = Get-WinEvent -FilterHashtable @{
    LogName='Application'; 
    StartTime=(Get-Date).AddMinutes(-30)
} -ErrorAction SilentlyContinue | Where-Object {
    $_.LevelDisplayName -eq 'Error' -and 
    ($_.Message -like "*SnipperPro*" -or $_.Message -like "*SnipperClone*")
} | Select-Object TimeCreated, Message -First 3

if ($errors) {
    foreach ($error in $errors) {
        Write-Host "   ❌ $($error.TimeCreated): $($error.Message.Substring(0, [Math]::Min(100, $error.Message.Length)))..." -ForegroundColor Red
    }
} else {
    Write-Host "   ✅ No recent errors found" -ForegroundColor Green
}

Write-Host ""
Write-Host "5. Next Steps:" -ForegroundColor Yellow
if ($regEntry.LoadBehavior -eq 2) {
    Write-Host "   ⚠️  Excel disabled the add-in due to an error" -ForegroundColor Yellow
    Write-Host "   🔧 Solution: Fix the error and reset LoadBehavior to 3" -ForegroundColor Cyan
} elseif ($regEntry.LoadBehavior -eq 3) {
    Write-Host "   ✅ Registry looks correct" -ForegroundColor Green
    Write-Host "   🔧 Try: Restart Excel and check COM Add-ins dialog" -ForegroundColor Cyan
}

Write-Host ""
Write-Host "💡 If the add-in appears in COM Add-ins but ribbon doesn't show:" -ForegroundColor Cyan
Write-Host "   - Check if add-in loads without errors (look for SNIPPER PRO tab)" -ForegroundColor White
Write-Host "   - Try running Excel as Administrator once" -ForegroundColor White
Write-Host "   - Check Windows Event Viewer for detailed error messages" -ForegroundColor White 