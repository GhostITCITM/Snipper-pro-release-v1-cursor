Write-Host "Installing Snipper Pro COM Add-in..." -ForegroundColor Green
try {
    $dll = "$env:LOCALAPPDATA\SnipperPro\SnipperClone.dll"
    $regasm = "${env:windir}\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
    Copy-Item "SnipperCloneCleanFinal\bin\x86\Release\SnipperCloneCleanFinal.dll" $dll -Force
    & $regasm $dll /codebase /verbose
    Write-Host "SUCCESS! COM registration completed" -ForegroundColor Green
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
