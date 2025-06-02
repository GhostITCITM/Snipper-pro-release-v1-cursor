# Force Register Snipper Pro
Write-Host "Force registering Snipper Pro..." -ForegroundColor Yellow

# Stop Excel
Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

# Register COM component
$dllPath = Join-Path $PSScriptRoot "SnipperCloneCleanFinal\bin\Release\SnipperCloneCleanFinal.dll"
Write-Host "Registering: $dllPath"

try {
    & "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" "$dllPath" /codebase /verbose
    Write-Host "✓ COM registration completed" -ForegroundColor Green
} catch {
    Write-Host "✗ COM registration failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Create Excel registry entries
Write-Host "Creating Excel registry entries..." -ForegroundColor Yellow

$regPath = "HKLM:\SOFTWARE\Microsoft\Office\Excel\Addins\SnipperPro.Connect"

try {
    # Remove existing
    if (Test-Path $regPath) {
        Remove-Item -Path $regPath -Recurse -Force
    }
    
    # Create new
    New-Item -Path $regPath -Force | Out-Null
    Set-ItemProperty -Path $regPath -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty -Path $regPath -Name "FriendlyName" -Value "Snipper Pro Excel Add-in" -Type String
    Set-ItemProperty -Path $regPath -Name "Description" -Value "DataSnipper Clone for Excel" -Type String
    
    Write-Host "✓ Registry entries created" -ForegroundColor Green
} catch {
    Write-Host "✗ Registry creation failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "✅ SNIPPER PRO FORCE REGISTERED!" -ForegroundColor Green
Write-Host "Start Excel and look for SNIPPER PRO tab" -ForegroundColor Yellow

Read-Host "Press Enter to start Excel"
Start-Process "excel" 