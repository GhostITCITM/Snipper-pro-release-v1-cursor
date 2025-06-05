# Snipper Pro Registration Script
# This script must be run as Administrator

param(
    [switch]$Unregister
)

$ErrorActionPreference = "Stop"

# Paths
$dllPath = "$PSScriptRoot\SnipperCloneCleanFinal\bin\Release\SnipperCloneCleanFinal.dll"
$regAsmPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$regPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\SnipperPro"

Write-Host "Snipper Pro Registration Script" -ForegroundColor Green
Write-Host "===============================" -ForegroundColor Green

# Check if running as Administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")

if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator for COM registration." -ForegroundColor Red
    Write-Host "Please right-click PowerShell and select 'Run as Administrator', then run this script again." -ForegroundColor Yellow
    exit 1
}

# Check if DLL exists
if (-not (Test-Path $dllPath)) {
    Write-Host "ERROR: SnipperCloneCleanFinal.dll not found at: $dllPath" -ForegroundColor Red
    Write-Host "Make sure you have built the project first." -ForegroundColor Yellow
    exit 1
}

if ($Unregister) {
    Write-Host "Unregistering Snipper Pro..." -ForegroundColor Yellow
    
    # Remove Excel registry entries
    if (Test-Path $regPath) {
        Remove-Item $regPath -Recurse -Force
        Write-Host "✓ Removed Excel add-in registry entries" -ForegroundColor Green
    }
    
    # Unregister COM assembly
    try {
        & $regAsmPath $dllPath /unregister
        Write-Host "✓ Unregistered COM assembly" -ForegroundColor Green
    } catch {
        Write-Host "⚠ Warning: Failed to unregister COM assembly: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    
    Write-Host "Snipper Pro has been unregistered." -ForegroundColor Green
} else {
    Write-Host "Registering Snipper Pro..." -ForegroundColor Yellow
    
    # Register COM assembly
    Write-Host "Registering COM assembly..." -ForegroundColor Cyan
    & $regAsmPath $dllPath /codebase
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERROR: Failed to register COM assembly" -ForegroundColor Red
        exit 1
    }
    Write-Host "✓ Registered COM assembly" -ForegroundColor Green
    
    # Set up Excel add-in registry entries
    Write-Host "Setting up Excel add-in registry entries..." -ForegroundColor Cyan
    
    if (Test-Path $regPath) {
        Remove-Item $regPath -Recurse -Force
    }
    
    New-Item $regPath -Force | Out-Null
    Set-ItemProperty $regPath "FriendlyName" "Snipper Pro v1"
    Set-ItemProperty $regPath "Description" "Professional PDF data extraction tool"
    Set-ItemProperty $regPath "LoadBehavior" 3 -Type DWord
    Set-ItemProperty $regPath "ProgId" "SnipperCloneCleanFinal.ThisAddIn"
    
    Write-Host "✓ Created Excel add-in registry entries" -ForegroundColor Green
    
    Write-Host ""
    Write-Host "✅ Snipper Pro has been successfully registered!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor Yellow
    Write-Host "1. Start Microsoft Excel" -ForegroundColor White
    Write-Host "2. You should see a 'SNIPPER PRO' ribbon tab" -ForegroundColor White
    Write-Host "3. If not visible, go to File → Options → Add-Ins → COM Add-ins → Go..." -ForegroundColor White
    Write-Host "4. Check the box next to 'Snipper Pro v1'" -ForegroundColor White
}

Write-Host ""
Write-Host "Script completed." -ForegroundColor Green 