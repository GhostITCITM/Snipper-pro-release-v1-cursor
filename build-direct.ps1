#!/usr/bin/env pwsh

Write-Host "Building SnipperClone (Direct PowerShell)" -ForegroundColor Green
Write-Host "=" * 50

# Set MSBuild path
$msbuildPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"

if (!(Test-Path $msbuildPath)) {
    $msbuildPath = "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
}

if (!(Test-Path $msbuildPath)) {
    Write-Host "ERROR: MSBuild not found!" -ForegroundColor Red
    exit 1
}

Write-Host "Using MSBuild: $msbuildPath" -ForegroundColor Yellow

# Clean first
Write-Host "Cleaning previous build..." -ForegroundColor Cyan
& $msbuildPath "SnipperClone\SnipperClone-Simple.csproj" /t:Clean /p:Configuration=Release /p:Platform=AnyCPU /v:minimal

# Build the project
Write-Host "Building SnipperClone..." -ForegroundColor Cyan
$buildResult = & $msbuildPath "SnipperClone\SnipperClone-Simple.csproj" /p:Configuration=Release /p:Platform=AnyCPU /v:normal

if ($LASTEXITCODE -eq 0) {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "BUILD SUCCESSFUL!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Assembly created at: SnipperClone\bin\Release\SnipperClone.dll" -ForegroundColor White
    Write-Host ""
    Write-Host "You can now:" -ForegroundColor Yellow
    Write-Host "1. Install the add-in: .\Install-SnipperClone.ps1" -ForegroundColor White
    Write-Host "2. Build MSI installer: .\Build-MSI.ps1" -ForegroundColor White
    Write-Host ""
} else {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "BUILD FAILED!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "Error code: $LASTEXITCODE" -ForegroundColor Red
    Write-Host ""
    Write-Host "Build output above shows the specific errors." -ForegroundColor Yellow
}

Write-Host "Press Enter to continue..."
Read-Host 