# Install-SnipperCloneCleanFinal.ps1
# User-level installation for SnipperCloneCleanFinal Excel Add-in

param(
    [string]$InstallPath = "$env:USERPROFILE\AppData\Local\SnipperCloneCleanFinal"
)

Write-Host "Installing SnipperCloneCleanFinal Excel Add-in (user-level installation)..." -ForegroundColor Green

try {
    # Create installation directory
    if (!(Test-Path $InstallPath)) {
        New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
        Write-Host "Created installation directory: $InstallPath"
    }

    # Copy files from build output
    $sourceDir = "SnipperCloneCleanFinal\bin\x64\Release"
    if (!(Test-Path $sourceDir)) {
        Write-Error "Build output not found at $sourceDir. Please build the project first."
        exit 1
    }

    Write-Host "Copying files to installation directory..."
    Copy-Item "$sourceDir\*" -Destination $InstallPath -Recurse -Force

    # Unblock the DLL file to prevent security warnings
    $dllPath = Join-Path $InstallPath "SnipperCloneCleanFinal.dll"
    if (Test-Path $dllPath) {
        Unblock-File $dllPath
        Write-Host "Unblocked DLL file to prevent security warnings"
    }

    # Register the add-in using user-level automation add-in approach
    Write-Host "Registering add-in in Excel user registry..."
    
    # Use Automation Add-ins instead of COM Add-ins (no admin required)
    $registryPath = 'HKCU:\Software\Microsoft\Office\Excel\Addins\SnipperCloneCleanFinal'
    
    # Remove existing registration if it exists
    if (Test-Path $registryPath) {
        Remove-Item $registryPath -Force -Recurse
    }
    
    # Create new registration for Excel Automation Add-in
    New-Item $registryPath -Force | Out-Null
    Set-ItemProperty $registryPath -Name "Description" -Value "Snipper Pro - Document Analysis Excel Add-in"
    Set-ItemProperty $registryPath -Name "FriendlyName" -Value "Snipper Pro"
    Set-ItemProperty $registryPath -Name "LoadBehavior" -Value 3 -Type DWord
    Set-ItemProperty $registryPath -Name "Manifest" -Value $dllPath
    
    Write-Host "Excel registry entries created successfully!" -ForegroundColor Green
    
    # Also try XLL registration path as alternative
    $xllRegistryPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
    if (!(Test-Path $xllRegistryPath)) {
        New-Item $xllRegistryPath -Force | Out-Null
    }
    
    # Enable VBA if needed for certain add-in features
    $securityPath = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    if (!(Test-Path $securityPath)) {
        New-Item $securityPath -Force | Out-Null
    }
    Set-ItemProperty $securityPath -Name "VBAWarnings" -Value 1 -Type DWord -ErrorAction SilentlyContinue
    
    # Enable logging for troubleshooting
    Write-Host "Enabling comprehensive logging..."
    [Environment]::SetEnvironmentVariable("VSTO_LOGALERTS", "1", "User")
    [Environment]::SetEnvironmentVariable("VSTO_SUPPRESSDISPLAYALERTS", "0", "User")
    
    Write-Host ""
    Write-Host "Installation completed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "The add-in has been installed as a user-level Excel add-in." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "To activate the add-in:" -ForegroundColor Cyan
    Write-Host "1. Open Excel"
    Write-Host "2. Go to File > Options > Add-ins"
    Write-Host "3. At the bottom, select 'Excel Add-ins' from the dropdown and click 'Go...'"
    Write-Host "4. Click 'Browse...' and navigate to:"
    Write-Host "   $dllPath"
    Write-Host "5. Select the DLL file and click OK"
    Write-Host "6. Check the checkbox next to 'Snipper Pro' and click OK"
    Write-Host ""
    Write-Host "Alternative method (Automation Add-ins):" -ForegroundColor Cyan
    Write-Host "1. Go to File > Options > Add-ins"
    Write-Host "2. Select 'COM Add-ins' from the dropdown and click 'Go...'"
    Write-Host "3. The add-in should appear as 'SnipperCloneCleanFinal'"
    Write-Host "4. Check the checkbox and click OK"
    Write-Host ""
    Write-Host "If neither method works, try:" -ForegroundColor Yellow
    Write-Host "- Running Excel as Administrator once"
    Write-Host "- Checking Windows Event Viewer for error messages"
    Write-Host "- Verifying no antivirus software is blocking the DLL"
    Write-Host ""
    Write-Host "Installation path: $InstallPath" -ForegroundColor Green
    Write-Host "Registry path: $registryPath" -ForegroundColor Green
    
    # Create a simple batch file for manual registration if needed
    $batContent = @"
@echo off
echo Manual registration script for SnipperCloneCleanFinal
echo.
echo Run this as Administrator if automatic registration failed
echo.
pause
cd /d "$InstallPath"
"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" "SnipperCloneCleanFinal.dll" /codebase /verbose
pause
"@
    $batPath = Join-Path $InstallPath "manual-register.bat"
    $batContent | Out-File -FilePath $batPath -Encoding ASCII
    Write-Host "Created manual registration script: $batPath" -ForegroundColor Gray
    
} catch {
    Write-Error "Installation failed: $($_.Exception.Message)"
    Write-Host "Stack trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    exit 1
} 