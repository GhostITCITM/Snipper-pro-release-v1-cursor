param(
    [switch]$Force = $false
)

Write-Host "=== Snipper Pro Complete Installation ===" -ForegroundColor Green
Write-Host "Installing fully functional Snipper Pro Excel Add-in..." -ForegroundColor Yellow

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Host "ERROR: This script must be run as Administrator!" -ForegroundColor Red
    Write-Host "Please right-click and select 'Run as Administrator'" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

try {
    # Stop Excel if running
    Write-Host "Stopping Excel processes..." -ForegroundColor Yellow
    try {
        Stop-Process -Name EXCEL -Force -Wait
        Write-Host "Excel closed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Excel was not running or already closed" -ForegroundColor Yellow
    }
    Start-Sleep -Seconds 3

    # Verify build
    $dllPath = Join-Path $PSScriptRoot "SnipperCloneCleanFinal\bin\Release\SnipperCloneCleanFinal.dll"
    if (-not (Test-Path $dllPath)) {
        Write-Host "ERROR: DLL not found. Building project first..." -ForegroundColor Red
        
        # Build the project
        Set-Location "SnipperCloneCleanFinal"
        $msbuildPath = "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
        if (-not (Test-Path $msbuildPath)) {
            $msbuildPath = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
        }
        
        if (Test-Path $msbuildPath) {
            Write-Host "Building with MSBuild..." -ForegroundColor Yellow
            & "$msbuildPath" "SnipperCloneCleanFinal.csproj" /p:Configuration=Release /p:Platform="Any CPU"
            if ($LASTEXITCODE -ne 0) {
                Write-Host "ERROR: Build failed!" -ForegroundColor Red
                exit 1
            }
        } else {
            Write-Host "ERROR: MSBuild not found!" -ForegroundColor Red
            exit 1
        }
        
        Set-Location ..
    }

    # Verify DLL exists after build
    if (-not (Test-Path $dllPath)) {
        Write-Host "ERROR: DLL still not found after build!" -ForegroundColor Red
        exit 1
    }

    Write-Host "✓ DLL found: $dllPath" -ForegroundColor Green

    # Create Program Files directory
    $programFilesDir = "${env:ProgramFiles}\SnipperPro"
    Write-Host "Creating installation directory: $programFilesDir" -ForegroundColor Yellow
    
    if (Test-Path $programFilesDir) {
        Remove-Item -Path $programFilesDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $programFilesDir -Force | Out-Null

    # Copy all necessary files
    Write-Host "Copying files to Program Files..." -ForegroundColor Yellow
    
    # Copy main DLL
    Copy-Item -Path $dllPath -Destination $programFilesDir -Force
    Write-Host "✓ Copied main DLL" -ForegroundColor Green

    # Copy Assets folder if it exists
    $assetsSource = Join-Path $PSScriptRoot "SnipperCloneCleanFinal\Assets"
    $assetsDest = Join-Path $programFilesDir "Assets"
    if (Test-Path $assetsSource) {
        Copy-Item -Path $assetsSource -Destination $assetsDest -Recurse -Force
        Write-Host "✓ Copied Assets folder" -ForegroundColor Green
    }

    # Copy dependencies if they exist
    $binDir = Join-Path $PSScriptRoot "SnipperCloneCleanFinal\bin\Release"
    $dependencies = @("*.dll", "*.pdb", "*.config")
    
    foreach ($dep in $dependencies) {
        $files = Get-ChildItem -Path $binDir -Filter $dep -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            Copy-Item -Path $file.FullName -Destination $programFilesDir -Force
            Write-Host "✓ Copied $($file.Name)" -ForegroundColor Green
        }
    }

    # Specifically ensure pdfium.dll is copied
    $pdfiumPath = Join-Path $binDir "pdfium.dll"
    if (Test-Path $pdfiumPath) {
        Copy-Item -Path $pdfiumPath -Destination $programFilesDir -Force
        Write-Host "✓ Copied pdfium.dll (native PDF renderer)" -ForegroundColor Green
    } else {
        Write-Host "⚠ WARNING: pdfium.dll not found - PDF rendering will not work!" -ForegroundColor Red
    }

    # Copy tessdata folder for OCR
    $tessSource = Join-Path $PSScriptRoot "SnipperCloneCleanFinal\tessdata"
    $tessDest = Join-Path $programFilesDir "tessdata"
    if (Test-Path $tessSource) {
        Copy-Item -Path $tessSource -Destination $tessDest -Recurse -Force
        Write-Host "✓ Copied tessdata folder" -ForegroundColor Green
    }

    # Update DLL path to Program Files location
    $installDllPath = Join-Path $programFilesDir "SnipperCloneCleanFinal.dll"

    # Register COM component
    Write-Host "Registering COM component..." -ForegroundColor Yellow
    
    $regasmPath = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
    if (-not (Test-Path $regasmPath)) {
        Write-Host "ERROR: RegAsm.exe not found at $regasmPath" -ForegroundColor Red
        exit 1
    }

    try {
        & $regasmPath "$installDllPath" /codebase
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ COM component registered successfully" -ForegroundColor Green
        } else {
            throw "RegAsm failed with exit code $LASTEXITCODE"
        }
    }
    catch {
        Write-Host "ERROR: COM registration failed: $_" -ForegroundColor Red
        exit 1
    }

    # Create registry entries for Excel add-in
    Write-Host "Creating Excel add-in registry entries..." -ForegroundColor Yellow
    
    $addinRegPath = "HKLM:\SOFTWARE\Microsoft\Office\Excel\Addins\SnipperPro.Connect"
    
    # Remove existing entries first
    if (Test-Path $addinRegPath) {
        Remove-Item -Path $addinRegPath -Recurse -Force
    }
    
    # Create new registry entries
    New-Item -Path $addinRegPath -Force | Out-Null
    New-ItemProperty -Path $addinRegPath -Name "Description" -Value "Snipper Pro Excel Add-in - DataSnipper Clone" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $addinRegPath -Name "FriendlyName" -Value "Snipper Pro" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $addinRegPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $addinRegPath -Name "Manifest" -Value $installDllPath -PropertyType String -Force | Out-Null
    
    Write-Host "✓ Registry entries created" -ForegroundColor Green

    # Also create entries for current user (backup)
    $userAddinRegPath = "HKCU:\SOFTWARE\Microsoft\Office\Excel\Addins\SnipperPro.Connect"
    if (Test-Path $userAddinRegPath) {
        Remove-Item -Path $userAddinRegPath -Recurse -Force
    }
    
    New-Item -Path $userAddinRegPath -Force | Out-Null
    New-ItemProperty -Path $userAddinRegPath -Name "Description" -Value "Snipper Pro Excel Add-in - DataSnipper Clone" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $userAddinRegPath -Name "FriendlyName" -Value "Snipper Pro" -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $userAddinRegPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $userAddinRegPath -Name "Manifest" -Value $installDllPath -PropertyType String -Force | Out-Null

    # Set security permissions for the installation directory
    Write-Host "Setting security permissions..." -ForegroundColor Yellow
    
    try {
        $acl = Get-Acl $programFilesDir
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Users", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow")
        $acl.SetAccessRule($accessRule)
        Set-Acl -Path $programFilesDir -AclObject $acl
        Write-Host "✓ Security permissions set" -ForegroundColor Green
    }
    catch {
        Write-Host "⚠ Could not set security permissions (non-critical)" -ForegroundColor Yellow
    }

    # Create a test script to verify installation
    $testScript = @"
Write-Host "Testing Snipper Pro Installation..." -ForegroundColor Yellow

try {
    `$excel = New-Object -ComObject Excel.Application
    `$excel.Visible = `$true
    
    Start-Sleep -Seconds 3
    
    # Check if add-in is loaded
    `$addin = `$excel.COMAddIns | Where-Object { `$_.ProgID -eq "SnipperPro.Connect" }
    if (`$addin) {
        if (`$addin.Connect) {
            Write-Host "✓ Snipper Pro add-in is loaded and connected!" -ForegroundColor Green
            Write-Host "Look for the 'SNIPPER PRO' tab in Excel ribbon" -ForegroundColor Yellow
        } else {
            Write-Host "⚠ Add-in found but not connected" -ForegroundColor Yellow
        }
    } else {
        Write-Host "✗ Add-in not found in Excel" -ForegroundColor Red
    }
    
    # Don't close Excel automatically
    Write-Host "Excel is now open. Check for the SNIPPER PRO tab in the ribbon." -ForegroundColor Green
}
catch {
    Write-Host "Error testing installation: `$_" -ForegroundColor Red
}
"@

    $testScriptPath = Join-Path $PSScriptRoot "test-snipper-installation.ps1"
    $testScript | Out-File -FilePath $testScriptPath -Encoding UTF8 -Force

    Write-Host ""
    Write-Host "=== Installation Complete ===" -ForegroundColor Green
    Write-Host "✓ Snipper Pro has been successfully installed!" -ForegroundColor Green
    Write-Host ""
    Write-Host "Installation Details:" -ForegroundColor Yellow
    Write-Host "- Location: $programFilesDir" -ForegroundColor White
    Write-Host "- COM Registration: Complete" -ForegroundColor White
    Write-Host "- Excel Registry: Complete" -ForegroundColor White
    Write-Host ""
    Write-Host "Next Steps:" -ForegroundColor Yellow
    Write-Host "1. Open Excel (will start automatically for testing)" -ForegroundColor White
    Write-Host "2. Look for the 'SNIPPER PRO' tab in the ribbon" -ForegroundColor White
    Write-Host "3. Click 'Open Viewer' to load documents" -ForegroundColor White
    Write-Host "4. Use Text Snip, Sum Snip, etc. to extract data" -ForegroundColor White
    Write-Host ""
    
    # Auto-test the installation
    Write-Host "Starting Excel to test installation..." -ForegroundColor Yellow
    & powershell -ExecutionPolicy Bypass -File $testScriptPath

}
catch {
    Write-Host "ERROR during installation: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack Trace: $($_.Exception.StackTrace)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Installation script completed. Check Excel for the SNIPPER PRO tab." -ForegroundColor Green
