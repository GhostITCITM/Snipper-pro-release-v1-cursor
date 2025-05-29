# SnipperClone MSI Build Script
# This script builds a professional MSI installer for enterprise deployment

param(
    [string]$Configuration = "Release",
    [switch]$Clean,
    [switch]$Verbose,
    [string]$OutputPath = ".\dist"
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$WixSourceFile = Join-Path $ScriptDir "installer.wxs"
$LicenseFile = Join-Path $ScriptDir "License.rtf"

Write-Host "SnipperClone MSI Build Script" -ForegroundColor Cyan
Write-Host "=============================" -ForegroundColor Cyan
Write-Host ""

# Validate prerequisites
function Test-Prerequisites {
    Write-Host "Checking MSI build prerequisites..." -ForegroundColor Yellow
    
    # Check if WiX Toolset is installed
    $WixPath = Get-WixToolsetPath
    if (-not $WixPath) {
        Write-Error "WiX Toolset not found. Please install WiX Toolset v3.11 or later from https://wixtoolset.org/"
        return $false
    }
    Write-Host "[OK] WiX Toolset found: $WixPath" -ForegroundColor Green
    
    # Check if main project is built
    $MainAssembly = Join-Path $ScriptDir "SnipperClone\bin\$Configuration\SnipperClone.dll"
    if (-not (Test-Path $MainAssembly)) {
        Write-Error "Main assembly not found: $MainAssembly. Please build the project first using Build-SnipperClone.ps1"
        return $false
    }
    Write-Host "[OK] Main assembly found" -ForegroundColor Green
    
    # Check if WiX source file exists
    if (-not (Test-Path $WixSourceFile)) {
        Write-Error "WiX source file not found: $WixSourceFile"
        return $false
    }
    Write-Host "[OK] WiX source file found" -ForegroundColor Green
    
    return $true
}

# Find WiX Toolset installation
function Get-WixToolsetPath {
    $WixPaths = @(
        "${env:ProgramFiles(x86)}\WiX Toolset v3.11\bin",
        "${env:ProgramFiles}\WiX Toolset v3.11\bin",
        "${env:ProgramFiles(x86)}\Windows Installer XML v3.5\bin",
        "${env:ProgramFiles}\Windows Installer XML v3.5\bin"
    )

    foreach ($path in $WixPaths) {
        $candlePath = Join-Path $path "candle.exe"
        $lightPath = Join-Path $path "light.exe"
        if ((Test-Path $candlePath) -and (Test-Path $lightPath)) {
            return $path
        }
    }

    # Try to find in PATH
    try {
        $candleCmd = Get-Command "candle.exe" -ErrorAction SilentlyContinue
        if ($candleCmd) {
            return Split-Path -Parent $candleCmd.Source
        }
    } catch {
        # Ignore error
    }

    return $null
}

# Create license file if it doesn't exist
function New-LicenseFile {
    if (-not (Test-Path $LicenseFile)) {
        Write-Host "Creating license file..." -ForegroundColor Yellow
        
        $licenseContent = @"
{\rtf1\ansi\deff0 {\fonttbl {\f0 Times New Roman;}}
\f0\fs24
SnipperClone - DataSnipper Alternative\par
\par
MIT License\par
\par
Copyright (c) 2024 Internal Development Team\par
\par
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\par
\par
The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\par
\par
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.\par
\par
Third-Party Components:\par
\par
- Tesseract.js: Apache License 2.0\par
- PDF.js: Apache License 2.0\par
- WebView2: Microsoft Software License\par
- Newtonsoft.Json: MIT License\par
}
"@
        
        Set-Content -Path $LicenseFile -Value $licenseContent -Encoding UTF8
        Write-Host "[OK] License file created" -ForegroundColor Green
    }
}

# Build MSI package
function Build-MSIPackage {
    param([string]$WixPath)
    
    Write-Host "Building MSI package..." -ForegroundColor Yellow
    
    # Ensure output directory exists
    if (-not (Test-Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }
    
    $CandleExe = Join-Path $WixPath "candle.exe"
    $LightExe = Join-Path $WixPath "light.exe"
    $WixObjFile = Join-Path $OutputPath "installer.wixobj"
    $MSIFile = Join-Path $OutputPath "SnipperClone.msi"
    
    # Set verbosity level
    $VerbosityFlag = if ($Verbose) { "-v" } else { "" }
    
    try {
        # Step 1: Compile WiX source to object file
        Write-Host "Compiling WiX source..." -ForegroundColor Gray
        $candleArgs = @(
            $WixSourceFile
            "-out", $WixObjFile
            "-arch", "x64"
            "-dConfiguration=$Configuration"
        )
        
        if ($Verbose) {
            $candleArgs += "-v"
        }
        
        & $CandleExe @candleArgs
        if ($LASTEXITCODE -ne 0) {
            throw "Candle compilation failed with exit code $LASTEXITCODE"
        }
        Write-Host "[OK] WiX source compiled successfully" -ForegroundColor Green
        
        # Step 2: Link object file to create MSI
        Write-Host "Linking MSI package..." -ForegroundColor Gray
        $lightArgs = @(
            $WixObjFile
            "-out", $MSIFile
            "-ext", "WixUIExtension"
            "-cultures:en-US"
            "-loc", "en-US"
        )
        
        if ($Verbose) {
            $lightArgs += "-v"
        } else {
            $lightArgs += "-sw1076"  # Suppress ICE validation warnings in non-verbose mode
        }
        
        & $LightExe @lightArgs
        if ($LASTEXITCODE -ne 0) {
            throw "Light linking failed with exit code $LASTEXITCODE"
        }
        Write-Host "[OK] MSI package created successfully" -ForegroundColor Green
        
        # Validate MSI file
        if (Test-Path $MSIFile) {
            $msiInfo = Get-ItemProperty $MSIFile
            Write-Host "[OK] MSI file created: $MSIFile" -ForegroundColor Green
            Write-Host "     Size: $([math]::Round($msiInfo.Length / 1MB, 2)) MB" -ForegroundColor Gray
            Write-Host "     Created: $($msiInfo.CreationTime)" -ForegroundColor Gray
        } else {
            throw "MSI file was not created"
        }
        
    } catch {
        Write-Error "MSI build failed: $($_.Exception.Message)"
        return $false
    }
    
    return $true
}

# Validate MSI package
function Test-MSIPackage {
    param([string]$MSIPath)
    
    Write-Host "Validating MSI package..." -ForegroundColor Yellow
    
    try {
        # Use Windows Installer API to validate MSI
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database = $installer.OpenDatabase($MSIPath, 0)
        
        # Check basic properties
        $view = $database.OpenView("SELECT * FROM Property WHERE Property = 'ProductName'")
        $view.Execute()
        $record = $view.Fetch()
        if ($record) {
            $productName = $record.StringData(2)
            Write-Host "[OK] Product Name: $productName" -ForegroundColor Green
        }
        
        # Check version
        $view = $database.OpenView("SELECT * FROM Property WHERE Property = 'ProductVersion'")
        $view.Execute()
        $record = $view.Fetch()
        if ($record) {
            $version = $record.StringData(2)
            Write-Host "[OK] Product Version: $version" -ForegroundColor Green
        }
        
        # Check manufacturer
        $view = $database.OpenView("SELECT * FROM Property WHERE Property = 'Manufacturer'")
        $view.Execute()
        $record = $view.Fetch()
        if ($record) {
            $manufacturer = $record.StringData(2)
            Write-Host "[OK] Manufacturer: $manufacturer" -ForegroundColor Green
        }
        
        # Check file count
        $view = $database.OpenView("SELECT COUNT(*) FROM File")
        $view.Execute()
        $record = $view.Fetch()
        if ($record) {
            $fileCount = $record.IntegerData(1)
            Write-Host "[OK] Files included: $fileCount" -ForegroundColor Green
        }
        
        # Cleanup COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($database) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($installer) | Out-Null
        
        Write-Host "[OK] MSI package validation completed" -ForegroundColor Green
        return $true
        
    } catch {
        Write-Warning "MSI validation failed: $($_.Exception.Message)"
        return $false
    }
}

# Generate installation instructions
function New-InstallationInstructions {
    param([string]$MSIPath)
    
    $instructionsFile = Join-Path $OutputPath "INSTALLATION-INSTRUCTIONS.md"
    
    $instructions = @"
# SnipperClone Installation Instructions

## System Requirements

- Windows 10 or Windows 11 (64-bit)
- Microsoft Excel 2016 or later
- .NET Framework 4.8 or later
- Administrator privileges for installation

## Installation Methods

### Method 1: Interactive Installation (Recommended)

1. **Download** the MSI package: ``SnipperClone.msi``
2. **Right-click** on the MSI file and select "Run as administrator"
3. **Follow** the installation wizard:
   - Accept the license agreement
   - Choose installation directory (default recommended)
   - Select features to install
   - Click "Install" to begin installation
4. **Restart Excel** after installation completes
5. **Verify** the installation by looking for the "DATASNIPPER" tab in Excel's ribbon

### Method 2: Silent Installation (Enterprise Deployment)

For automated deployment in corporate environments:

``````powershell
# Install silently with all features
msiexec /i "SnipperClone.msi" /quiet /norestart

# Install with logging
msiexec /i "SnipperClone.msi" /quiet /norestart /l*v "install.log"

# Install to custom directory
msiexec /i "SnipperClone.msi" /quiet /norestart INSTALLFOLDER="C:\CustomPath\SnipperClone"
``````

### Method 3: Group Policy Deployment

1. Copy the MSI file to a network share accessible by target computers
2. Open Group Policy Management Console
3. Navigate to Computer Configuration > Software Settings > Software Installation
4. Right-click and select "New > Package"
5. Browse to the MSI file and select "Assigned" deployment method

## Post-Installation Verification

1. **Open Excel** and look for the "DATASNIPPER" tab in the ribbon
2. **Click** "Open Viewer" to test the document viewer
3. **Import** a test PDF or image file
4. **Try** a Text Snip operation to verify OCR functionality

## Troubleshooting

### Add-in Not Loading

1. Check Excel Add-ins:
   - File > Options > Add-ins > COM Add-ins > Go...
   - Ensure "SnipperClone" is checked
2. Check Windows Event Viewer for errors
3. Verify .NET Framework 4.8 is installed
4. Run Excel as administrator (temporarily)

### OCR Not Working

1. Ensure internet connectivity for initial OCR engine download
2. Check Windows Defender/antivirus settings
3. Verify WebView2 runtime is installed (usually automatic)

### Performance Issues

1. Close other Office applications
2. Ensure sufficient RAM (8GB+ recommended)
3. Check for Windows updates
4. Restart Excel periodically

## Uninstallation

### Interactive Uninstall
1. Go to Settings > Apps & features
2. Find "SnipperClone - DataSnipper Alternative"
3. Click "Uninstall" and follow prompts

### Silent Uninstall
``````powershell
msiexec /x "SnipperClone.msi" /quiet /norestart
``````

## Support

- **Documentation**: Check the installed documentation in the Start Menu
- **Issues**: Report problems through your IT support channel
- **Updates**: New versions will be distributed through the same deployment method

---

**Installation Package**: $(Split-Path -Leaf $MSIPath)
**Generated**: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
**Version**: 1.0.0.0
"@

    Set-Content -Path $instructionsFile -Value $instructions -Encoding UTF8
    Write-Host "[OK] Installation instructions created: $instructionsFile" -ForegroundColor Green
}

# Main build process
try {
    # Validate prerequisites
    if (-not (Test-Prerequisites)) {
        exit 1
    }
    
    # Get WiX path
    $WixPath = Get-WixToolsetPath
    
    Write-Host "Configuration: $Configuration" -ForegroundColor Gray
    Write-Host "Output Path: $OutputPath" -ForegroundColor Gray
    if ($Verbose) {
        Write-Host "Verbose output enabled" -ForegroundColor Gray
    }
    Write-Host ""

    # Clean if requested
    if ($Clean) {
        Write-Host "Cleaning output directory..." -ForegroundColor Yellow
        if (Test-Path $OutputPath) {
            Remove-Item -Path $OutputPath -Recurse -Force
        }
        Write-Host "[OK] Output directory cleaned" -ForegroundColor Green
    }

    # Create license file
    New-LicenseFile

    # Build MSI package
    $buildSuccess = Build-MSIPackage -WixPath $WixPath
    if (-not $buildSuccess) {
        exit 1
    }

    # Validate MSI package
    $MSIFile = Join-Path $OutputPath "SnipperClone.msi"
    Test-MSIPackage -MSIPath $MSIFile

    # Generate installation instructions
    New-InstallationInstructions -MSIPath $MSIFile

    Write-Host ""
    Write-Host "MSI Build Completed Successfully!" -ForegroundColor Green
    Write-Host "=================================" -ForegroundColor Green
    Write-Host "MSI Package: $MSIFile" -ForegroundColor Cyan
    Write-Host "Instructions: $(Join-Path $OutputPath "INSTALLATION-INSTRUCTIONS.md")" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Ready for enterprise deployment!" -ForegroundColor Yellow

} catch {
    Write-Error "MSI build failed: $($_.Exception.Message)"
    exit 1
}

Write-Host ""
Write-Host "Build script completed." -ForegroundColor White 