# Install-User.ps1
# User-level installation script for Snipper Pro Excel Add-in (no admin required)

$ErrorActionPreference = "Stop"

# Configuration for user installation
$config = @{
    InstallPath = "$env:APPDATA\SnipperPro"
    DllName = "SnipperCloneCleanFinal.dll"
    ProgId = "SnipperPro.Connect"
    ClsId = "{D9A6E8B7-F3E1-47B0-B76B-C8DE050D1111}"
    Name = "Snipper Pro"
}

function Write-Step {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Green
}

function Write-Error {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Red
}

# Stop Excel processes
Write-Step "Stopping Excel processes..."
Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

# Clean up old installation
Write-Step "Cleaning up old installation..."

$regPaths = @(
    "HKCU:\Software\Microsoft\Office\Excel\Addins\$($config.ProgId)",
    "HKCU:\Software\Classes\$($config.ProgId)",
    "HKCU:\Software\Classes\CLSID\$($config.ClsId)"
)

foreach ($path in $regPaths) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
    }
}

if (Test-Path $config.InstallPath) {
    Remove-Item -Path $config.InstallPath -Recurse -Force -ErrorAction SilentlyContinue
}

# Create install directory
Write-Step "Setting up installation..."
New-Item -ItemType Directory -Path $config.InstallPath -Force | Out-Null

# Copy DLL
$sourceDll = ".\bin\x64\Release\$($config.DllName)"
$targetDll = Join-Path $config.InstallPath $config.DllName

if (-not (Test-Path $sourceDll)) {
    Write-Error "ERROR: Source DLL not found at: $sourceDll"
    exit 1
}

Copy-Item $sourceDll $targetDll -Force
Write-Success "DLL copied successfully"

# Register COM component for current user
Write-Step "Registering COM component..."
$regasm = "$env:windir\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$result = & $regasm $targetDll /codebase

if ($LASTEXITCODE -ne 0) {
    Write-Error "ERROR: RegAsm failed with exit code: $LASTEXITCODE"
    Write-Error "Output: $result"
    exit 1
}

Write-Success "COM registration successful"

# Configure Excel add-in registry for current user
Write-Step "Configuring Excel add-in registry..."
$regPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\$($config.ProgId)"
New-Item -Path $regPath -Force | Out-Null

$regValues = @{
    "Description" = "$($config.Name) Excel Add-in"
    "FriendlyName" = $config.Name
    "LoadBehavior" = 3
}

foreach ($key in $regValues.Keys) {
    $value = $regValues[$key]
    $type = if ($value -is [int]) { "DWord" } else { "String" }
    Set-ItemProperty -Path $regPath -Name $key -Value $value -Type $type -Force
}

Write-Success "Registry configuration complete"

# Verify installation
Write-Step "Verifying installation..."

$tests = @(
    @{ Path = $targetDll; Message = "DLL exists" },
    @{ Path = "HKCU:\Software\Classes\CLSID\$($config.ClsId)"; Message = "COM registration exists" },
    @{ Path = $regPath; Message = "Excel add-in registration exists" }
)

$failed = $false
foreach ($test in $tests) {
    if (Test-Path $test.Path) {
        Write-Success "checkmark $($test.Message)"
    } else {
        Write-Error "X $($test.Message)"
        $failed = $true
    }
}

if ($failed) {
    Write-Error "Installation verification failed"
    exit 1
}

Write-Success "Installation completed successfully!"

# Display next steps
Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "1. Start Excel" -ForegroundColor White
Write-Host "2. Go to File menu then Options then Add-ins" -ForegroundColor White
Write-Host "3. Look for Snipper Pro in the COM Add-ins list" -ForegroundColor White
Write-Host "4. Check the box to enable it" -ForegroundColor White
Write-Host "5. Look for the SNIPPER PRO tab in the ribbon" -ForegroundColor White 