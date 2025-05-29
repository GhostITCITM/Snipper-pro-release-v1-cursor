# SnipperClone COM Add-in Installation Script
# This script registers the COM add-in with Excel

param(
    [switch]$Uninstall,
    [string]$AssemblyPath = ""
)

$ErrorActionPreference = "Stop"

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as Administrator. Please run PowerShell as Administrator and try again."
    exit 1
}

# Add-in configuration
$AddInName = "SnipperClone"
$AddInDescription = "DataSnipper Clone - Document Analysis Excel Add-in"
$AddInProgId = "SnipperClone.Connect"
$AddInGuid = "{12345678-1234-1234-1234-123456789012}"
$AddInClass = "SnipperClone.Connect"

# Define registry roots
$HKCU_Classes = "HKCU:\\Software\\Classes"
$HKCU_ExcelAddinsBase = "HKCU:\\Software\\Microsoft\\Office\\Excel\\Addins"

# Determine assembly path
if ([string]::IsNullOrEmpty($AssemblyPath)) {
    $ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    $AssemblyPath = Join-Path $ScriptDir "SnipperClone\bin\Release\SnipperClone.dll"
}

if ($Uninstall) {
    Write-Host "Uninstalling SnipperClone COM Add-in from HKCU..." -ForegroundColor Yellow
    
    try {
        # Remove Excel add-in registry entries
        $ExcelAddInsPath = "$HKCU_ExcelAddinsBase\\$AddInProgId"
        if (Test-Path $ExcelAddInsPath) {
            Remove-Item -Path $ExcelAddInsPath -Recurse -Force
            Write-Host "Removed Excel add-in registry entries from HKCU" -ForegroundColor Green
        }
        
        # Remove COM registration
        $ComPath = "$HKCU_Classes\\$AddInProgId"
        if (Test-Path $ComPath) {
            Remove-Item -Path $ComPath -Recurse -Force
            Write-Host "Removed COM registration from HKCU" -ForegroundColor Green
        }
        
        # Remove CLSID registration
        $ClsidPath = "$HKCU_Classes\\CLSID\\$AddInGuid"
        if (Test-Path $ClsidPath) {
            Remove-Item -Path $ClsidPath -Recurse -Force
            Write-Host "Removed CLSID registration from HKCU" -ForegroundColor Green
        }
        
        Write-Host "SnipperClone COM Add-in uninstalled successfully from HKCU!" -ForegroundColor Green
    }
    catch {
        Write-Error "Error during uninstallation: $($_.Exception.Message)"
    }
}
else {
    Write-Host "Installing SnipperClone COM Add-in to HKCU..." -ForegroundColor Cyan
    
    # Verify assembly exists
    if (-not (Test-Path $AssemblyPath)) {
        Write-Error "Assembly not found at: $AssemblyPath"
        Write-Host "Please build the project first or specify the correct path with -AssemblyPath parameter"
        exit 1
    }
    
    $FullAssemblyPath = (Resolve-Path $AssemblyPath).Path
    $AssemblyFileName = Split-Path -Leaf $FullAssemblyPath
    
    # Build the manifest string
    $manifestValue = "file:///$FullAssemblyPath|vstolocal"
    # For .NET COM add-ins, we need to use mscoree.dll as the InprocServer32
    
    Write-Host "Using assembly: $FullAssemblyPath" -ForegroundColor Gray
    Write-Host "Manifest value will be: $manifestValue" -ForegroundColor Gray
    
    try {
        # Set up registry values for HKCU installation
        $comAddInKey = "$HKCU_Classes\\$AddInProgId"
        $clsidKey = "$HKCU_Classes\\CLSID\\$AddInGuid"
        $inprocKey = "$clsidKey\\InprocServer32"
        $progIdKey = "$clsidKey\\ProgId"
        
        Write-Host "Registering COM component in HKCU..." -ForegroundColor Yellow
        
        # Create COM registration
        New-Item -Path $comAddInKey -Force | Out-Null
        New-Item -Path $clsidKey -Force | Out-Null
        New-Item -Path $inprocKey -Force | Out-Null
        New-Item -Path $progIdKey -Force | Out-Null
        
        # Set COM registration values
        Set-ItemProperty -Path $comAddInKey -Name "(Default)" -Value $AddInDescription
        Set-ItemProperty -Path $comAddInKey -Name "CLSID" -Value $AddInGuid
        
        Set-ItemProperty -Path $clsidKey -Name "(Default)" -Value $AddInDescription
        
        # For .NET COM components, InprocServer32 should point to mscoree.dll
        $mscoree = [Environment]::ExpandEnvironmentVariables("%windir%\System32\mscoree.dll")
        Set-ItemProperty -Path $inprocKey -Name "(Default)" -Value $mscoree
        Set-ItemProperty -Path $inprocKey -Name "ThreadingModel" -Value "Both"
        Set-ItemProperty -Path $inprocKey -Name "Assembly" -Value "SnipperClone, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
        Set-ItemProperty -Path $inprocKey -Name "Class" -Value $AddInClass
        Set-ItemProperty -Path $inprocKey -Name "RuntimeVersion" -Value "v4.0.30319"
        Set-ItemProperty -Path $inprocKey -Name "CodeBase" -Value "file:///$($FullAssemblyPath.Replace('\', '/'))"
        
        # Create Programmable subkey (required for COM visibility)
        New-Item -Path "$clsidKey\Programmable" -Force | Out-Null
        
        Set-ItemProperty -Path $progIdKey -Name "(Default)" -Value $AddInProgId
        
        Write-Host "COM component registered successfully in HKCU" -ForegroundColor Green
        
        # Register with Excel in HKCU
        Write-Host "Registering with Excel in HKCU..." -ForegroundColor Yellow
        
        $ExcelAddInsPath = "$HKCU_ExcelAddinsBase\\$AddInProgId"
        New-Item -Path $ExcelAddInsPath -Force | Out-Null
        Set-ItemProperty -Path $ExcelAddInsPath -Name "Description" -Value $AddInDescription
        Set-ItemProperty -Path $ExcelAddInsPath -Name "FriendlyName" -Value $AddInName
        Set-ItemProperty -Path $ExcelAddInsPath -Name "LoadBehavior" -Value 3 -Type DWord
        Set-ItemProperty -Path $ExcelAddInsPath -Name "Manifest" -Value $manifestValue -Type String # Add the Manifest value
        
        Write-Host "Excel registration completed successfully in HKCU" -ForegroundColor Green
        
        # Attempt to register assembly using RegAsm (optional, but can help with type library generation if needed elsewhere)
        # This part is less critical if HKCU registration + Manifest works, and might be removed if problematic.
        Write-Host "Attempting to register assembly with RegAsm (optional)..." -ForegroundColor Yellow
        try {
            $regasm = "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\RegAsm.exe"
            if (-not (Test-Path $regasm)) {
                $regasm = "${env:ProgramFiles}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\RegAsm.exe"
            }
            if (-not (Test-Path $regasm)) {
                # Try .NET Framework tools directory
                $regasm = "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v8.1A\bin\NETFX 4.5.1 Tools\RegAsm.exe"
            }
            
            if (Test-Path $regasm) {
                & $regasm $FullAssemblyPath /codebase /tlb
                Write-Host "Assembly registered with RegAsm" -ForegroundColor Green
            } else {
                Write-Warning "RegAsm.exe not found. Assembly registration skipped."
            }
        }
        catch {
            Write-Warning "RegAsm registration failed: $($_.Exception.Message)"
        }
        
        Write-Host ""
        Write-Host "SnipperClone COM Add-in installed successfully to HKCU!" -ForegroundColor Green
        Write-Host "Please restart Excel to load the add-in." -ForegroundColor Cyan
        Write-Host ""
        Write-Host "If the add-in doesn't appear, check:" -ForegroundColor Yellow
        Write-Host "1. Excel > File > Options > Add-ins > COM Add-ins > Go..." -ForegroundColor Gray
        Write-Host "2. Ensure 'SnipperClone' is checked in the list" -ForegroundColor Gray
        Write-Host "3. Check Windows Event Viewer for any loading errors" -ForegroundColor Gray
    }
    catch {
        Write-Error "Error during installation: $($_.Exception.Message)"
        Write-Host "Rolling back changes..." -ForegroundColor Yellow
        
        # Cleanup on error
        try {
            if (Test-Path "$HKCU_ExcelAddinsBase\\$AddInProgId") {
                Remove-Item -Path "$HKCU_ExcelAddinsBase\\$AddInProgId" -Recurse -Force
            }
            if (Test-Path "$HKCU_Classes\\$AddInProgId") {
                Remove-Item -Path "$HKCU_Classes\\$AddInProgId" -Recurse -Force
            }
            if (Test-Path "$HKCU_Classes\\CLSID\\$AddInGuid") {
                Remove-Item -Path "$HKCU_Classes\\CLSID\\$AddInGuid" -Recurse -Force
            }
        }
        catch {
            Write-Warning "Error during rollback: $($_.Exception.Message)"
        }
    }
}

Write-Host ""
Write-Host "Script completed." -ForegroundColor White 