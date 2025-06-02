param(
    [switch]$Force = $false
)

Write-Host "=== Snipper Pro Complete Uninstall ===" -ForegroundColor Red
Write-Host "Removing all traces of Snipper Pro add-in..." -ForegroundColor Yellow

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
    Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2

    # Unregister COM component
    Write-Host "Unregistering COM component..." -ForegroundColor Yellow
    $dllPath = Join-Path $PSScriptRoot "SnipperCloneCleanFinal\bin\Release\SnipperCloneCleanFinal.dll"
    if (Test-Path $dllPath) {
        try {
            & "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" /unregister "$dllPath" /silent
            Write-Host "✓ COM component unregistered" -ForegroundColor Green
        }
        catch {
            Write-Host "⚠ COM unregister failed or was already unregistered" -ForegroundColor Yellow
        }
    }

    # Remove registry entries
    Write-Host "Removing registry entries..." -ForegroundColor Yellow
    
    $registryPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Office\Excel\Addins\SnipperPro.Connect",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Excel\Addins\SnipperPro.Connect",
        "HKCU:\SOFTWARE\Microsoft\Office\Excel\Addins\SnipperPro.Connect",
        "HKLM:\SOFTWARE\Classes\SnipperPro.Connect",
        "HKLM:\SOFTWARE\WOW6432Node\Classes\SnipperPro.Connect",
        "HKLM:\SOFTWARE\Classes\SnipperCloneCleanFinal.ThisAddIn",
        "HKLM:\SOFTWARE\WOW6432Node\Classes\SnipperCloneCleanFinal.ThisAddIn"
    )

    foreach ($path in $registryPaths) {
        try {
            if (Test-Path $path) {
                Remove-Item -Path $path -Recurse -Force
                Write-Host "✓ Removed: $path" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "⚠ Could not remove: $path" -ForegroundColor Yellow
        }
    }

    # Remove from GAC if present
    Write-Host "Checking Global Assembly Cache..." -ForegroundColor Yellow
    try {
        & "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\gacutil.exe" /u "SnipperCloneCleanFinal" /silent
        Write-Host "✓ Removed from GAC" -ForegroundColor Green
    }
    catch {
        Write-Host "⚠ Not in GAC or removal failed" -ForegroundColor Yellow
    }

    # Clear Excel add-in cache
    Write-Host "Clearing Excel add-in cache..." -ForegroundColor Yellow
    $excelCachePaths = @(
        "$env:LOCALAPPDATA\Microsoft\Office\Excel\",
        "$env:APPDATA\Microsoft\Excel\",
        "$env:TEMP\Excel*"
    )

    foreach ($cachePath in $excelCachePaths) {
        try {
            if (Test-Path $cachePath) {
                Remove-Item -Path $cachePath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            # Ignore errors for cache cleanup
        }
    }

    # Remove from Program Files if copied there
    $programFilesPath = "${env:ProgramFiles}\SnipperPro"
    if (Test-Path $programFilesPath) {
        try {
            Remove-Item -Path $programFilesPath -Recurse -Force
            Write-Host "✓ Removed Program Files installation" -ForegroundColor Green
        }
        catch {
            Write-Host "⚠ Could not remove Program Files installation" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "=== Uninstall Complete ===" -ForegroundColor Green
    Write-Host "All Snipper Pro components have been removed." -ForegroundColor Green
    Write-Host "You can now restart Excel and reinstall if needed." -ForegroundColor Yellow

}
catch {
    Write-Host "ERROR during uninstall: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Press Enter to continue..."
Read-Host 