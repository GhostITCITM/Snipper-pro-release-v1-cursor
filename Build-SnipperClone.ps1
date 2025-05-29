# SnipperClone Build Script
# This script builds the COM add-in project with comprehensive validation

param(
    [string]$Configuration = "Release",
    [switch]$Clean,
    [switch]$Install,
    [switch]$Verbose
)

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$SolutionFile = Join-Path $ScriptDir "SnipperClone.sln"
$ProjectFile = Join-Path $ScriptDir "SnipperClone\SnipperClone.csproj"

Write-Host "SnipperClone Build Script" -ForegroundColor Cyan
Write-Host "=========================" -ForegroundColor Cyan
Write-Host ""

# Validate prerequisites
function Test-Prerequisites {
    Write-Host "Checking prerequisites..." -ForegroundColor Yellow
    
    # Check .NET Framework
    $dotNetVersion = Get-ItemProperty "HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\" -Name Release -ErrorAction SilentlyContinue
    if (-not $dotNetVersion -or $dotNetVersion.Release -lt 528040) {
        Write-Error ".NET Framework 4.8 or later is required"
        return $false
    }
    Write-Host "[OK] .NET Framework 4.8+ detected" -ForegroundColor Green
    
    # Check if solution file exists
    if (-not (Test-Path $SolutionFile)) {
        Write-Error "Solution file not found: $SolutionFile"
        return $false
    }
    Write-Host "[OK] Solution file found" -ForegroundColor Green
    
    # Check if project file exists
    if (-not (Test-Path $ProjectFile)) {
        Write-Error "Project file not found: $ProjectFile"
        return $false
    }
    Write-Host "[OK] Project file found" -ForegroundColor Green
    
    return $true
}

# Find MSBuild
function Find-MSBuild {
    Write-Host "Locating MSBuild..." -ForegroundColor Yellow
    
    $MSBuildPaths = @(
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Enterprise\MSBuild\15.0\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Professional\MSBuild\15.0\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2017\Community\MSBuild\15.0\Bin\MSBuild.exe",
        "${env:ProgramFiles(x86)}\MSBuild\14.0\Bin\MSBuild.exe"
    )

    foreach ($path in $MSBuildPaths) {
        if (Test-Path $path) {
            Write-Host "[OK] MSBuild found: $path" -ForegroundColor Green
            return $path
        }
    }

    Write-Error "MSBuild not found. Please install Visual Studio or Build Tools for Visual Studio."
    return $null
}

# Validate build output
function Test-BuildOutput {
    param([string]$OutputDir)
    
    Write-Host "Validating build output..." -ForegroundColor Yellow
    
    $AssemblyPath = Join-Path $OutputDir "SnipperClone.dll"
    
    if (-not (Test-Path $AssemblyPath)) {
        Write-Error "Output assembly not found: $AssemblyPath"
        return $false
    }
    Write-Host "[OK] Assembly found: $AssemblyPath" -ForegroundColor Green
    
    try {
        # Load and inspect assembly
        $Assembly = [System.Reflection.Assembly]::LoadFrom($AssemblyPath)
        $Version = $Assembly.GetName().Version
        Write-Host "[OK] Assembly version: $Version" -ForegroundColor Green
        
        # Verify key components are included
        $Types = $Assembly.GetTypes()
        $RequiredTypes = @(
            "SnipperClone.Connect",
            "SnipperClone.Core.SnipEngine", 
            "SnipperClone.Core.TableParser",
            "SnipperClone.Core.MetadataManager",
            "SnipperClone.Core.OCREngine",
            "SnipperClone.Core.ExcelHelper",
            "SnipperClone.DocumentViewer"
        )
        
        $MissingTypes = @()
        foreach ($RequiredType in $RequiredTypes) {
            if (-not ($Types | Where-Object { $_.FullName -eq $RequiredType })) {
                $MissingTypes += $RequiredType
            }
        }
        
        if ($MissingTypes.Count -eq 0) {
            Write-Host "[OK] All required components found in assembly" -ForegroundColor Green
        } else {
            Write-Warning "Missing components: $($MissingTypes -join ', ')"
            return $false
        }
        
        # Check for COM visibility
        $ConnectType = $Types | Where-Object { $_.FullName -eq "SnipperClone.Connect" }
        if ($ConnectType) {
            $ComVisibleAttr = $ConnectType.GetCustomAttributes([System.Runtime.InteropServices.ComVisibleAttribute], $false)
            if ($ComVisibleAttr.Length -gt 0 -and $ComVisibleAttr[0].Value) {
                Write-Host "[OK] COM visibility confirmed" -ForegroundColor Green
            } else {
                Write-Warning "COM visibility not detected"
            }
        }
        
    } catch {
        Write-Warning "Could not fully validate assembly: $($_.Exception.Message)"
        return $false
    }
    
    # Check for WebAssets
    $WebAssetsDir = Join-Path $OutputDir "WebAssets"
    if (Test-Path $WebAssetsDir) {
        $ViewerHtml = Join-Path $WebAssetsDir "viewer.html"
        if (Test-Path $ViewerHtml) {
            Write-Host "[OK] WebAssets found: viewer.html" -ForegroundColor Green
        } else {
            Write-Warning "WebAssets directory found but viewer.html is missing"
            return $false
        }
    } else {
        Write-Warning "WebAssets directory not found in output"
        return $false
    }
    
    # Check dependencies
    $ExpectedDependencies = @(
        "Microsoft.Web.WebView2.Core.dll",
        "Microsoft.Web.WebView2.WinForms.dll",
        "Newtonsoft.Json.dll"
    )
    
    foreach ($Dependency in $ExpectedDependencies) {
        $DependencyPath = Join-Path $OutputDir $Dependency
        if (Test-Path $DependencyPath) {
            Write-Host "[OK] Dependency found: $Dependency" -ForegroundColor Green
        } else {
            Write-Warning "Missing dependency: $Dependency"
        }
    }
    
    return $true
}

# Main build process
try {
    # Validate prerequisites
    if (-not (Test-Prerequisites)) {
        exit 1
    }
    
    # Find MSBuild
    $MSBuild = Find-MSBuild
    if (-not $MSBuild) {
        exit 1
    }
    
    Write-Host "Configuration: $Configuration" -ForegroundColor Gray
    if ($Verbose) {
        Write-Host "Verbose output enabled" -ForegroundColor Gray
    }
    Write-Host ""

    # Set verbosity level
    $VerbosityLevel = if ($Verbose) { "normal" } else { "minimal" }

    # Clean if requested
    if ($Clean) {
        Write-Host "Cleaning solution..." -ForegroundColor Yellow
        & $MSBuild $SolutionFile /t:Clean /p:Configuration=$Configuration /v:$VerbosityLevel
        if ($LASTEXITCODE -ne 0) {
            throw "Clean failed with exit code $LASTEXITCODE"
        }
        Write-Host "[OK] Clean completed successfully" -ForegroundColor Green
        Write-Host ""
    }

    # Restore NuGet packages first
    Write-Host "Restoring NuGet packages..." -ForegroundColor Yellow
    & $MSBuild $SolutionFile /t:Restore /p:Configuration=$Configuration /v:$VerbosityLevel
    if ($LASTEXITCODE -ne 0) {
        Write-Error "NuGet package restoration failed"
        exit 1
    }
    Write-Host "[OK] NuGet packages restored successfully" -ForegroundColor Green
    Write-Host ""

    # Build the solution
    Write-Host "Building solution..." -ForegroundColor Yellow
    $BuildStartTime = Get-Date
    
    & $MSBuild $SolutionFile /p:Configuration=$Configuration /v:$VerbosityLevel /p:Platform="Any CPU" /p:OutputPath="bin\$Configuration\" /flp:logfile=build.log`;verbosity=diagnostic
    
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Build failed. Check build.log for details."
        if (Test-Path "build.log") {
            Write-Host "Last 20 lines of build log:" -ForegroundColor Red
            Get-Content "build.log" | Select-Object -Last 20 | ForEach-Object { Write-Host $_ -ForegroundColor Gray }
        }
        exit 1
    }
    
    $BuildEndTime = Get-Date
    $BuildDuration = $BuildEndTime - $BuildStartTime
    Write-Host "[OK] Build completed successfully in $($BuildDuration.TotalSeconds.ToString('F1')) seconds" -ForegroundColor Green
    Write-Host ""

    # Validate build output
    $OutputDir = Join-Path $ScriptDir "SnipperClone\bin\$Configuration"
    if (-not (Test-BuildOutput $OutputDir)) {
        Write-Error "Build output validation failed"
        exit 1
    }
    Write-Host ""

    # Copy additional files if needed
    Write-Host "Copying additional files..." -ForegroundColor Yellow
    
    # Ensure WebAssets are copied
    $WebAssetsSource = Join-Path $ScriptDir "SnipperClone\WebAssets"
    $WebAssetsTarget = Join-Path $OutputDir "WebAssets"
    
    if (Test-Path $WebAssetsSource) {
        if (Test-Path $WebAssetsTarget) {
            Remove-Item $WebAssetsTarget -Recurse -Force
        }
        Copy-Item $WebAssetsSource $WebAssetsTarget -Recurse -Force
        Write-Host "[OK] WebAssets copied to output directory" -ForegroundColor Green
    }
    
    # Copy any additional configuration files
    $ConfigFiles = @("app.config", "SnipperClone.exe.config")
    foreach ($ConfigFile in $ConfigFiles) {
        $SourcePath = Join-Path $ScriptDir "SnipperClone\$ConfigFile"
        $TargetPath = Join-Path $OutputDir $ConfigFile
        
        if (Test-Path $SourcePath) {
            Copy-Item $SourcePath $TargetPath -Force
            Write-Host "[OK] Copied $ConfigFile" -ForegroundColor Green
        }
    }
    
    Write-Host ""

    # Generate build summary
    Write-Host "Build Summary" -ForegroundColor Cyan
    Write-Host "=============" -ForegroundColor Cyan
    Write-Host "Configuration: $Configuration" -ForegroundColor White
    Write-Host "Build Time: $($BuildDuration.TotalSeconds.ToString('F1')) seconds" -ForegroundColor White
    Write-Host "Output Directory: $OutputDir" -ForegroundColor White
    
    # List key output files
    $KeyFiles = @(
        "SnipperClone.dll",
        "SnipperClone.tlb",
        "Microsoft.Web.WebView2.Core.dll",
        "Microsoft.Web.WebView2.WinForms.dll",
        "Newtonsoft.Json.dll"
    )
    
    Write-Host "Key Output Files:" -ForegroundColor White
    foreach ($File in $KeyFiles) {
        $FilePath = Join-Path $OutputDir $File
        if (Test-Path $FilePath) {
            $FileInfo = Get-Item $FilePath
            $FileSize = [math]::Round($FileInfo.Length / 1KB, 1)
            Write-Host "  [OK] $File ($FileSize KB)" -ForegroundColor Green
        } else {
            Write-Host "  [MISSING] $File" -ForegroundColor Red
        }
    }
    
    Write-Host ""
    Write-Host "Build completed successfully!" -ForegroundColor Green
    Write-Host "To install the add-in, run: .\Install-SnipperClone.ps1" -ForegroundColor Cyan
    
    # Optional installation
    if ($Install) {
        Write-Host ""
        Write-Host "Installing add-in..." -ForegroundColor Yellow
        & "$ScriptDir\Install-SnipperClone.ps1" -AssemblyPath (Join-Path $OutputDir "SnipperClone.dll")
    }
}
catch {
    Write-Host ""
    Write-Host "Build Failed!" -ForegroundColor Red
    Write-Host "=============" -ForegroundColor Red
    Write-Error "Build failed: $($_.Exception.Message)"
    
    if ($Verbose) {
        Write-Host ""
        Write-Host "Full Error Details:" -ForegroundColor Red
        Write-Host $_.Exception.ToString() -ForegroundColor Red
    }
    
    exit 1
}

Write-Host ""
Write-Host "Build script completed successfully." -ForegroundColor Green 