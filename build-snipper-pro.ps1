# build-snipper-pro.ps1
# Build script for Snipper Pro Excel Add-in

Write-Host "Building Snipper Pro Excel Add-in..." -ForegroundColor Green

# Find MSBuild from Visual Studio 2022 installation
$vsWhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$msBuildPath = ""

if (Test-Path $vsWhere) {
    # Try to find VS2022 installation
    $vsPath = & $vsWhere -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe | Select-Object -First 1
    if ($vsPath) {
        $msBuildPath = $vsPath
    }
}

# Fallback paths if vswhere doesn't find it
$fallbackPaths = @(
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe",
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
)

if (-not $msBuildPath) {
    foreach ($path in $fallbackPaths) {
        if (Test-Path $path) {
            $msBuildPath = $path
            break
        }
    }
}

if (-not $msBuildPath) {
    Write-Host "ERROR: MSBuild not found. Please install Visual Studio 2022 or Build Tools." -ForegroundColor Red
    Write-Host "Download VS2022 Build Tools: https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022" -ForegroundColor Yellow
    exit 1
}

Write-Host "Found MSBuild at: $msBuildPath" -ForegroundColor Green

# Change to the project directory
$projectDir = "SnipperCloneCleanFinal"
if (Test-Path $projectDir) {
    Set-Location $projectDir
    Write-Host "Changed to project directory: $projectDir" -ForegroundColor Yellow
} else {
    Write-Host "ERROR: Project directory not found: $projectDir" -ForegroundColor Red
    exit 1
}

# Ensure tessdata is available
$tessDir = "tessdata"
if (-not (Test-Path $tessDir)) {
    New-Item -ItemType Directory -Path $tessDir | Out-Null
}
$engData = Join-Path $tessDir "eng.traineddata"
if (-not (Test-Path $engData)) {
    Write-Host "Downloading default tessdata files..." -ForegroundColor Yellow
    try {
        $url = "https://github.com/tesseract-ocr/tessdata/raw/main/eng.traineddata"
        Invoke-WebRequest -Uri $url -OutFile $engData
        Write-Host "Downloaded eng.traineddata" -ForegroundColor Green
    } catch {
        Write-Host "Failed to download tessdata: $_" -ForegroundColor Red
    }
}

$projectFile = "SnipperCloneCleanFinal.csproj"
if (!(Test-Path $projectFile)) {
    Write-Host "ERROR: Project file not found: $projectFile" -ForegroundColor Red
    exit 1
}

$nugetExe = "$PSScriptRoot\nuget.exe"
if (!(Test-Path $nugetExe)) {
    Write-Host "Downloading nuget.exe..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri "https://dist.nuget.org/win-x86-commandline/latest/nuget.exe" -OutFile $nugetExe
}

# After changing directory, restore
Write-Host "Restoring NuGet packages..." -ForegroundColor Yellow
& $nugetExe restore "packages.config" -PackagesDirectory "..\packages" | Out-Null

# Build the project
Write-Host "Building project..." -ForegroundColor Yellow
$buildArgs = @(
    $projectFile,
    "/p:Configuration=Release",
    "/p:Platform=AnyCPU",
    "/t:Rebuild",
    "/v:m"
)

& $msBuildPath $buildArgs

if ($LASTEXITCODE -eq 0) {
    # Verify the output DLL exists
    $outputDll = "bin\Release\SnipperCloneCleanFinal.dll"
    $outputDllDir = Split-Path $outputDll -Parent
    if (Test-Path $outputDll) {
        Write-Host "Build completed successfully!" -ForegroundColor Green
        Write-Host "Output DLL: $outputDll" -ForegroundColor Green
        Write-Host "Next steps:" -ForegroundColor Cyan
        Write-Host "1. Run install-snipper-pro.ps1 as Administrator to install the add-in" -ForegroundColor White

        # Copy native Pdfium DLLs
        $packageRoot = "..\packages"
        $x86Dll = Get-ChildItem -Path $packageRoot -Recurse -Filter pdfium.dll | Where-Object { $_.FullName -match "x86.no_v8" -and $_.FullName -notmatch "x86_64" } | Select-Object -First 1
        $x64Dll = Get-ChildItem -Path $packageRoot -Recurse -Filter pdfium.dll | Where-Object { $_.FullName -match "x86_64.no_v8" } | Select-Object -First 1
        if ($x64Dll) {
            Copy-Item $x64Dll.FullName "$outputDllDir\pdfium.dll" -Force
            Write-Host "Copied x64 pdfium.dll to output directory" -ForegroundColor Green
        } elseif ($x86Dll) {
            Copy-Item $x86Dll.FullName "$outputDllDir\pdfium.dll" -Force
            Write-Host "Copied x86 pdfium.dll to output directory" -ForegroundColor Green
        } else {
            Write-Host "WARNING: Pdfium native DLLs not found; PDF rendering will fail" -ForegroundColor Red
        }

        # Copy native Tesseract DLLs
        $tesseractNativePath = "..\packages\Tesseract.5.2.0\x64"
        if (Test-Path $tesseractNativePath) {
            $leptonicaDll = Join-Path $tesseractNativePath "leptonica-1.82.0.dll"
            $tesseractDll = Join-Path $tesseractNativePath "tesseract50.dll"
            
            if (Test-Path $leptonicaDll) {
                Copy-Item $leptonicaDll "$outputDllDir\" -Force
                Write-Host "Copied leptonica-1.82.0.dll to output directory" -ForegroundColor Green
            }
            
            if (Test-Path $tesseractDll) {
                Copy-Item $tesseractDll "$outputDllDir\" -Force
                Write-Host "Copied tesseract50.dll to output directory" -ForegroundColor Green
            }
        } else {
            Write-Host "WARNING: Tesseract native DLLs not found; OCR will fail" -ForegroundColor Red
        }

        # Copy tessdata directory to output
        $tessdataSource = "tessdata"
        $tessdataTarget = "$outputDllDir\tessdata"
        if (Test-Path $tessdataSource) {
            if (Test-Path $tessdataTarget) {
                Remove-Item $tessdataTarget -Recurse -Force
            }
            Copy-Item $tessdataSource $tessdataTarget -Recurse -Force
            Write-Host "Copied tessdata directory to output" -ForegroundColor Green
        }
    } else {
        Write-Host "Build succeeded but output DLL not found at: $outputDll" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "Build failed with exit code: $LASTEXITCODE" -ForegroundColor Red
    exit $LASTEXITCODE
}
