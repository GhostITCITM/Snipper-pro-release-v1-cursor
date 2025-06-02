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
    Write-Host "❌ ERROR: MSBuild not found. Please install Visual Studio 2022 or Build Tools." -ForegroundColor Red
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
    Write-Host "❌ ERROR: Project directory not found: $projectDir" -ForegroundColor Red
    exit 1
}

$projectFile = "SnipperCloneCleanFinal.csproj"
if (!(Test-Path $projectFile)) {
    Write-Host "❌ ERROR: Project file not found: $projectFile" -ForegroundColor Red
    exit 1
}

# Build the project
Write-Host "`nBuilding project..." -ForegroundColor Yellow
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
    if (Test-Path $outputDll) {
        Write-Host "`n✅ Build completed successfully!" -ForegroundColor Green
        Write-Host "Output DLL: $outputDll" -ForegroundColor Green
        Write-Host "`nNext steps:" -ForegroundColor Cyan
        Write-Host "1. Run install-snipper-pro.ps1 as Administrator to install the add-in" -ForegroundColor White
    } else {
        Write-Host "`n❌ Build succeeded but output DLL not found at: $outputDll" -ForegroundColor Red
        exit 1
    }
} else {
    Write-Host "`n❌ Build failed with exit code: $LASTEXITCODE" -ForegroundColor Red
    exit $LASTEXITCODE
} 