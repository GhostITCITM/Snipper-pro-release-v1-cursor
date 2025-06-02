# restore-packages.ps1
# Script to download NuGet and restore packages

Write-Host "Setting up NuGet and restoring packages..." -ForegroundColor Green

# Create NuGet directory if it doesn't exist
$nugetDir = "$env:LOCALAPPDATA\NuGet"
if (!(Test-Path $nugetDir)) {
    New-Item -ItemType Directory -Path $nugetDir | Out-Null
}

# Download latest NuGet.exe if not present
$nugetExe = "$nugetDir\nuget.exe"
if (!(Test-Path $nugetExe)) {
    Write-Host "Downloading NuGet.exe..." -ForegroundColor Yellow
    $webClient = New-Object System.Net.WebClient
    $webClient.DownloadFile("https://dist.nuget.org/win-x86-commandline/latest/nuget.exe", $nugetExe)
    Write-Host "   ✅ NuGet.exe downloaded successfully" -ForegroundColor Green
}

# Add NuGet to PATH temporarily
$env:Path = "$nugetDir;$env:Path"

# Create solution file if it doesn't exist
$slnContent = @"
Microsoft Visual Studio Solution File, Format Version 12.00
# Visual Studio Version 17
VisualStudioVersion = 17.0.0.0
MinimumVisualStudioVersion = 10.0.40219.1
Project("{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}") = "SnipperCloneCleanFinal", "SnipperCloneCleanFinal.csproj", "{12345678-1234-1234-1234-123456789ABC}"
EndProject
Global
    GlobalSection(SolutionConfigurationPlatforms) = preSolution
        Release|x64 = Release|x64
    EndGlobalSection
    GlobalSection(ProjectConfigurationPlatforms) = postSolution
        {12345678-1234-1234-1234-123456789ABC}.Release|x64.ActiveCfg = Release|x64
        {12345678-1234-1234-1234-123456789ABC}.Release|x64.Build.0 = Release|x64
    EndGlobalSection
EndGlobal
"@

$slnFile = "SnipperCloneCleanFinal.sln"
if (!(Test-Path $slnFile)) {
    Write-Host "Creating solution file..." -ForegroundColor Yellow
    Set-Content -Path $slnFile -Value $slnContent
    Write-Host "   ✅ Solution file created" -ForegroundColor Green
}

# Create packages directory if it doesn't exist
$packagesDir = "packages"
if (!(Test-Path $packagesDir)) {
    New-Item -ItemType Directory -Path $packagesDir | Out-Null
}

# Install packages
Write-Host "`nInstalling NuGet packages..." -ForegroundColor Yellow

# Install Microsoft.Office.Interop.Excel
Write-Host "Installing Microsoft.Office.Interop.Excel..." -ForegroundColor Yellow
& $nugetExe install Microsoft.Office.Interop.Excel -Version 15.0.4795.1001 -OutputDirectory packages

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n✅ Packages installed successfully!" -ForegroundColor Green
    Write-Host "`nNext steps:" -ForegroundColor Cyan
    Write-Host "1. Run build-snipper-pro.ps1 to build the project" -ForegroundColor White
} else {
    Write-Host "`n❌ Package installation failed with exit code: $LASTEXITCODE" -ForegroundColor Red
    exit $LASTEXITCODE
} 