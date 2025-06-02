# create-strong-name.ps1
# Script to generate strong name key for assembly signing

Write-Host "Creating strong name key for Snipper Pro..." -ForegroundColor Green

$snPath = "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\sn.exe"
if (!(Test-Path $snPath)) {
    $snPath = "${env:ProgramFiles(x86)}\Microsoft SDKs\Windows\v10.0A\bin\NETFX 4.8 Tools\x64\sn.exe"
}

if (!(Test-Path $snPath)) {
    Write-Host "❌ ERROR: Strong Name Tool (sn.exe) not found." -ForegroundColor Red
    Write-Host "Please install the Windows SDK or Visual Studio with .NET Framework 4.8 SDK." -ForegroundColor Yellow
    exit 1
}

$keyFile = "SnipperPro.snk"
Write-Host "Generating strong name key file: $keyFile" -ForegroundColor Yellow

& $snPath -k $keyFile

if ($LASTEXITCODE -eq 0) {
    Write-Host "`n✅ Strong name key generated successfully!" -ForegroundColor Green
    Write-Host "`nNext steps:" -ForegroundColor Cyan
    Write-Host "1. Run build-snipper-pro.ps1 to rebuild the project" -ForegroundColor White
    Write-Host "2. Run install-snipper-pro.ps1 as Administrator to reinstall" -ForegroundColor White
} else {
    Write-Host "`n❌ Failed to generate strong name key" -ForegroundColor Red
    exit $LASTEXITCODE
} 