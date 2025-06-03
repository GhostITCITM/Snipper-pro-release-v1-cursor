#!/usr/bin/env bash
set -euo pipefail

# Script to build and install Snipper Pro on Windows using Git Bash
# Update paths if Visual Studio or Windows are installed in different locations

REPO_ROOT="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$REPO_ROOT/SnipperCloneCleanFinal"
MSBUILD="/c/Program Files (x86)/Microsoft Visual Studio/2019/BuildTools/MSBuild/Current/Bin/MSBuild.exe"
REGASM="/c/Windows/Microsoft.NET/Framework64/v4.0.30319/regasm.exe"
INSTALL_DIR="/c/Program Files/SnipperPro"

# 1. build the add-in
cd "$PROJECT_DIR"
"$MSBUILD" SnipperCloneCleanFinal.csproj \
  /p:Configuration=Release \
  /p:Platform=AnyCPU \
  /verbosity:minimal

# 2. copy DLL
mkdir -p "$INSTALL_DIR"
cp -f "bin/Release/SnipperCloneCleanFinal.dll" "$INSTALL_DIR/"

# 3. register COM component
"$REGASM" "$INSTALL_DIR/SnipperCloneCleanFinal.dll" /codebase /tlb

# 4. create Excel add-in registry keys
powershell.exe -NoLogo -NoProfile -Command '
$reg = "HKCU:\\Software\\Microsoft\\Office\\Excel\\Addins\\SnipperPro"
New-Item -Path $reg -Force | Out-Null
New-ItemProperty -Path $reg -Name Description  -Value "Snipper Pro Excel Add-in" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $reg -Name FriendlyName -Value "Snipper Pro"            -PropertyType String -Force | Out-Null
New-ItemProperty -Path $reg -Name LoadBehavior  -Value 3                       -PropertyType DWord  -Force | Out-Null
New-ItemProperty -Path $reg -Name Manifest      -Value "C:\\Program Files\\SnipperPro\\SnipperCloneCleanFinal.dll" -PropertyType String -Force | Out-Null
'

echo "Snipper Pro built and installed. Open Excel to verify the Snipper Pro tab."
