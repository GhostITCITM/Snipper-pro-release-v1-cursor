@echo off
echo Building SnipperClone (Working Version)...
echo.

REM Set the correct MSBuild path
set MSBUILD_PATH="C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"

REM Check if MSBuild exists
if not exist %MSBUILD_PATH% (
    echo MSBuild not found at Community location, trying BuildTools...
    set MSBUILD_PATH="C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
)

REM Check again
if not exist %MSBUILD_PATH% (
    echo ERROR: MSBuild not found!
    pause
    exit /b 1
)

echo Using MSBuild at: %MSBUILD_PATH%
echo.

REM Clean first
echo Cleaning previous build...
%MSBUILD_PATH% "SnipperClone\SnipperClone-Simple.csproj" /t:Clean /p:Configuration=Release /p:Platform=AnyCPU /v:minimal

REM Build the project
echo Building SnipperClone...
%MSBUILD_PATH% "SnipperClone\SnipperClone-Simple.csproj" /p:Configuration=Release /p:Platform=AnyCPU /v:minimal

if %ERRORLEVEL% EQU 0 (
    echo.
    echo ========================================
    echo BUILD SUCCESSFUL!
    echo ========================================
    echo Assembly created at: SnipperClone\bin\Release\SnipperClone.dll
    echo.
    echo You can now install the add-in:
    echo .\Install-SnipperClone.ps1
    echo.
    echo Or build the MSI installer:
    echo .\Build-MSI.ps1
    echo.
) else (
    echo.
    echo ========================================
    echo BUILD FAILED!
    echo ========================================
    echo Error code: %ERRORLEVEL%
    echo.
    echo Check the error messages above.
)

echo.
pause 