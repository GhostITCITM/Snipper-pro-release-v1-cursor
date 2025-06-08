@echo off
echo Building Snipper Pro with validation/exception fixes...

REM Try different MSBuild locations
set MSBUILD_PATH=""

if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH="C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    goto BUILD
)

if exist "C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH="C:\Program Files\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
    goto BUILD
)

if exist "C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" (
    set MSBUILD_PATH="C:\Program Files\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
    goto BUILD
)

if exist "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe" (
    set MSBUILD_PATH="C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
    goto BUILD
)

echo ERROR: MSBuild not found. Please install Visual Studio or .NET Framework SDK.
pause
exit /b 1

:BUILD
echo Found MSBuild at: %MSBUILD_PATH%
echo Building project...

%MSBUILD_PATH% SnipperCloneCleanFinal.sln /p:Configuration=Release /p:Platform="Any CPU" /v:minimal

if %ERRORLEVEL% == 0 (
    echo.
    echo BUILD SUCCESSFUL!
    echo Updated DLL: SnipperCloneCleanFinal\bin\Release\SnipperCloneCleanFinal.dll
    echo.
    echo To register the updated add-in, run as Administrator:
    echo .\run_as_admin.bat
) else (
    echo.
    echo BUILD FAILED! Check the error messages above.
)

pause 