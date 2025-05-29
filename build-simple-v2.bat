@echo off
echo Building SnipperClone with simplified project...
"C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" "SnipperClone\SnipperClone-Simple.csproj" /p:Configuration=Release /p:Platform="Any CPU" /v:minimal
if %ERRORLEVEL% EQU 0 (
    echo Build successful!
    echo Assembly created at: SnipperClone\bin\Release\SnipperClone.dll
    echo.
    echo You can now run the installation:
    echo .\Install-SnipperClone.ps1
) else (
    echo Build failed with error code %ERRORLEVEL%
)
pause 