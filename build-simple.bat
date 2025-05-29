@echo off
echo Building SnipperClone...
"C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" "SnipperClone\SnipperClone.csproj" /p:Configuration=Release /p:Platform="Any CPU" /v:minimal
if %ERRORLEVEL% EQU 0 (
    echo Build successful!
    echo Assembly created at: SnipperClone\bin\Release\SnipperClone.dll
) else (
    echo Build failed with error code %ERRORLEVEL%
)
pause 