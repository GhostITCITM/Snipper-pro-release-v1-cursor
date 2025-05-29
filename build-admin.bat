@echo off
cd /d "%~dp0"
echo Building SnipperClone in Release mode...
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" SnipperClone\SnipperClone.csproj /p:Configuration=Release /p:Platform=AnyCPU /t:Clean,Build
if %ERRORLEVEL% EQU 0 (
    echo Build completed successfully!
) else (
    echo Build failed with error code %ERRORLEVEL%
    pause
    exit /b %ERRORLEVEL%
)
pause 