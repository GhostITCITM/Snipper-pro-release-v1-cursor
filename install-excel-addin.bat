@echo off
cd /d "%~dp0"

echo Building SnipperClone...
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" SnipperClone\SnipperClone.csproj /p:Configuration=Release /p:Platform=AnyCPU /t:Clean,Build
if %ERRORLEVEL% NEQ 0 (
    echo Build failed!
    pause
    exit /b %ERRORLEVEL%
)

echo Installing WebView2...
start /wait MicrosoftEdgeWebview2Setup.exe /silent /install

echo Registering COM components...
regsvr32 /s "SnipperClone\bin\Release\SnipperClone.dll"

echo Installing Office tools...
start /wait vstor_redist.exe /quiet

echo Setup complete! The add-in should now be available in Excel.
echo If you don't see it, please restart Excel.
pause 