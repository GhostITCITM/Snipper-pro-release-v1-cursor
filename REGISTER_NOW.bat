@echo off
echo ====================================
echo FORCE REGISTERING SNIPPER PRO
echo ====================================

echo Stopping Excel...
taskkill /f /im excel.exe >nul 2>&1

echo.
echo Registering COM component...
"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" "%~dp0SnipperCloneCleanFinal\bin\Release\SnipperCloneCleanFinal.dll" /codebase /verbose

echo.
echo Creating Excel registry entry...
reg add "HKLM\SOFTWARE\Microsoft\Office\Excel\Addins\SnipperPro.Connect" /v LoadBehavior /t REG_DWORD /d 3 /f
reg add "HKLM\SOFTWARE\Microsoft\Office\Excel\Addins\SnipperPro.Connect" /v FriendlyName /t REG_SZ /d "Snipper Pro Excel Add-in" /f
reg add "HKLM\SOFTWARE\Microsoft\Office\Excel\Addins\SnipperPro.Connect" /v Description /t REG_SZ /d "DataSnipper Clone for Excel" /f

echo.
echo ====================================
echo REGISTRATION COMPLETE!
echo ====================================
echo Starting Excel...
start excel
echo.
echo Look for SNIPPER PRO tab in Excel ribbon!
pause 