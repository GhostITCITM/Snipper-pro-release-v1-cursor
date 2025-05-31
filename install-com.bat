@echo off
echo Installing Snipper Pro COM Add-in (Pure COM)...
echo.

:: Check for admin rights
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: This script requires administrator privileges.
    echo Please run as Administrator.
    pause
    exit /b 1
)

echo Stopping Excel processes...
taskkill /f /im EXCEL.EXE 2>nul
timeout /t 2 /nobreak >nul

echo.
echo Removing old registry entries...
:: Remove old CLSID entries
reg delete "HKLM\SOFTWARE\Classes\CLSID\{12345678-1234-1234-1234-123456789012}" /f 2>nul
reg delete "HKCR\CLSID\{12345678-1234-1234-1234-123456789012}" /f 2>nul

:: Remove old ProgID entries  
reg delete "HKLM\SOFTWARE\Classes\SnipperPro.Connect" /f 2>nul
reg delete "HKCR\SnipperPro.Connect" /f 2>nul

:: Remove user registry entry
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\SnipperPro.Connect" /f 2>nul

echo.
echo Registering new COM add-in...
C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe "C:\Users\piete\AppData\Local\SnipperPro\SnipperCloneCleanFinal_v2.dll" /codebase

if %errorLevel% neq 0 (
    echo ERROR: Registration failed!
    pause
    exit /b 1
)

echo.
echo Adding Excel add-in registry entry...
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\SnipperPro.Connect" /v "FriendlyName" /t REG_SZ /d "Snipper Pro" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\SnipperPro.Connect" /v "Description" /t REG_SZ /d "Document analysis add-in" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\SnipperPro.Connect" /v "LoadBehavior" /t REG_DWORD /d 3 /f

echo.
echo ✅ Installation completed successfully!
echo.
echo You can now:
echo 1. Open Excel
echo 2. Go to File ^> Options ^> Add-ins
echo 3. Select 'COM Add-ins' and click 'Go...'
echo 4. Check 'Snipper Pro' and click OK
echo 5. Look for 'SNIPPER PRO' tab in the ribbon
echo.
pause 