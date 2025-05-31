@echo off
echo Installing Snipper Pro COM Add-in (Following Playbook)...
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
echo Cleaning old registrations...
:: Remove old COM registrations
C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe "C:\Users\piete\AppData\Local\SnipperPro\SnipperCloneCleanFinal_v2.dll" /unregister /silent 2>nul

:: Remove old registry entries with different GUIDs
reg delete "HKLM\SOFTWARE\Classes\CLSID\{12345678-1234-1234-1234-123456789012}" /f 2>nul
reg delete "HKCR\CLSID\{12345678-1234-1234-1234-123456789012}" /f 2>nul
reg delete "HKLM\SOFTWARE\Classes\SnipperPro.Connect" /f 2>nul
reg delete "HKCR\SnipperPro.Connect" /f 2>nul
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\SnipperPro.Connect" /f 2>nul

:: Remove new GUID entries to ensure clean start
reg delete "HKLM\SOFTWARE\Classes\CLSID\{D9A6E8B7-F3E1-47B0-B76B-C8DE050D1111}" /f 2>nul
reg delete "HKCR\CLSID\{D9A6E8B7-F3E1-47B0-B76B-C8DE050D1111}" /f 2>nul
reg delete "HKLM\SOFTWARE\Classes\SnipperClone.AddIn" /f 2>nul
reg delete "HKCR\SnipperClone.AddIn" /f 2>nul
reg delete "HKCU\Software\Microsoft\Office\Excel\Addins\SnipperClone.AddIn" /f 2>nul

echo.
echo Copying new DLL with IDTExtensibility2 interface...
copy /Y "SnipperCloneCleanFinal\bin\x86\Release\SnipperCloneCleanFinal.dll" "C:\Users\piete\AppData\Local\SnipperPro\SnipperCloneCleanFinal_playbook.dll"

echo.
echo Registering COM add-in with RegAsm...
C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe "C:\Users\piete\AppData\Local\SnipperPro\SnipperCloneCleanFinal_playbook.dll" /codebase

if %errorLevel% neq 0 (
    echo ERROR: RegAsm registration failed!
    pause
    exit /b 1
)

echo.
echo Verifying COM registration...
reg query "HKCR\SnipperClone.AddIn\CLSID" 2>nul
if %errorLevel% neq 0 (
    echo ERROR: ProgID not found in registry
    pause
    exit /b 1
)

reg query "HKCR\CLSID\{D9A6E8B7-F3E1-47B0-B76B-C8DE050D1111}\InprocServer32" 2>nul
if %errorLevel% neq 0 (
    echo ERROR: CLSID not found in registry
    pause
    exit /b 1
)

echo ✅ COM registration verified!

echo.
echo Setting up Excel add-in registry (no Manifest value for COM)...
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\SnipperClone.AddIn" /v "FriendlyName" /t REG_SZ /d "Snipper Pro" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\SnipperClone.AddIn" /v "Description" /t REG_SZ /d "Snipper Pro - Document Viewer" /f
reg add "HKCU\Software\Microsoft\Office\Excel\Addins\SnipperClone.AddIn" /v "LoadBehavior" /t REG_DWORD /d 3 /f

echo.
echo ✅ Installation completed successfully!
echo.
echo 🔍 Key changes made (following playbook):
echo   - Implemented IDTExtensibility2 interface (MANDATORY for Excel COM add-ins)
echo   - Built with x86 platform for broad compatibility
echo   - Proper COM registration with RegAsm /codebase
echo   - No Manifest registry value (that's for VSTO only)
echo   - LoadBehavior = 3 (load at startup)
echo.
echo You can now:
echo 1. Open Excel
echo 2. Go to File ^> Options ^> Add-ins
echo 3. Select 'COM Add-ins' and click 'Go...'
echo 4. Check 'Snipper Pro' and click OK
echo 5. Look for 'SNIPPER PRO' tab in the ribbon
echo.
pause 