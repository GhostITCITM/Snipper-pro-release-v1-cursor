@echo off
echo Registering SnipperClone COM Add-in...

REM Try 64-bit RegAsm first
if exist "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" (
    echo Using 64-bit RegAsm...
    "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe" "%~dp0SnipperClone\bin\Release\SnipperClone.dll" /codebase /tlb
) else (
    echo Using 32-bit RegAsm...
    "C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe" "%~dp0SnipperClone\bin\Release\SnipperClone.dll" /codebase /tlb
)

echo.
echo Registration complete. Please check for any errors above.
echo.
pause 