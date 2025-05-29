@echo off
echo Building FULL SnipperClone with Office Tools...
echo.

REM Try Visual Studio Community first
if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" (
    echo Using Visual Studio Community 2022...
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "SnipperClone\SnipperClone-COM.csproj" /p:Configuration=Release /p:Platform="Any CPU" /v:minimal
    goto :check_result
)

REM Fallback to Build Tools
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" (
    echo Using Visual Studio Build Tools 2022...
    "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" "SnipperClone\SnipperClone-COM.csproj" /p:Configuration=Release /p:Platform="Any CPU" /v:minimal
    goto :check_result
)

echo ERROR: No suitable MSBuild found!
goto :end

:check_result
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
) else (
    echo.
    echo ========================================
    echo BUILD FAILED!
    echo ========================================
    echo Error code: %ERRORLEVEL%
    echo.
    echo Please check the error messages above.
)

:end
pause 