@echo off
echo Setting up Office Tools for SnipperClone build...

REM Create the OfficeTools directory
mkdir "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Microsoft\VisualStudio\v17.0\OfficeTools" 2>nul

REM Copy the targets file
copy "Microsoft.VisualStudio.Tools.Office.targets" "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Microsoft\VisualStudio\v17.0\OfficeTools\"

if %ERRORLEVEL% EQU 0 (
    echo Office tools setup complete!
    echo.
    echo Now building the full SnipperClone project...
    echo.
    call build-full.bat
) else (
    echo Failed to copy Office tools targets file.
    echo Trying alternative approach...
    echo.
    
    REM Try with Community edition
    mkdir "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Microsoft\VisualStudio\v17.0\OfficeTools" 2>nul
    copy "Microsoft.VisualStudio.Tools.Office.targets" "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Microsoft\VisualStudio\v17.0\OfficeTools\"
    
    if %ERRORLEVEL% EQU 0 (
        echo Office tools setup complete for Community edition!
        echo.
        call build-full.bat
    ) else (
        echo Failed to setup Office tools. Please run as administrator.
    )
)

pause 