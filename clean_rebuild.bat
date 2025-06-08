@echo off
echo ========================================
echo Snipper Pro - Complete Clean Rebuild
echo ========================================
echo.

echo Step 1: Cleaning old build files...
if exist "SnipperCloneCleanFinal\bin" rmdir /s /q "SnipperCloneCleanFinal\bin"
if exist "SnipperCloneCleanFinal\obj" rmdir /s /q "SnipperCloneCleanFinal\obj"
echo ✓ Cleaned build directories

echo.
echo Step 2: Restoring NuGet packages...
nuget.exe restore SnipperCloneCleanFinal.sln
if %ERRORLEVEL% neq 0 (
    echo ❌ Failed to restore NuGet packages
    pause
    exit /b 1
)
echo ✓ NuGet packages restored

echo.
echo Step 3: Building solution...
call build.cmd
if %ERRORLEVEL% neq 0 (
    echo ❌ Build failed
    pause
    exit /b 1
)

echo.
echo Step 4: Registering add-in...
echo Please run as Administrator: run_as_admin.bat
echo.
echo ========================================
echo Clean rebuild completed successfully!
echo ========================================
echo.
echo Next steps:
echo 1. Run 'run_as_admin.bat' as Administrator
echo 2. Launch Excel with 'start_excel_with_snipper.bat'
echo.
pause 