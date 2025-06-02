@echo off
echo Building Snipper Pro...
"C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" SnipperCloneCleanFinal.csproj /p:Configuration=Release /p:Platform=AnyCPU /p:RestorePackages=false /p:EnableNuGetPackageRestore=false /p:RestoreProjectStyle=None /t:Build /v:normal
if %errorlevel% equ 0 (
    echo Build successful!
    dir bin\Release\SnipperCloneCleanFinal.dll
) else (
    echo Build failed with error %errorlevel%
)
pause 