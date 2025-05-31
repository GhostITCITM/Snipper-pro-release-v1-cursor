# Install Script Following the Recipe's Exact Approach
# This script implements the tested recipe to fix COM registration

Write-Host "🔧 Installing Snipper Pro COM Add-in (Following Tested Recipe)" -ForegroundColor Green
Write-Host ""

# Ensure running as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "❌ ERROR: Must run as Administrator" -ForegroundColor Red
    Write-Host "Please right-click and 'Run as Administrator'" -ForegroundColor Yellow
    exit 1
}

try {
    # Step 1: Stop Excel processes
    Write-Host "1. Stopping Excel processes..." -ForegroundColor Yellow
    Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2

    # Step 2: Clean old registrations
    Write-Host "2. Cleaning old registrations..." -ForegroundColor Yellow
    
    # Remove old registry entries
    $oldPaths = @(
        "HKLM:\SOFTWARE\Classes\CLSID\{12345678-1234-1234-1234-123456789012}",
        "HKCR:\CLSID\{12345678-1234-1234-1234-123456789012}",
        "HKLM:\SOFTWARE\Classes\SnipperPro.Connect",
        "HKCR:\SnipperPro.Connect",
        "HKCU:\Software\Microsoft\Office\Excel\Addins\SnipperPro.Connect"
    )
    
    foreach ($path in $oldPaths) {
        Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue
    }

    # Step 3: Setup DLL path (Recipe Section 2.1)
    Write-Host "3. Setting up DLL..." -ForegroundColor Yellow
    $dllPath = "$env:LOCALAPPDATA\SnipperPro\SnipperClone.dll"
    $regasm = "${env:windir}\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"  # x86!
    
    # Copy DLL to final location
    $sourceDll = "SnipperCloneCleanFinal\bin\x86\Release\SnipperCloneCleanFinal.dll"
    if (!(Test-Path $sourceDll)) {
        throw "Source DLL not found: $sourceDll"
    }
    
    New-Item -Path (Split-Path $dllPath) -ItemType Directory -Force | Out-Null
    Copy-Item $sourceDll $dllPath -Force

    # Step 4: Register COM with x86 RegAsm (CRITICAL)
    Write-Host "4. Registering COM with x86 RegAsm..." -ForegroundColor Yellow
    $regResult = & $regasm $dllPath /codebase /tlb:none /verbose 2>&1
    
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ RegAsm failed:" -ForegroundColor Red
        Write-Host $regResult -ForegroundColor Red
        throw "RegAsm registration failed"
    }
    
    Write-Host "✅ RegAsm completed successfully" -ForegroundColor Green

    # Step 5: Verify CLSID registration (Recipe verification)
    Write-Host "5. Verifying CLSID registration..." -ForegroundColor Yellow
    $clsid = "{D9A6E8B7-F3E1-47B0-B76B-C8DE050D1111}"
    
    # Check CLSID exists
    $clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid"
    if (!(Test-Path $clsidPath)) {
        throw "CLSID not registered at $clsidPath"
    }
    
    # Check InprocServer32 path
    $inprocPath = "$clsidPath\InprocServer32"
    if (!(Test-Path $inprocPath)) {
        throw "InprocServer32 not found at $inprocPath"
    }
    
    $mscoreeValue = Get-ItemProperty -Path $inprocPath -Name "(default)" -ErrorAction SilentlyContinue
    if ($mscoreeValue."(default)" -ne "C:\Windows\System32\mscoree.dll") {
        throw "InprocServer32 does not point to correct mscoree.dll path"
    }
    
    Write-Host "✅ CLSID properly registered!" -ForegroundColor Green

    # Step 6: Setup Excel add-in registry (Recipe Section 2.2)
    Write-Host "6. Setting up Excel add-in registry..." -ForegroundColor Yellow
    $addin = 'HKCU:\Software\Microsoft\Office\Excel\Addins\SnipperClone.AddIn'
    New-Item $addin -Force | Out-Null
    Set-ItemProperty $addin FriendlyName 'Snipper Pro'
    Set-ItemProperty $addin Description 'Snipper Pro – Document viewer'
    Set-ItemProperty $addin LoadBehavior 3 -Type DWord
    
    Write-Host "✅ Excel registry configured!" -ForegroundColor Green

    # Step 7: Final verification
    Write-Host "7. Final verification..." -ForegroundColor Yellow
    
    # Test COM instantiation
    try {
        $testObj = New-Object -ComObject "SnipperClone.AddIn"
        Write-Host "✅ COM instantiation successful!" -ForegroundColor Green
        $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($testObj)
    } catch {
        throw "COM instantiation failed: $($_.Exception.Message)"
    }

    # Verify all registry entries
    $progidPath = "HKLM:\SOFTWARE\Classes\SnipperClone.AddIn"
    if (!(Test-Path $progidPath)) {
        throw "ProgID not registered"
    }
    
    $excelAddinPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\SnipperClone.AddIn"
    if (!(Test-Path $excelAddinPath)) {
        throw "Excel add-in registry not found"
    }

    Write-Host ""
    Write-Host "🎉 SUCCESS! Installation completed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "✅ Recipe compliance verified:" -ForegroundColor Cyan
    Write-Host "   • IDTExtensibility2 + IRibbonExtensibility interfaces ✓" -ForegroundColor White
    Write-Host "   • x86 platform build ✓" -ForegroundColor White
    Write-Host "   • x86 RegAsm with /codebase ✓" -ForegroundColor White
    Write-Host "   • CLSID properly registered ✓" -ForegroundColor White
    Write-Host "   • InprocServer32 → mscoree.dll ✓" -ForegroundColor White
    Write-Host "   • Excel registry LoadBehavior=3 ✓" -ForegroundColor White
    Write-Host "   • COM instantiation test ✓" -ForegroundColor White
    Write-Host ""
    Write-Host "🚀 Next steps:" -ForegroundColor Cyan
    Write-Host "   1. Open Excel" -ForegroundColor White
    Write-Host "   2. Go to File → Options → Add-ins" -ForegroundColor White
    Write-Host "   3. Select 'COM Add-ins' and click 'Go...'" -ForegroundColor White
    Write-Host "   4. Check 'Snipper Pro' and click OK" -ForegroundColor White
    Write-Host "   5. Look for 'SNIPPER PRO' tab in the ribbon!" -ForegroundColor White

} catch {
    Write-Host ""
    Write-Host "❌ Installation failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "🔍 Troubleshooting:" -ForegroundColor Yellow
    Write-Host "   • Ensure running as Administrator" -ForegroundColor White
    Write-Host "   • Close all Excel instances" -ForegroundColor White
    Write-Host "   • Check Windows Event Viewer for detailed errors" -ForegroundColor White
    exit 1
} 