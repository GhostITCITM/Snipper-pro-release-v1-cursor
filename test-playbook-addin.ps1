# Test the new COM add-in with IDTExtensibility2 interface

Write-Host "Testing Snipper Pro COM Add-in (Playbook Version)..." -ForegroundColor Green

try {
    # Test COM instantiation with new ProgID
    Write-Host "1. Testing COM instantiation with new ProgID..." -ForegroundColor Yellow
    $addin = New-Object -ComObject "SnipperClone.AddIn"
    Write-Host "   ✅ COM instantiation successful!" -ForegroundColor Green
    
    # Test if we can get the type info
    Write-Host "2. Testing type information..." -ForegroundColor Yellow
    $type = $addin.GetType()
    Write-Host "   ✅ Type: $($type.FullName)" -ForegroundColor Green
    
    # Check if IDTExtensibility2 interface is implemented
    Write-Host "3. Checking interface implementation..." -ForegroundColor Yellow
    $interfaces = $type.GetInterfaces() | Where-Object { $_.Name -like "*Extensibility*" -or $_.Name -like "*Ribbon*" }
    foreach ($interface in $interfaces) {
        Write-Host "   ✅ Interface: $($interface.Name)" -ForegroundColor Green
    }
    
    Write-Host "4. COM add-in test completed successfully!" -ForegroundColor Green
    Write-Host ""
    Write-Host "✅ The add-in should now work in Excel. Try:" -ForegroundColor Cyan
    Write-Host "   1. Open Excel" -ForegroundColor White
    Write-Host "   2. Go to File > Options > Add-ins" -ForegroundColor White  
    Write-Host "   3. Select 'COM Add-ins' and click 'Go...'" -ForegroundColor White
    Write-Host "   4. Check 'Snipper Pro' and click OK" -ForegroundColor White
    Write-Host "   5. Look for 'SNIPPER PRO' tab in the ribbon" -ForegroundColor White
    
} catch {
    Write-Host "❌ COM instantiation failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "This indicates a registration problem." -ForegroundColor Yellow
    
    # Check registry entries
    Write-Host ""
    Write-Host "Checking registry entries..." -ForegroundColor Yellow
    
    $clsid = "{D9A6E8B7-F3E1-47B0-B76B-C8DE050D1111}"
    if (Test-Path "HKLM:\SOFTWARE\Classes\CLSID\$clsid") {
        Write-Host "✅ CLSID is registered" -ForegroundColor Green
    } else {
        Write-Host "❌ CLSID not found in registry" -ForegroundColor Red
    }
    
    if (Test-Path "HKLM:\SOFTWARE\Classes\SnipperClone.AddIn") {
        Write-Host "✅ ProgID is registered" -ForegroundColor Green
    } else {
        Write-Host "❌ ProgID not found in registry" -ForegroundColor Red
    }
    
    # Check Excel add-in registry
    if (Test-Path "HKCU:\Software\Microsoft\Office\Excel\Addins\SnipperClone.AddIn") {
        $props = Get-ItemProperty "HKCU:\Software\Microsoft\Office\Excel\Addins\SnipperClone.AddIn"
        Write-Host "✅ Excel add-in registry entry exists" -ForegroundColor Green
        Write-Host "   LoadBehavior: $($props.LoadBehavior)" -ForegroundColor White
    } else {
        Write-Host "❌ Excel add-in registry entry not found" -ForegroundColor Red
    }
} 