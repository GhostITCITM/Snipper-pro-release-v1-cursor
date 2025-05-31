# Test-Addin.ps1 - Test if the COM add-in can be instantiated

Write-Host "Testing Snipper Pro COM Add-in..." -ForegroundColor Green

try {
    # Test COM instantiation
    Write-Host "1. Testing COM instantiation..." -ForegroundColor Yellow
    $addin = New-Object -ComObject "SnipperPro.Connect"
    Write-Host "   ✅ COM instantiation successful!" -ForegroundColor Green
    
    # Test if we can get the type info
    Write-Host "2. Testing type information..." -ForegroundColor Yellow
    $type = $addin.GetType()
    Write-Host "   ✅ Type: $($type.FullName)" -ForegroundColor Green
    
    Write-Host "3. COM add-in test completed successfully!" -ForegroundColor Green
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
    
    # Check if the CLSID exists
    $clsid = "{12345678-1234-1234-1234-123456789012}"
    if (Test-Path "HKLM:\SOFTWARE\Classes\CLSID\$clsid") {
        Write-Host "✅ CLSID is registered" -ForegroundColor Green
    } else {
        Write-Host "❌ CLSID not found in registry" -ForegroundColor Red
    }
    
    # Check if ProgID exists
    if (Test-Path "HKLM:\SOFTWARE\Classes\SnipperPro.Connect") {
        Write-Host "✅ ProgID is registered" -ForegroundColor Green
    } else {
        Write-Host "❌ ProgID not found in registry" -ForegroundColor Red
    }
} 