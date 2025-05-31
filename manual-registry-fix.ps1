# Manual Registry Creation for Snipper Pro COM Add-in
# This bypasses RegAsm and creates the registry entries manually

Write-Host "🔧 Manually creating registry entries for COM registration..." -ForegroundColor Green

# Ensure running as Administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "❌ ERROR: Must run as Administrator" -ForegroundColor Red
    exit 1
}

try {
    $clsid = "{D9A6E8B7-F3E1-47B0-B76B-C8DE050D1111}"
    $progid = "SnipperClone.AddIn"
    $className = "SnipperCloneCleanFinal.ThisAddIn"
    $assemblyInfo = "SnipperCloneCleanFinal, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null"
    $dllPath = "C:\Users\piete\AppData\Local\SnipperPro\SnipperClone.dll"

    # 1. Create ProgID entries
    Write-Host "1. Creating ProgID entries..." -ForegroundColor Yellow
    
    $progidPath = "HKLM:\SOFTWARE\Classes\$progid"
    New-Item -Path $progidPath -Force | Out-Null
    Set-ItemProperty -Path $progidPath -Name "(default)" -Value $className
    
    $progidClsidPath = "$progidPath\CLSID"
    New-Item -Path $progidClsidPath -Force | Out-Null
    Set-ItemProperty -Path $progidClsidPath -Name "(default)" -Value $clsid

    # 2. Create CLSID main entry
    Write-Host "2. Creating CLSID main entry..." -ForegroundColor Yellow
    
    $clsidPath = "HKLM:\SOFTWARE\Classes\CLSID\$clsid"
    New-Item -Path $clsidPath -Force | Out-Null
    Set-ItemProperty -Path $clsidPath -Name "(default)" -Value $className

    # 3. Create InprocServer32 entry (CRITICAL)
    Write-Host "3. Creating InprocServer32 entry..." -ForegroundColor Yellow
    
    $inprocPath = "$clsidPath\InprocServer32"
    New-Item -Path $inprocPath -Force | Out-Null
    Set-ItemProperty -Path $inprocPath -Name "(default)" -Value "C:\Windows\System32\mscoree.dll"
    Set-ItemProperty -Path $inprocPath -Name "ThreadingModel" -Value "Both"
    Set-ItemProperty -Path $inprocPath -Name "Class" -Value $className
    Set-ItemProperty -Path $inprocPath -Name "Assembly" -Value $assemblyInfo
    Set-ItemProperty -Path $inprocPath -Name "RuntimeVersion" -Value "v4.0.30319"
    Set-ItemProperty -Path $inprocPath -Name "CodeBase" -Value "file:///$($dllPath -replace '\\', '/')"

    # 4. Create version-specific entry
    Write-Host "4. Creating version-specific entry..." -ForegroundColor Yellow
    
    $versionPath = "$inprocPath\1.0.0.0"
    New-Item -Path $versionPath -Force | Out-Null
    Set-ItemProperty -Path $versionPath -Name "Class" -Value $className
    Set-ItemProperty -Path $versionPath -Name "Assembly" -Value $assemblyInfo
    Set-ItemProperty -Path $versionPath -Name "RuntimeVersion" -Value "v4.0.30319"
    Set-ItemProperty -Path $versionPath -Name "CodeBase" -Value "file:///$($dllPath -replace '\\', '/')"

    # 5. Create ProgId reference
    Write-Host "5. Creating ProgId reference..." -ForegroundColor Yellow
    
    $progIdRefPath = "$clsidPath\ProgId"
    New-Item -Path $progIdRefPath -Force | Out-Null
    Set-ItemProperty -Path $progIdRefPath -Name "(default)" -Value $progid

    # 6. Create Implemented Categories
    Write-Host "6. Creating Implemented Categories..." -ForegroundColor Yellow
    
    $categoryPath = "$clsidPath\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
    New-Item -Path $categoryPath -Force | Out-Null

    # 7. Setup Excel add-in registry
    Write-Host "7. Setting up Excel add-in registry..." -ForegroundColor Yellow
    
    $excelAddinPath = "HKCU:\Software\Microsoft\Office\Excel\Addins\$progid"
    New-Item -Path $excelAddinPath -Force | Out-Null
    Set-ItemProperty -Path $excelAddinPath -Name "FriendlyName" -Value "Snipper Pro"
    Set-ItemProperty -Path $excelAddinPath -Name "Description" -Value "Snipper Pro - Document Viewer"
    Set-ItemProperty -Path $excelAddinPath -Name "LoadBehavior" -Value 3 -Type DWord

    Write-Host ""
    Write-Host "✅ Registry entries created successfully!" -ForegroundColor Green

    # 8. Verification
    Write-Host "8. Verifying registration..." -ForegroundColor Yellow
    
    if (Test-Path $clsidPath) {
        Write-Host "   ✅ CLSID registered" -ForegroundColor Green
    } else {
        throw "CLSID not found"
    }
    
    if (Test-Path $inprocPath) {
        Write-Host "   ✅ InprocServer32 created" -ForegroundColor Green
    } else {
        throw "InprocServer32 not found"
    }
    
    $mscoreeValue = Get-ItemProperty -Path $inprocPath -Name "(default)" -ErrorAction SilentlyContinue
    if ($mscoreeValue."(default)" -eq "C:\Windows\System32\mscoree.dll") {
        Write-Host "   ✅ mscoree.dll path correct" -ForegroundColor Green
    } else {
        throw "mscoree.dll path incorrect"
    }

    # 9. Test COM instantiation
    Write-Host "9. Testing COM instantiation..." -ForegroundColor Yellow
    
    try {
        $testObj = New-Object -ComObject $progid
        Write-Host "   ✅ COM instantiation successful!" -ForegroundColor Green
        $null = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($testObj)
    } catch {
        Write-Host "   ❌ COM instantiation failed: $($_.Exception.Message)" -ForegroundColor Red
        throw "COM test failed"
    }

    Write-Host ""
    Write-Host "🎉 SUCCESS! Manual COM registration completed!" -ForegroundColor Green
    Write-Host ""
    Write-Host "✅ All recipe requirements satisfied:" -ForegroundColor Cyan
    Write-Host "   • CLSID registered with full mscoree.dll path ✓" -ForegroundColor White
    Write-Host "   • InprocServer32 pointing to C:\Windows\System32\mscoree.dll ✓" -ForegroundColor White
    Write-Host "   • ProgID registered ✓" -ForegroundColor White
    Write-Host "   • Excel add-in registry (LoadBehavior=3) ✓" -ForegroundColor White
    Write-Host "   • COM instantiation test passed ✓" -ForegroundColor White

} catch {
    Write-Host ""
    Write-Host "❌ Manual registration failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
} 