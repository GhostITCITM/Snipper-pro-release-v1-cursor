Write-Host "=== Snipper Pro Installation Verification ===" -ForegroundColor Green
Write-Host ""

# Check if DLL exists
$dllPath = "${env:ProgramFiles}\SnipperPro\SnipperCloneCleanFinal.dll"
if (Test-Path $dllPath) {
    Write-Host "✓ DLL Found: $dllPath" -ForegroundColor Green
} else {
    Write-Host "✗ DLL Not Found: $dllPath" -ForegroundColor Red
}

# Check registry entries
try {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Office\Excel\Addins\SnipperPro.Connect"
    $regEntry = Get-ItemProperty -Path $regPath -ErrorAction Stop
    Write-Host "✓ Registry Entry Found" -ForegroundColor Green
    Write-Host "  Description: $($regEntry.Description)" -ForegroundColor White
    Write-Host "  FriendlyName: $($regEntry.FriendlyName)" -ForegroundColor White
    Write-Host "  LoadBehavior: $($regEntry.LoadBehavior)" -ForegroundColor White
    Write-Host "  Manifest: $($regEntry.Manifest)" -ForegroundColor White
} 
catch {
    Write-Host "✗ Registry Entry Not Found" -ForegroundColor Red
}

# Check COM registration
Write-Host ""
Write-Host "Checking COM Registration..." -ForegroundColor Yellow
try {
    $comClass = [System.Runtime.InteropServices.Marshal]::GetTypeFromProgID("SnipperPro.Connect")
    if ($comClass) {
        Write-Host "✓ COM Class Registered: SnipperPro.Connect" -ForegroundColor Green
    } else {
        Write-Host "✗ COM Class Not Found" -ForegroundColor Red
    }
} 
catch {
    Write-Host "✗ COM Registration Error: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ""
Write-Host "=== Installation Status ===" -ForegroundColor Yellow

$dllExists = Test-Path $dllPath
$regExists = $null -ne (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\Excel\Addins\SnipperPro.Connect" -ErrorAction SilentlyContinue)

if ($dllExists -and $regExists) {
    Write-Host "✓ INSTALLATION COMPLETE" -ForegroundColor Green
    Write-Host ""
    Write-Host "Next Steps:" -ForegroundColor Yellow
    Write-Host "1. Open Excel" -ForegroundColor White
    Write-Host "2. Look for 'SNIPPER PRO' tab in ribbon" -ForegroundColor White
    Write-Host "3. Click 'Open Viewer' to load documents" -ForegroundColor White
    Write-Host "4. Use snip buttons to extract data" -ForegroundColor White
    Write-Host ""
    Write-Host "Features Available:" -ForegroundColor Yellow
    Write-Host "• Load multiple images and PDFs" -ForegroundColor White
    Write-Host "• Real OCR text extraction" -ForegroundColor White
    Write-Host "• Number detection and summation" -ForegroundColor White
    Write-Host "• Table structure detection" -ForegroundColor White
    Write-Host "• DataSnipper-style DS formulas" -ForegroundColor White
    Write-Host "• Visual snip highlighting" -ForegroundColor White
    Write-Host "• Multiple document support" -ForegroundColor White
} else {
    Write-Host "✗ INSTALLATION INCOMPLETE" -ForegroundColor Red
    Write-Host "Run install-snipper-pro-complete.ps1 as Administrator" -ForegroundColor Yellow
}

Write-Host "" 