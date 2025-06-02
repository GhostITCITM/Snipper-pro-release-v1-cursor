param(
    [switch]$Force = $false
)

Write-Host "Installing Snipper Pro Required Packages..." -ForegroundColor Green

try {
    # Ensure we're in the right directory
    $projectPath = "SnipperCloneCleanFinal"
    
    if (-not (Test-Path $projectPath)) {
        Write-Error "Project directory not found: $projectPath"
        exit 1
    }

    # Create packages directory if it doesn't exist
    if (-not (Test-Path "packages")) {
        New-Item -ItemType Directory -Path "packages" -Force
        Write-Host "Created packages directory"
    }

    # Install Tesseract OCR package
    Write-Host "Installing Tesseract.NET..." -ForegroundColor Yellow
    $tesseractUrl = "https://www.nuget.org/api/v2/package/Tesseract/5.2.0"
    $tesseractPath = "packages\Tesseract.5.2.0"
    
    if (-not (Test-Path $tesseractPath) -or $Force) {
        Write-Host "Downloading Tesseract package..."
        Invoke-WebRequest -Uri $tesseractUrl -OutFile "tesseract.nupkg"
        Expand-Archive -Path "tesseract.nupkg" -DestinationPath $tesseractPath -Force
        Remove-Item "tesseract.nupkg"
        Write-Host "Tesseract package installed" -ForegroundColor Green
    }

    # Install PdfiumViewer package
    Write-Host "Installing PdfiumViewer..." -ForegroundColor Yellow
    $pdfiumUrl = "https://www.nuget.org/api/v2/package/PdfiumViewer/2.13.0"
    $pdfiumPath = "packages\PdfiumViewer.2.13.0"
    
    if (-not (Test-Path $pdfiumPath) -or $Force) {
        Write-Host "Downloading PdfiumViewer package..."
        Invoke-WebRequest -Uri $pdfiumUrl -OutFile "pdfium.nupkg"
        Expand-Archive -Path "pdfium.nupkg" -DestinationPath $pdfiumPath -Force
        Remove-Item "pdfium.nupkg"
        Write-Host "PdfiumViewer package installed" -ForegroundColor Green
    }

    # Install System.Drawing.Common package
    Write-Host "Installing System.Drawing.Common..." -ForegroundColor Yellow
    $drawingUrl = "https://www.nuget.org/api/v2/package/System.Drawing.Common/7.0.0"
    $drawingPath = "packages\System.Drawing.Common.7.0.0"
    
    if (-not (Test-Path $drawingPath) -or $Force) {
        Write-Host "Downloading System.Drawing.Common package..."
        Invoke-WebRequest -Uri $drawingUrl -OutFile "drawing.nupkg"
        Expand-Archive -Path "drawing.nupkg" -DestinationPath $drawingPath -Force
        Remove-Item "drawing.nupkg"
        Write-Host "System.Drawing.Common package installed" -ForegroundColor Green
    }

    # Download Tesseract language data if needed
    Write-Host "Checking for Tesseract language data..." -ForegroundColor Yellow
    $tessDataPath = "tessdata"
    
    if (-not (Test-Path $tessDataPath)) {
        Write-Host "Creating tessdata directory and downloading English language data..."
        New-Item -ItemType Directory -Path $tessDataPath -Force
        
        $engDataUrl = "https://github.com/tesseract-ocr/tessdata/raw/main/eng.traineddata"
        $engDataPath = "$tessDataPath\eng.traineddata"
        
        Write-Host "Downloading English OCR data..."
        Invoke-WebRequest -Uri $engDataUrl -OutFile $engDataPath
        Write-Host "English OCR data downloaded" -ForegroundColor Green
    }

    Write-Host "`nPackage installation completed successfully!" -ForegroundColor Green
    Write-Host "You can now build the project with OCR and PDF support." -ForegroundColor Cyan
    
    # Instructions for manual Tesseract installation
    Write-Host "`nIMPORTANT:" -ForegroundColor Red
    Write-Host "For best OCR results, install Tesseract OCR manually:" -ForegroundColor Yellow
    Write-Host "1. Download from: https://github.com/UB-Mannheim/tesseract/wiki" -ForegroundColor White
    Write-Host "2. Install to: C:\Program Files\Tesseract-OCR" -ForegroundColor White
    Write-Host "3. Add to PATH environment variable" -ForegroundColor White
    Write-Host "`nIf Tesseract is not available, the app will use fallback pattern recognition." -ForegroundColor Cyan

} catch {
    Write-Error "Failed to install packages: $_"
    exit 1
} 