Write-Host "Testing Handwriting Recognition in Snipper Pro" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green

# Create a simple test image with handwritten-style text using .NET
Add-Type -AssemblyName System.Drawing

$width = 400
$height = 200
$bitmap = New-Object System.Drawing.Bitmap($width, $height)
$graphics = [System.Drawing.Graphics]::FromImage($bitmap)

# White background
$graphics.Clear([System.Drawing.Color]::White)

# Create handwriting-style text
$font = New-Object System.Drawing.Font("Segoe Script", 24)
$brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::Black)

# Draw some handwritten-style text and numbers
$graphics.DrawString("Hello World", $font, $brush, 20, 20)
$graphics.DrawString("Total: $123.45", $font, $brush, 20, 70)
$graphics.DrawString("Items: 567", $font, $brush, 20, 120)

# Save test image
$testImagePath = Join-Path $PSScriptRoot "test_handwriting.png"
$bitmap.Save($testImagePath, [System.Drawing.Imaging.ImageFormat]::Png)

$graphics.Dispose()
$bitmap.Dispose()

Write-Host "`nTest image created at: $testImagePath" -ForegroundColor Yellow
Write-Host "`nTo test handwriting recognition:" -ForegroundColor Cyan
Write-Host "1. Open Excel with Snipper Pro loaded"
Write-Host "2. Load the test image: $testImagePath"
Write-Host "3. Try snipping the handwritten text"
Write-Host "4. Check if the text is recognized correctly"

Write-Host "`nThe handwriting recognition will automatically activate when:" -ForegroundColor Magenta
Write-Host "- Regular OCR confidence is low (<30%)"
Write-Host "- No text is detected by regular OCR"
Write-Host "- The image contains curved strokes typical of handwriting"

Write-Host "`nPress any key to open the test image..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

Start-Process $testImagePath 