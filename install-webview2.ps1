# Download and install WebView2 Runtime
$downloadUrl = "https://go.microsoft.com/fwlink/p/?LinkId=2124703"
$installerPath = Join-Path $env:TEMP "MicrosoftEdgeWebview2Setup.exe"

Write-Host "Downloading WebView2 Runtime..."
Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath

Write-Host "Installing WebView2 Runtime..."
Start-Process -FilePath $installerPath -ArgumentList "/silent /install" -Wait

Write-Host "Installation complete. Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") 