# Minimal deployment script for SnipperClone
# Sets up the exact folder structure Excel expects for Developer sideloading

# Paths
$projectRoot     = (Get-Location).Path
$targetPath      = "C:\Program Files\SnipperClone"
$appPath         = Join-Path $targetPath "app"
$manifestSource  = Join-Path $projectRoot "manifest_local.xml"
$manifestTarget  = Join-Path $targetPath "manifest_local.xml"
$distAppFolder   = Join-Path $projectRoot "dist\app"

Write-Host "Setting up minimal SnipperClone deployment..."

# Create folders
Write-Host "Creating folder: $appPath"
New-Item -Path $appPath -ItemType Directory -Force | Out-Null

# Copy manifest
if (Test-Path $manifestSource) {
    Copy-Item -Path $manifestSource -Destination $manifestTarget -Force
    Write-Host "Copied manifest to: $manifestTarget"
} else {
    Write-Host "ERROR: manifest_local.xml not found in project root." -ForegroundColor Red
    exit 1
}

# Copy app files
if (Test-Path $distAppFolder) {
    Write-Host "Copying application files to: $appPath"
    Copy-Item -Path (Join-Path $distAppFolder "*") -Destination $appPath -Recurse -Force
    Write-Host "Copied application files to: $appPath"
} else {
    Write-Host "ERROR: dist/app folder not found. Run the build step first." -ForegroundColor Red
    exit 1
}

Write-Host "Deployment complete."
Write-Host "Next: In Excel, enable Developer tab → Add-ins → Add from File → $manifestTarget"