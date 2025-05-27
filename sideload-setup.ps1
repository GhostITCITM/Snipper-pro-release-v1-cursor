# Sideload setup script for SnipperClone
# Registers a sideloaded add-in and creates a starter workbook without emojis

$addinName    = "SnipperClone"
$manifestPath = "C:\Users\piete\Desktop\snipper pro v1\sideload\manifest_local.xml"
$docPath      = "C:\Users\piete\Desktop\snipper pro v1\sideload\SnipperCloneStarter.xlsx"
$registryKey  = "HKCU:\Software\Microsoft\Office\16.0\WEF\SideloadedAddins\$addinName"

Write-Host "Setting up direct sideload for $addinName..."

# 1. Create sideload registry entry
New-Item -Path $registryKey -Force | Out-Null
Set-ItemProperty -Path $registryKey -Name "ManifestPath" -Value $manifestPath
Set-ItemProperty -Path $registryKey -Name "DocumentPath" -Value $docPath
Set-ItemProperty -Path $registryKey -Name "IsEnabled"    -Value 1

Write-Host "Registry entry created at $registryKey"

# 2. Create starter workbook if it doesn't exist
if (-not (Test-Path $docPath)) {
    Write-Host "Creating starter workbook at $docPath"
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Add()
        $workbook.SaveAs($docPath)
        $workbook.Close()
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)     | Out-Null
        Write-Host "Starter workbook created at $docPath"
    } catch {
        Write-Host "Could not create workbook automatically. Please create one at $docPath manually." -ForegroundColor Yellow
        Read-Host "Press Enter once you have created the starter workbook"
    }
} else {
    Write-Host "Starter workbook already exists at $docPath"
}

# 3. Copy built files to C:\Program Files\SnipperClone\app
$appPath = "C:\Program Files\SnipperClone\app"
Write-Host "Copying built files to $appPath..."
New-Item -ItemType Directory -Path $appPath -Force | Out-Null

$distPath = "dist\app\*"
if (Test-Path $distPath) {
    Copy-Item $distPath $appPath -Recurse -Force
    Write-Host "Copied app files to $appPath"
} else {
    Write-Host "dist\app\ not found. Please run the build step first." -ForegroundColor Yellow
}

Write-Host "Setup complete. Run launch-snipper.ps1 to start Excel with SnipperClone."