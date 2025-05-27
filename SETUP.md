# SnipperClone - Complete Installation Guide

## Overview

SnipperClone is a fully offline DataSnipper clone that runs entirely within Excel without requiring internet connectivity or external servers. This guide covers installation via MSI installer and manual deployment.

## Prerequisites

- **Windows 10/11** (required for Office.js add-ins)
- **Excel 2016+** with Microsoft 365 or standalone license
- **Administrator privileges** (for installation only)
- **WiX Toolset v3.11+** (only if building from source)

## Method 1: MSI Installer (Recommended)

### Step 1: Install WiX Toolset (Build Environment Only)

If building from source, install WiX first:

```powershell
# Download from https://wixtoolset.org/releases/
# Install WiX Toolset v3.11 or later
```

### Step 2: Build Application and Installer

```powershell
# Navigate to project directory
cd "C:\Users\piete\Desktop\snipper pro v1"

# Install dependencies and build everything
.\build.ps1
```

This script will:
- Install npm dependencies
- Build the React application
- Download and bundle office.js locally
- Copy PDF.js worker files
- Generate self-signed certificate
- Create SnipperClone.msi installer

### Step 3: Install via MSI

```powershell
# Run the installer (requires Administrator)
.\SnipperClone.msi
```

The installer automatically:
- Copies files to `C:\Program Files\SnipperClone\`
- Registers add-in in Windows Registry
- Installs self-signed certificate
- No additional configuration needed

### Step 4: Verify Installation

1. **Open Excel**
2. **Look for DATASNIPPER tab** in the ribbon
3. **Click Import Docs** to test the task pane opens
4. **Load a PDF file** to verify full functionality

## Method 2: Manual Installation

### Step 1: Build Application

```powershell
# Install dependencies
npm install

# Build application
npm run build

# Run build script to copy all required files
.\build.ps1
```

### Step 2: Deploy Files Manually

```powershell
# Create installation directory
mkdir "C:\Program Files\SnipperClone\app"

# Copy built application files
Copy-Item "dist\*" "C:\Program Files\SnipperClone\app\" -Recurse -Force

# Copy manifest file
Copy-Item "manifest.xml" "C:\Program Files\SnipperClone\" -Force
```

### Step 3: Register Add-in

```powershell
# Add registry entry for Excel add-in discovery
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer\Catalogs" /v "SnipperClone" /t REG_SZ /d "C:\Program Files\SnipperClone\manifest.xml" /f

# Install self-signed certificate
Import-Certificate -FilePath "assets\snipperclone.cer" -CertStoreLocation Cert:\CurrentUser\Root
```

### Step 4: Update Manifest URLs

Edit `C:\Program Files\SnipperClone\manifest.xml` and replace all `~remoteAppUrl` with:
```
file:///C:/Program/Files/SnipperClone/app
```

## File Structure After Installation

```
C:\Program Files\SnipperClone\
├── app\
│   ├── taskpane.html          # Main task pane UI
│   ├── commands.html          # Ribbon command handlers  
│   ├── taskpane.js           # React application bundle
│   ├── commands.js           # Button function handlers
│   ├── pdf.worker.min.js     # PDF.js worker (offline)
│   ├── office.js             # Microsoft Office API (offline)
│   ├── 16.svg, 32.svg, etc.  # Button icons
│   └── [other assets]
└── manifest.xml              # Add-in registration file
```

## Troubleshooting Installation

### Add-in Not Appearing

**Check Registry Entry:**
```powershell
reg query "HKCU\Software\Microsoft\Office\16.0\WEF\Developer\Catalogs"
```

**Verify Files Exist:**
```powershell
dir "C:\Program Files\SnipperClone\"
dir "C:\Program Files\SnipperClone\app\"
```

**Restart Excel:** Close Excel completely and reopen.

### Certificate Issues

**Manual Certificate Install:**
```powershell
# Import certificate with Administrator privileges
Import-Certificate -FilePath "assets\snipperclone.cer" -CertStoreLocation Cert:\CurrentUser\Root
```

**Verify Certificate:**
```powershell
Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Subject -like "*SnipperClone*" }
```

### File Permission Errors

**Set Permissions:**
```powershell
# Grant read access to all users
icacls "C:\Program Files\SnipperClone" /grant "Users:(OI)(CI)R" /T
```

**Run Excel as Administrator:** (temporary troubleshooting step)

### MSI Build Failures

**Check WiX Installation:**
```powershell
# Verify WiX paths exist
Test-Path "C:\Program Files (x86)\WiX Toolset v3.11\bin\candle.exe"
Test-Path "C:\Program Files (x86)\WiX Toolset v3.11\bin\light.exe"
```

**Manual MSI Build:**
```powershell
# If build.ps1 fails, run commands manually
& "C:\Program Files (x86)\WiX Toolset v3.11\bin\candle.exe" installer.wxs
& "C:\Program Files (x86)\WiX Toolset v3.11\bin\light.exe" installer.wixobj -o SnipperClone.msi
```

## Usage Verification

### Test Basic Functionality

1. **Open Excel** and verify DATASNIPPER tab appears
2. **Click Import Docs** - task pane should open  
3. **Load a test PDF** - file should display correctly
4. **Select an Excel cell** - cell address should be detected
5. **Try Text Snip** - click button, then draw rectangle on PDF
6. **Verify data appears** in the selected Excel cell

### Test All Snip Types

- **Text Snip**: Extract text from PDF area
- **Sum Snip**: Calculate sum of numbers in area
- **Table Snip**: Extract table data to multiple cells
- **Validation**: Add checkmark (immediate on cell selection)
- **Exception**: Add cross mark (immediate on cell selection)

### Check Metadata Storage

```powershell
# Verify hidden sheet is created
# Open Excel workbook, press Ctrl+G, type: _Snips
# Should show hidden sheet with snip records
```

## Performance and Limitations

### Expected Performance

- **First Load**: 10-15 seconds (OCR engine initialization)
- **PDF Loading**: 2-5 seconds depending on file size
- **OCR Processing**: 5-30 seconds per snip
- **Text/Validation Snips**: Near-instant after OCR

### File Size Limits

- **Recommended**: PDFs under 50MB
- **Maximum**: Limited by available system memory
- **Page Count**: No hard limit, renders on-demand

### Supported Formats

- **PDF Files**: All versions, including scanned documents
- **Images**: PNG, JPG, GIF via PDF embedding
- **Text**: Both searchable and image-based PDFs

## Security Considerations

### Self-signed Certificate

The installer creates a self-signed certificate for local HTTPS requirements:

- **Purpose**: Required for Office.js add-in loading
- **Scope**: Local machine only, not for internet communication
- **Risk**: Minimal - only enables local file serving
- **Management**: Automatically handled by installer/uninstaller

### Data Privacy

- **Local Processing**: All OCR and PDF processing occurs on local machine
- **No Network**: No data transmitted to external servers
- **User Documents**: Remain entirely on user's machine
- **Metadata**: Stored only in Excel workbook files

### Antivirus Considerations

Some antivirus software may flag:

- **Self-signed certificate installation**
- **Registry modifications**
- **Program Files directory access**

Add exclusions for `C:\Program Files\SnipperClone\` if needed.

## Uninstallation

### Via MSI Installer

```powershell
# Use Windows Add/Remove Programs
# Or run MSI with uninstall flag
msiexec /x SnipperClone.msi
```

### Manual Removal

```powershell
# Remove registry entry
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer\Catalogs" /v "SnipperClone" /f

# Remove certificate
Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Subject -like "*SnipperClone*" } | Remove-Item

# Remove files
Remove-Item "C:\Program Files\SnipperClone\" -Recurse -Force
```

## Support and Maintenance

### Log Files

Check for issues in:
- **Windows Event Viewer** > Application Logs
- **Excel Developer Console** (F12 in task pane)
- **Build logs** from build.ps1 execution

### Updates

To update the application:

1. **Build new version** with updated code
2. **Create new MSI** with incremented version number
3. **Uninstall old version** first
4. **Install new version** with new MSI

### Backup

Important files to backup:
- **Source code** and build scripts
- **Certificate files** (if custom generated)
- **User workbooks** with snip metadata

For additional support, contact the internal development team with specific error messages and system configuration details.