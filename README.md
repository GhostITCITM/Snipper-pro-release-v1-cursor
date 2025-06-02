# Snipper Pro Excel Add-in

A professional Excel COM add-in that provides document analysis and data extraction capabilities.

## 🚀 Installation

### Prerequisites
- Windows 10/11
- Microsoft Excel 2016 or later (64-bit)
- .NET Framework 4.8
- Visual Studio 2022 or Build Tools for Visual Studio 2022

### Installation Steps

1. **Build the Project**
   ```powershell
   # Run in PowerShell
   .\build-snipper-pro.ps1
   ```
   This script will automatically locate MSBuild and build the project.

2. **Install the Add-in**
   ```powershell
   # Run PowerShell as Administrator
   .\install-snipper-pro.ps1
   ```

3. **Verify in Excel**
   - Open Excel
   - Go to File > Options > Add-ins
   - Select 'COM Add-ins' and click 'Go...'
   - Check 'Snipper Pro' and click OK
   - Look for 'SNIPPER PRO' tab in the ribbon

### Troubleshooting

If the build fails:
1. Install Visual Studio 2022 or Build Tools for Visual Studio 2022
2. Ensure .NET Framework 4.8 SDK is installed
3. Try running Developer PowerShell for VS 2022 manually:
   ```powershell
   # In Developer PowerShell for VS 2022
   msbuild SnipperCloneCleanFinal.csproj /p:Configuration=Release /p:Platform=x64
   ```

If the add-in doesn't appear:
1. Check Windows Event Viewer for .NET Runtime errors
2. Verify the add-in is enabled in Excel's COM Add-ins dialog
3. Try running Excel as Administrator once
4. Check that the DLL exists at `C:\Program Files\SnipperPro\SnipperCloneCleanFinal.dll`

## 📋 Features

The add-in provides a "SNIPPER PRO" tab with the following functionality:

- **Text Snip**: Extract text from selected areas using OCR
- **Sum Snip**: Extract and sum numerical values
- **Table Snip**: Extract structured table data
- **Validation**: Mark cells as validated (✓)
- **Exception**: Mark cells as exceptions (✗)

## 🔧 Technical Details

### COM Registration
```
CLSID: {D9A6E8B7-F3E1-47B0-B76B-C8DE050D1111}
ProgID: SnipperPro.Connect
Class: SnipperCloneCleanFinal.ThisAddIn
```

### Registry Location
```
HKCU:\Software\Microsoft\Office\Excel\Addins\SnipperPro.Connect
```

### Installation Directory
```
C:\Program Files\SnipperPro\
```

## 🛠️ Development

### Build Requirements
- Visual Studio 2022
- Office Development Tools for Visual Studio
- .NET Framework 4.8 SDK
- Microsoft Office (Excel) 2016+ (64-bit)

### Project Structure
- Pure COM add-in (no VSTO)
- Implements `IDTExtensibility2` and `IRibbonExtensibility`
- Custom ribbon UI via XML
- Windows Forms for document viewer

### Dependencies
- Microsoft.Office.Interop.Excel (16.0)
- Microsoft.Office.Core (16.0)
- System.Windows.Forms

## 📝 License

Copyright © 2024. All rights reserved. 