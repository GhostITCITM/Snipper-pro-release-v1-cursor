# Snipper Pro - Excel Add-in for PDF Data Extraction

A powerful Excel add-in for extracting data from PDFs and images, designed to work like DataSnipper.

## 🚀 Features

- **Table Snipping** - Extract tables from PDFs with column adjustment
- **Text Extraction** - Extract any text from documents  
- **Sum Snip** - Automatically sum numbers in selected areas
- **PDF & Image Support** - Full PDF rendering plus PNG, JPG, JPEG, BMP, TIFF, GIF
- **Excel Integration** - Direct output to Excel with proper formatting
- **Drag-and-Drop Loading** - Drop PDF or image files directly into the viewer
- **Right-Click Removal** - Remove loaded documents from the list via context menu

## 📋 Requirements

- **Windows 10/11**
- **Microsoft Excel 2016 or later**
- **.NET Framework 4.8**
- **Visual Studio 2019/2022** (for building)
- **Administrator privileges** (for registration)

## 🛠 Quick Setup (Recommended)

### 1. Build the Project
```cmd
build.cmd
```

### 2. Register the Add-in
```cmd
run_as_admin.bat
```

### 3. Start Excel
```cmd
start_excel_with_snipper.bat
```

The "SNIPPER PRO" tab should appear in Excel's ribbon.

## 🔧 Manual Setup

### Build from Source
1. Open `SnipperCloneCleanFinal.sln` in Visual Studio
2. Set configuration to **Release**
3. Build Solution (Ctrl+Shift+B)
4. Ensure all dependencies are copied to `bin\Release\`

### Register COM Add-in
Run PowerShell as Administrator:
```powershell
.\register_snipper_pro_simple.ps1
```

To unregister:
```powershell
.\register_snipper_pro_simple.ps1 -Unregister
```

### Verify Installation
```powershell
.\check_registration.ps1
```

## 📁 Project Structure

```
SnipperCloneCleanFinal/
├── Core/                 # Business logic and extraction engines
├── Infrastructure/       # Logging, configuration, authentication
├── UI/                   # Document viewer and user interface
├── Assets/               # Resources and ribbon XML
├── Properties/           # Assembly info and settings
└── tessdata/             # OCR language data files
```

## 🧪 Usage

1. **Open Excel** using `start_excel_with_snipper.bat`
2. **Click "Load Document"** in the SNIPPER PRO ribbon
3. **Select PDF or image file** or simply **drag and drop** files into the viewer
4. **Choose extraction mode**: Text, Sum, Table, Validation, Exception
5. **Draw selection** around the area to extract
6. **Adjust columns** (Table mode) using + and - buttons
7. **Double-click** to extract data to Excel

## 🔍 Troubleshooting

### Add-in Not Visible
- Run `check_registration.ps1` to verify registration
- Check Excel: File → Options → Add-ins → COM Add-ins
- Ensure "Snipper Pro v1" is checked and enabled

### Build Errors
- Ensure .NET Framework 4.8 is installed
- Restore NuGet packages: `nuget.exe restore SnipperCloneCleanFinal.sln`
- Check that all dependencies are available in `packages/` folder

### PDF Loading Issues
- Use `start_excel_with_snipper.bat` to launch Excel
- Verify PDFium DLLs exist in the Release folder
- Check Windows Event Viewer for DLL loading errors

### Registration Issues
- Always run registration scripts as Administrator
- Disable antivirus temporarily during registration
- Check Windows Registry for COM registration entries

## 🔄 Rebuild & Reinstall Process

### Complete Rebuild
1. Clean solution: `Remove-Item -Recurse SnipperCloneCleanFinal\bin, SnipperCloneCleanFinal\obj`
2. Restore packages: `nuget.exe restore SnipperCloneCleanFinal.sln`
3. Build: `build.cmd`
4. Re-register: `run_as_admin.bat`

### Reinstall Add-in
1. Unregister: `.\register_snipper_pro_simple.ps1 -Unregister`
2. Build: `build.cmd`
3. Register: `.\register_snipper_pro_simple.ps1`
4. Restart Excel

## 📦 Dependencies

The following packages are automatically managed via NuGet:
- **PDFium** - PDF rendering engine
- **Tesseract** - OCR text recognition
- **Newtonsoft.Json** - JSON processing
- **Microsoft.Office.Interop.Excel** - Excel COM integration

## 🤝 Support

For issues:
1. Check troubleshooting section above
2. Run `verify_installation.ps1` for diagnostic info
3. Check Windows Event Viewer for detailed error logs

---

**Ready to extract data like a pro!** 🎉 

## 📄 License

See [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md) for open source licenses of the libraries used in this project. This project itself is provided under the terms of the MIT License.

### Image-to-PDF conversion

As of vNEXT, any raster image (PNG/JPG/TIFF/BMP/GIF) dropped into Snipper is automatically converted to a single-page PDF using `SnipperCloneCleanFinal.Core.ImageToPdfConverter`. The converter tries to invoke **Tesseract** with `pdf` output mode (to embed an invisible text layer) and falls back to **PDFsharp** embedding if Tesseract is not available. This removes the legacy image OCR pathway and guarantees that all snips run through the PDF text-extraction engine.

Dependency list:

* **PdfSharp 1.51.518** (MIT) – generates fallback PDFs.
* **Tesseract OCR** (Apache 2.0). You can install the standard Windows build or include `tesseract.exe` alongside the add-in. If `tesseract` is not found on the PATH or in the add-in folder, images will still load (but without selectable text).

