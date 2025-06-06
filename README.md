# Snipper Pro - DataSnipper Clone

A powerful Excel add-in for extracting data from PDFs and images, designed to work exactly like DataSnipper.

## 🚀 Features

### **📊 Table Snipping**
- **Professional Column Adjustment** - Add/remove columns with + and - buttons
- **Real PDF Rendering** - View actual PDF content, not just text streams
- **Smart Text Extraction** - PDF text extraction with OCR fallback
- **Excel Integration** - Tab-delimited output for perfect Excel tables

### **📝 Text & Data Extraction**
- **Text Snip** - Extract any text from documents
- **Sum Snip** - Automatically sum numbers in selected areas
- **Validation** - Mark items as verified
- **Exception** - Flag items for review

### **📄 Document Support**
- **PDF Files** - Full PDF rendering with native text extraction
- **Image Files** - PNG, JPG, JPEG, BMP, TIFF, GIF
- **Multi-page Documents** - Navigate through pages with zoom controls
- **Real-time Preview** - See exactly what you're extracting

## 🛠 Installation

### **Quick Install**
1. Run `run_as_admin.bat` to register the add-in
2. Start Excel using `start_excel_with_snipper.bat`
3. Look for the "SNIPPER PRO" tab in Excel's ribbon

### **Manual Install**
1. Build the solution in Release mode
2. Run `register_snipper_pro_simple.ps1` as Administrator
3. Enable the add-in in Excel: File > Options > Add-ins > COM Add-ins

## 🧪 Usage

1. **Open Excel** - Use the provided batch script for best results
2. **Click "Load Document"** - Select PDF or image files
3. **Choose Snip Mode** - Text, Sum, Table, Validation, or Exception
4. **Draw Selection** - Rectangle around the area to extract
5. **Adjust Columns** (Table mode) - Use + and - buttons to adjust column dividers
6. **Double-click** to extract data to Excel

## 🔧 Technical Details

### **Built With**
- **.NET Framework 4.8** - Core framework
- **PDFium** - High-quality PDF rendering
- **Tesseract OCR** - Text recognition fallback
- **Excel COM Interop** - Direct Excel integration

### **Architecture**
- **COM Add-in** - Native Excel integration
- **WinForms UI** - Professional document viewer
- **Modular Design** - Separate engines for OCR, PDF, and Excel

## 📁 Project Structure

```
SnipperCloneCleanFinal/
├── Core/                 # Business logic
├── Infrastructure/       # Logging, config, auth
├── UI/                   # Document viewer interface
├── Assets/               # Resources and ribbon XML
└── bin/Release/          # Built assemblies and dependencies
```

## 🚀 Latest Updates

### **PDF Rendering Fix** (Latest)
- ✅ Fixed PDFium DLL loading issues
- ✅ Added automatic native library copying
- ✅ Enhanced error handling and logging
- ✅ Created launch script for optimal performance

### **Table Snip Enhancement**
- ✅ DataSnipper-style column adjustment UI
- ✅ Column-by-column text extraction
- ✅ Smart tab-delimited Excel output
- ✅ Professional + and - button interface

## 🔍 Troubleshooting

### **Add-in Not Visible**
- Run `check_registration.ps1` to verify registration
- Check Excel: File > Options > Add-ins > COM Add-ins
- Ensure "Snipper Pro" is checked

### **PDF Not Loading**
- Use `start_excel_with_snipper.bat` to launch Excel
- Verify `pdfium.dll` exists in the Release folder
- Check Windows Event Viewer for DLL errors

### **Table Extraction Issues**
- Ensure PDF contains actual text (not just images)
- Adjust column dividers using + and - buttons
- Try OCR fallback for image-based PDFs

## 📋 Requirements

- **Windows 10/11**
- **Microsoft Excel 2016 or later**
- **.NET Framework 4.8**
- **Visual C++ Redistributable** (for PDFium)

## 🤝 Contributing

This is a complete DataSnipper clone implementation. The codebase includes:
- Full table snipping functionality
- Professional document viewer
- Robust error handling
- Comprehensive logging

## 📄 License

Proprietary - Snipper Pro Project

---

**Ready to use!** 🎉 Start with `start_excel_with_snipper.bat` for the best experience. 