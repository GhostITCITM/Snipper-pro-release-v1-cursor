# SnipperClone - Enhanced DataSnipper Alternative COM Add-in

A comprehensive Excel COM add-in that replicates and enhances DataSnipper's core functionality for document analysis and data extraction. This native COM implementation provides superior performance, easier deployment, and full corporate compatibility with advanced OCR, table parsing, and document processing capabilities.

## 🚀 Enhanced Features

### Core Snipping Functionality
- **🔤 Text Snip**: Advanced OCR with Tesseract.js integration, image preprocessing, and intelligent text cleaning
- **🧮 Sum Snip**: Enhanced number detection supporting multiple formats (currency, percentages, parentheses notation)
- **📊 Table Snip**: Multi-strategy table parsing with quality scoring and professional Excel formatting
- **✅ Validation Snip**: Visual validation marks with comprehensive metadata tracking
- **❌ Exception Snip**: Exception markers with detailed logging and audit trails

### Advanced Document Support
- **📄 PDF Documents**: Full PDF viewing with PDF.js integration and text extraction
- **🖼️ Image Files**: Support for PNG, JPG, JPEG, BMP, TIFF, GIF formats
- **📑 Multi-page Navigation**: Smooth page-by-page document browsing with keyboard shortcuts
- **🔍 Zoom Controls**: Fit width, fit page, custom zoom levels with mouse wheel support
- **📏 File Size Validation**: Support for documents up to 100MB with progress indicators

### Enhanced OCR Engine
- **🎯 Image Preprocessing**: Contrast adjustment, brightness optimization, grayscale conversion
- **🔧 Noise Reduction**: Gaussian blur and sharpening filters for better text recognition
- **🌍 Multi-language Support**: Tesseract.js with configurable language packs
- **⚡ Performance Optimization**: WebView2 worker management and timeout handling
- **🎛️ OCR Parameters**: Configurable character whitelisting and page segmentation modes

### Intelligent Table Parser
- **📋 Multiple Strategies**: Markdown, tab-delimited, pipe-delimited, CSV, space-delimited parsing
- **🏆 Quality Scoring**: Advanced algorithm to select the best parsing strategy
- **🔍 Pattern Recognition**: Intelligent detection of headers, data types, and table structure
- **🧹 Data Cleaning**: Automatic OCR artifact correction and whitespace normalization
- **📐 Column Alignment**: Smart column boundary detection and content alignment

### Professional Excel Integration
- **🎨 Custom Ribbon Tab**: Dedicated "DATASNIPPER" tab with intuitive button layout
- **⚡ Real-time Selection**: Dynamic cell selection tracking with visual feedback
- **🎯 Automatic Formatting**: Professional table formatting with headers, borders, and styling
- **💾 Hidden Metadata**: Secure storage of snip information in hidden worksheets
- **🔗 Jump-back Navigation**: Click any snip cell to navigate to its source location

### Enhanced User Interface
- **🎨 Modern Design**: Fluent UI-inspired interface with smooth animations
- **⌨️ Keyboard Shortcuts**: Full keyboard navigation support (arrows, +/-, W, F, Escape)
- **🖱️ Mouse Interactions**: Wheel zoom, drag selection, context menus
- **📱 Responsive Layout**: Adaptive UI that works across different screen sizes
- **🌈 Visual Feedback**: Progress indicators, status updates, and error messages

## 🛠 Installation

### Prerequisites
- Microsoft Excel 2016 or later (Office 365 recommended)
- Windows 10/11 (64-bit recommended)
- .NET Framework 4.8 or later
- Visual Studio 2019/2022 (for building from source)
- Administrator privileges (for COM registration)

### Quick Install
1. **Download the latest release** from the releases page
2. **Extract** the files to a local directory
3. **Run PowerShell as Administrator**
4. **Execute the build script**:
   ```powershell
   .\Build-SnipperClone.ps1 -Configuration Release
   ```
5. **Install the add-in**:
   ```powershell
   .\Install-SnipperClone.ps1
   ```
6. **Restart Excel** to load the add-in

### Build from Source
```powershell
# Clone the repository
git clone https://github.com/your-repo/SnipperClone.git
cd SnipperClone

# Build and install in one step
.\Build-SnipperClone.ps1 -Configuration Release -Install

# Or build and install separately
.\Build-SnipperClone.ps1 -Configuration Release
.\Install-SnipperClone.ps1
```

### Verification
1. Open Excel
2. Look for the **"DATASNIPPER"** tab in the ribbon
3. Click **"Open Viewer"** to test document loading
4. Import a PDF or image file to verify functionality

## 📖 Usage Guide

### Getting Started
1. **Open Excel** and navigate to the **DATASNIPPER** tab
2. **Click "Open Viewer"** to launch the document viewer
3. **Import a document** using the "Import Document" button
4. **Select a cell** in Excel where you want to extract data
5. **Choose a snip mode** (Text, Sum, Table, Validation, or Exception)
6. **Draw a rectangle** around the area you want to extract
7. **Data is automatically extracted** and inserted into the selected cell

### Snip Modes Explained

#### 🔤 Text Snip
- **Purpose**: Extract text from any area of a document
- **How to use**: Select cell → Click "Text Snip" → Draw rectangle around text
- **Features**: Advanced OCR with preprocessing, text cleaning, and error correction
- **Best for**: Names, addresses, descriptions, notes

#### 🧮 Sum Snip
- **Purpose**: Automatically detect and sum all numbers in an area
- **How to use**: Select cell → Click "Sum Snip" → Draw rectangle around numbers
- **Features**: Supports currency symbols, percentages, negative numbers in parentheses
- **Best for**: Financial statements, invoices, calculation verification

#### 📊 Table Snip
- **Purpose**: Extract entire tables with structure preservation
- **How to use**: Select cell → Click "Table Snip" → Draw rectangle around table
- **Features**: Intelligent table detection, header recognition, professional formatting
- **Best for**: Financial data, schedules, lists, structured information

#### ✅ Validation Snip
- **Purpose**: Mark areas as validated without extracting data
- **How to use**: Select cell → Click "Validation" → Draw rectangle around verified area
- **Features**: Inserts checkmark symbol, maintains audit trail
- **Best for**: Audit procedures, verification workflows

#### ❌ Exception Snip
- **Purpose**: Flag areas as exceptions or issues
- **How to use**: Select cell → Click "Exception" → Draw rectangle around problem area
- **Features**: Inserts cross symbol, logs exception details
- **Best for**: Audit findings, discrepancies, items requiring attention

### Advanced Features

#### Jump-back Navigation
- **Click any snip cell** to automatically navigate to its source location in the document
- **Visual highlighting** shows the exact area that was snipped
- **Page navigation** automatically switches to the correct page

#### Snip Management
- **Highlight All Snips**: Shows all snip locations with color coding
- **Clear Highlights**: Removes visual indicators
- **Delete Snip**: Removes snip data and metadata
- **Show Snip Info**: Displays detailed information about any snip

#### Keyboard Shortcuts
- **Arrow Keys**: Navigate document pages
- **+/-**: Zoom in/out
- **W**: Fit width
- **F**: Fit page
- **Escape**: Clear current selection

## 🏗 Architecture

### Component Overview
```
SnipperClone COM Add-in
├── Connect.cs              # Main COM add-in entry point
├── DocumentViewer.cs       # Windows Forms document viewer
├── SnipperRibbon.xml      # Excel ribbon customization
├── Core/
│   ├── SnipEngine.cs      # Core snipping logic
│   ├── OCREngine.cs       # Enhanced OCR processing
│   ├── TableParser.cs     # Advanced table parsing
│   ├── ExcelHelper.cs     # Excel integration utilities
│   ├── MetadataManager.cs # Snip metadata management
│   └── SnipTypes.cs       # Data structures and enums
└── WebAssets/
    └── viewer.html        # Enhanced document viewer UI
```

### Technology Stack
- **Framework**: .NET Framework 4.8
- **UI**: Windows Forms with WebView2
- **Excel Integration**: Office Interop APIs
- **OCR**: Tesseract.js via WebView2
- **PDF Processing**: PDF.js
- **Data Storage**: Hidden Excel worksheets with JSON serialization
- **Build System**: MSBuild with PowerShell automation

### Performance Optimizations
- **Lazy Loading**: Components initialize only when needed
- **Image Preprocessing**: Optimized for OCR accuracy
- **Caching**: Document pages cached for faster navigation
- **Async Operations**: Non-blocking UI with progress indicators
- **Memory Management**: Proper disposal patterns and resource cleanup

## 🔧 Configuration

### OCR Settings
The OCR engine can be configured by modifying the parameters in `OCREngine.cs`:
```csharp
await worker.setParameters({
    tessedit_char_whitelist: '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz.,()$€£¥₹%+-=/:;!?@#&*[]{}|\"\'`~_<> \t\n\r',
    tessedit_pageseg_mode: Tesseract.PSM.AUTO,
    tessedit_ocr_engine_mode: Tesseract.OEM.LSTM_ONLY,
    preserve_interword_spaces: '1'
});
```

### Table Parser Settings
Table parsing strategies can be prioritized by modifying the order in `TableParser.cs`:
```csharp
var strategies = new Func<List<string>, TableData>[]
{
    ParseMarkdownTable,      // Highest priority
    ParseTabDelimited,       // High priority
    ParsePipeDelimited,      // High priority
    ParseCommaDelimited,     // Medium priority
    // ... other strategies
};
```

## 🐛 Troubleshooting

### Common Issues

#### Add-in Not Loading
1. **Check Excel version**: Requires Excel 2016 or later
2. **Verify COM registration**: Run `Install-SnipperClone.ps1` as Administrator
3. **Check Excel add-ins**: File → Options → Add-ins → COM Add-ins → Ensure SnipperClone is checked
4. **Review Event Viewer**: Look for loading errors in Windows Event Viewer

#### OCR Not Working
1. **Internet connection**: Tesseract.js requires CDN access for initial download
2. **WebView2 runtime**: Ensure Microsoft Edge WebView2 is installed
3. **Image quality**: Try higher resolution or better contrast images
4. **Timeout issues**: Large images may require longer processing time

#### Table Parsing Issues
1. **Document quality**: Ensure tables have clear structure and borders
2. **OCR accuracy**: Poor text recognition affects table detection
3. **Complex layouts**: Very complex tables may require manual adjustment
4. **File format**: PDFs generally work better than scanned images

#### Performance Issues
1. **File size**: Large documents (>50MB) may be slow to process
2. **Memory usage**: Close other applications if experiencing slowdowns
3. **Document complexity**: Very detailed documents require more processing time
4. **Hardware**: OCR processing is CPU-intensive

### Debug Mode
Enable verbose logging by setting the build configuration to Debug:
```powershell
.\Build-SnipperClone.ps1 -Configuration Debug -Verbose
```

### Log Files
- **Build logs**: `build.log` in the project directory
- **Runtime logs**: Check Visual Studio Output window or Debug console
- **Excel errors**: Windows Event Viewer → Applications and Services Logs → Microsoft Office Alerts

## 🤝 Contributing

### Development Setup
1. **Clone the repository**
2. **Install Visual Studio 2019/2022** with .NET Framework 4.8
3. **Install Office Developer Tools** for Visual Studio
4. **Open `SnipperClone.sln`** in Visual Studio
5. **Build and test** using the provided scripts

### Code Style
- Follow C# naming conventions
- Use XML documentation for public methods
- Include error handling and logging
- Write unit tests for new functionality

### Pull Request Process
1. Fork the repository
2. Create a feature branch
3. Make your changes with tests
4. Update documentation
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **DataSnipper** for the original concept and inspiration
- **Tesseract.js** team for the excellent OCR library
- **PDF.js** team for PDF processing capabilities
- **Microsoft** for Office Interop APIs and WebView2
- **Community contributors** for testing and feedback

## 📞 Support

- **Issues**: Report bugs and feature requests on GitHub Issues
- **Documentation**: Check the wiki for detailed guides
- **Community**: Join discussions in GitHub Discussions
- **Enterprise Support**: Contact for commercial licensing and support

---

**SnipperClone** - Bringing professional document analysis to Excel with enhanced performance and corporate-friendly deployment. 