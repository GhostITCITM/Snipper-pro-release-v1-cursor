# Snipper Pro - DataSnipper Clone

A powerful Excel add-in for extracting data from PDFs and images, designed for audit and finance professionals.

## Features

- **Real OCR Engine**: Extract text and numbers from images and PDFs
- **Multiple Snip Types**: Text, Sum, Table, Validation, and Exception snips
- **Excel Integration**: Full ribbon interface with DataSnipper-style formulas
- **Document Viewer**: Load and manage multiple documents simultaneously
- **Visual Selection**: Draw rectangles to select data regions
- **Professional UI**: Modern interface with color-coded content

## Installation

1. **Build the project**:
   ```
   .\build-snipper-pro.ps1
   ```

2. **Install (as Administrator)**:
   ```
   .\install-snipper-pro-complete.ps1
   ```

3. **Register COM components**:
   ```
   .\REGISTER_NOW.bat
   ```

4. **Verify installation**:
   ```
   .\verify-installation.ps1
   ```

## Usage

1. Open Excel
2. Look for "SNIPPER PRO" tab in ribbon
3. Click "Open Viewer" to load documents
4. Select snip mode (Text/Sum/Table/Validation/Exception)
5. Draw rectangles on documents to extract data
6. Data appears in Excel with DS formulas

## Project Structure

- `SnipperCloneCleanFinal/` - Main C# project source code
- `packages/` - NuGet packages and dependencies
- `SnipperPro.snk` - Strong name key for signing
- Build and installation scripts

## Requirements

- .NET Framework 4.8
- Microsoft Excel (COM Interop)
- Windows with PowerShell execution policy enabled

## Architecture

Built as a C# .NET Framework COM add-in with GUID: `D9A6E8B7-F3E1-47B0-B76B-C8DE050D1111`

Core components:
- **OCREngine**: Real text extraction from images
- **DocumentViewer**: PDF and image display
- **SnippingEngine**: Data extraction and Excel integration
- **RibbonInterface**: Excel ribbon UI

---

*DataSnipper-compatible Excel add-in for professional document analysis*
