# Search Functionality Test Guide

## 🔧 Testing the DataSnipper-Style Search Implementation

### 1. **Start Excel with the Add-in**
- Run: `.\start_excel_with_snipper.bat`
- This ensures the updated DLL is loaded with search functionality

### 2. **Open the Snipper Pro Document Viewer**
- Open Excel
- Click on the "Snipper Pro" tab in the ribbon
- Click "Open Viewer" button

### 3. **Load a PDF Document**
- In the Document Viewer, click "Load Document(s)"
- Select an IFRS PDF file (or any PDF with text)
- Wait for the document to load and text extraction to complete

### 4. **Test Search Functionality**

#### **Basic Search Test:**
1. Type "IFRS" in the search box
2. Press Enter or click the 🔍 button
3. **Expected Results:**
   - Status should show "Found X matches for 'IFRS'"
   - Yellow highlights should appear on the document
   - Current result should have orange highlight/border
   - Navigation buttons (◀ ▶) should be enabled
   - Results counter should show "1/X"

#### **Navigation Test:**
1. Click the ▶ (next) button to cycle through results
2. Click the ◀ (previous) button to go back
3. **Expected Results:**
   - Orange highlight should move to the next/previous result
   - Counter should update (e.g., "2/5", "3/5")
   - Viewport should center on the found text

#### **Keyboard Shortcuts Test:**
1. Press `Ctrl+F` - should focus the search box
2. Press `F3` - should go to next result
3. Press `Shift+F3` - should go to previous result
4. Press `Escape` - should close search mode

#### **Close Search Test:**
1. Click the ✕ button
2. **Expected Results:**
   - All highlights should disappear
   - Search box should clear
   - Navigation buttons should be disabled

### 5. **Advanced Tests**

#### **Case Insensitive Search:**
- Search for "ifrs" (lowercase) - should still find "IFRS"
- Search for "Revenue" - should find "revenue", "REVENUE", etc.

#### **Partial Word Search:**
- Search for "Rev" - should find words containing "Revenue"
- Search for "15" - should find "IFRS 15"

#### **Cross-Document Search:**
- Load multiple PDF files
- Search for a term that appears in different documents
- Navigation should switch between documents automatically

### 6. **Visual Verification**

#### **Ribbon Icons:**
- Check that snip category buttons have colored squares:
  - Text Snip: Blue square with professional gradient
  - Sum Snip: Purple square
  - Table Snip: Purple square
  - Validation: Green square
  - Exception: Red square

#### **Search UI Elements:**
- Search box in the toolbar (second row)
- 🔍 search button
- ◀ ▶ navigation buttons
- Results counter label
- ✕ close button

### 7. **Troubleshooting**

#### **If Search Doesn't Work:**
1. Check the status bar for error messages
2. Ensure the PDF loaded successfully
3. Try a simple search term like "a" or "the"
4. Check that text extraction completed (status should not show "Searching..." indefinitely)

#### **If No Highlights Appear:**
1. Ensure search found results (check status and counter)
2. Try zooming in/out to refresh the display
3. Check that the correct page is displayed

#### **If Highlights Are in Wrong Position:**
1. This is expected with the current implementation
2. The highlighting uses estimated text positions
3. Focus on verifying that searches work and navigation functions

### 8. **Log Verification**
- Check the application logs for search-related entries:
  - "Starting text extraction for search"
  - "Text extraction completed"
  - "Starting search for: '[term]'"
  - "FOUND '[term]' in word: '[word]'"
  - "Search completed. Found X total results"

## ✅ Success Criteria

The search functionality is working correctly if:
1. ✅ Search finds and counts results correctly
2. ✅ Yellow highlights appear on the document
3. ✅ Orange highlight marks the current result
4. ✅ Navigation buttons cycle through results
5. ✅ Keyboard shortcuts work (Ctrl+F, F3, etc.)
6. ✅ Search can be closed with ✕ or Escape
7. ✅ Ribbon shows colored category squares
8. ✅ Text extraction happens automatically on document load

This implementation provides a solid foundation for DataSnipper-style search functionality! 
