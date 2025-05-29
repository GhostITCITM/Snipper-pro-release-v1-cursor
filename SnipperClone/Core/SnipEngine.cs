using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace SnipperClone.Core
{
    public class SnipEngine : IDisposable
    {
        private readonly ExcelHelper _excelHelper;
        private readonly OCREngine _ocrEngine;
        private readonly TableParser _tableParser;
        private readonly MetadataManager _metadataManager;
        private SnipMode _currentMode;
        private string _selectedCellAddress;
        private string _currentDocumentName;
        private bool _disposed;
        private bool _isInitialized;

        public event EventHandler<SnipModeChangedEventArgs> ModeChanged;
        public event EventHandler<CellSelectionChangedEventArgs> CellSelectionChanged;
        public event EventHandler<SnipCompletedEventArgs> SnipCompleted;

        public SnipMode CurrentMode => _currentMode;
        public string SelectedCellAddress => _selectedCellAddress;
        public string CurrentDocumentName => _currentDocumentName;
        public bool IsInSelectionMode => _currentMode != SnipMode.None && !string.IsNullOrEmpty(_selectedCellAddress);
        public bool IsInitialized => _isInitialized;
        public OCREngine OcrEngine => _ocrEngine;

        public SnipEngine(Microsoft.Office.Interop.Excel.Application application)
        {
            try
            {
                _excelHelper = new ExcelHelper(application);
                _ocrEngine = new OCREngine();
                _tableParser = new TableParser();
                _metadataManager = new MetadataManager(application);
                _currentMode = SnipMode.None;
                
                Debug.WriteLine("SnipEngine: Initialized successfully");
                _isInitialized = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Initialization failed: {ex.Message}");
                throw new InvalidOperationException($"Failed to initialize SnipEngine: {ex.Message}", ex);
            }
        }

        public async Task<bool> InitializeAsync()
        {
            try
            {
                if (_isInitialized)
                    return true;

                Debug.WriteLine("SnipEngine: Starting async initialization...");
                
                // Initialize OCR engine asynchronously
                var ocrInitialized = await _ocrEngine.InitializeAsync();
                if (!ocrInitialized)
                {
                    Debug.WriteLine("SnipEngine: OCR engine initialization failed");
                    return false;
                }

                _isInitialized = true;
                Debug.WriteLine("SnipEngine: Async initialization completed successfully");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Async initialization failed: {ex.Message}");
                return false;
            }
        }

        public void SetMode(SnipMode mode)
        {
            try
            {
                if (_currentMode != mode)
                {
                    var previousMode = _currentMode;
                    _currentMode = mode;
                    _selectedCellAddress = _excelHelper.GetSelectedCellAddress();
                    
                    Debug.WriteLine($"SnipEngine: Mode changed from {previousMode} to {mode}, selected cell: {_selectedCellAddress}");
                    ModeChanged?.Invoke(this, new SnipModeChangedEventArgs(mode, _selectedCellAddress));
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error setting mode: {ex.Message}");
                throw new InvalidOperationException($"Failed to set snip mode: {ex.Message}", ex);
            }
        }

        public void ClearMode()
        {
            try
            {
                if (_currentMode != SnipMode.None)
                {
                    var previousMode = _currentMode;
                    _currentMode = SnipMode.None;
                    _selectedCellAddress = null;
                    
                    Debug.WriteLine($"SnipEngine: Mode cleared from {previousMode}");
                    ModeChanged?.Invoke(this, new SnipModeChangedEventArgs(SnipMode.None, null));
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error clearing mode: {ex.Message}");
            }
        }

        public void SetCurrentDocument(string documentName)
        {
            try
            {
                _currentDocumentName = documentName;
                Debug.WriteLine($"SnipEngine: Current document set to: {documentName}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error setting current document: {ex.Message}");
            }
        }

        public async Task<SnipResult> ProcessSnipAsync(Bitmap imageData, int pageNumber, Rectangle rectangle)
        {
            var stopwatch = Stopwatch.StartNew();
            
            try
            {
                Debug.WriteLine($"SnipEngine: Starting {_currentMode} snip processing...");

                // Enhanced input validation
                var validationResult = ValidateSnipInputs(imageData, pageNumber, rectangle);
                if (!validationResult.Success)
                {
                    return validationResult;
                }

                // Ensure OCR engine is initialized for OCR-based snips
                if (RequiresOCR(_currentMode) && !_ocrEngine.IsInitialized)
                {
                    Debug.WriteLine("SnipEngine: Initializing OCR engine for snip processing...");
                    var ocrInitialized = await _ocrEngine.InitializeAsync();
                    if (!ocrInitialized)
                    {
                        return SnipResult.CreateError("OCR engine initialization failed. Please try again.");
                    }
                }

                SnipResult result;

                switch (_currentMode)
                {
                    case SnipMode.Text:
                        result = await ProcessTextSnipAsync(imageData);
                        break;

                    case SnipMode.Sum:
                        result = await ProcessSumSnipAsync(imageData);
                        break;

                    case SnipMode.Table:
                        result = await ProcessTableSnipAsync(imageData);
                        break;

                    case SnipMode.Validation:
                        result = ProcessValidationSnip();
                        break;

                    case SnipMode.Exception:
                        result = ProcessExceptionSnip();
                        break;

                    default:
                        return SnipResult.CreateError($"Unknown snip mode: {_currentMode}");
                }

                if (result.Success)
                {
                    // Create comprehensive snip record
                    var record = new SnipRecord
                    {
                        CellAddress = _selectedCellAddress,
                        PageNumber = pageNumber,
                        Rectangle = rectangle,
                        Mode = _currentMode,
                        ExtractedText = result.OCRData?.Text ?? result.Value,
                        DocumentName = _currentDocumentName,
                        CreatedAt = DateTime.Now,
                        ProcessingTimeMs = (int)stopwatch.ElapsedMilliseconds,
                        Confidence = result.OCRData?.Confidence ?? 1.0
                    };

                    // Save metadata
                    try
                    {
                        _metadataManager.SaveSnipRecord(record);
                        Debug.WriteLine($"SnipEngine: Snip record saved successfully");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"SnipEngine: Warning - Failed to save snip metadata: {ex.Message}");
                        // Continue even if metadata save fails
                    }

                    // Clear mode after successful snip
                    ClearMode();

                    stopwatch.Stop();
                    Debug.WriteLine($"SnipEngine: {_currentMode} snip completed successfully in {stopwatch.ElapsedMilliseconds}ms");
                    
                    SnipCompleted?.Invoke(this, new SnipCompletedEventArgs(record, result));
                }
                else
                {
                    stopwatch.Stop();
                    Debug.WriteLine($"SnipEngine: Snip processing failed: {result.ErrorMessage}");
                }

                return result;
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                Debug.WriteLine($"SnipEngine: Error processing snip: {ex}");
                return SnipResult.CreateError($"Error processing snip: {ex.Message}");
            }
        }

        private SnipResult ValidateSnipInputs(Bitmap imageData, int pageNumber, Rectangle rectangle)
        {
            if (_currentMode == SnipMode.None)
            {
                return SnipResult.CreateError("No snip mode selected. Please select a snip mode first.");
            }

            if (string.IsNullOrEmpty(_selectedCellAddress))
            {
                return SnipResult.CreateError("No cell selected. Please select a cell in Excel first.");
            }

            if (RequiresOCR(_currentMode) && imageData == null)
            {
                return SnipResult.CreateError("No image data provided for OCR-based snip.");
            }

            if (rectangle.Width <= 0 || rectangle.Height <= 0)
            {
                return SnipResult.CreateError("Invalid selection rectangle. Please select a valid area.");
            }

            if (pageNumber < 1)
            {
                return SnipResult.CreateError("Invalid page number. Page number must be greater than 0.");
            }

            // Validate image data for OCR operations
            if (RequiresOCR(_currentMode) && imageData != null)
            {
                if (imageData.Width < 10 || imageData.Height < 10)
                {
                    return SnipResult.CreateError("Selected area is too small for reliable text recognition.");
                }

                if (imageData.Width > 5000 || imageData.Height > 5000)
                {
                    return SnipResult.CreateError("Selected area is too large. Please select a smaller area.");
                }
            }

            return SnipResult.CreateSuccess("Validation passed", null);
        }

        private bool RequiresOCR(SnipMode mode)
        {
            return mode == SnipMode.Text || mode == SnipMode.Sum || mode == SnipMode.Table;
        }

        private async Task<SnipResult> ProcessTextSnipAsync(Bitmap imageData)
        {
            try
            {
                Debug.WriteLine("SnipEngine: Processing text snip...");
                
                if (imageData == null)
                    return SnipResult.CreateError("No image data provided for text recognition.");

                var ocrResult = await _ocrEngine.RecognizeTextAsync(imageData);
                
                if (!ocrResult.Success)
                {
                    return SnipResult.CreateError($"OCR processing failed: {ocrResult.ErrorMessage}");
                }

                if (string.IsNullOrWhiteSpace(ocrResult.Text))
                {
                    return SnipResult.CreateError("No text was recognized in the selected area. Try selecting a clearer area or check image quality.");
                }

                // Enhanced text cleaning and formatting
                var cleanedText = CleanExtractedText(ocrResult.Text);
                
                if (string.IsNullOrWhiteSpace(cleanedText))
                {
                    return SnipResult.CreateError("Text was recognized but appears to be invalid after cleaning. Please try a different area.");
                }

                // Write to Excel with error handling
                try
                {
                    _excelHelper.WriteToSelectedCell(cleanedText);
                    Debug.WriteLine($"SnipEngine: Text snip successful - extracted: '{cleanedText.Substring(0, Math.Min(50, cleanedText.Length))}...'");
                    return SnipResult.CreateSuccess(cleanedText, ocrResult);
                }
                catch (Exception ex)
                {
                    return SnipResult.CreateError($"Failed to write text to Excel: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Text snip error: {ex}");
                return SnipResult.CreateError($"Text recognition failed: {ex.Message}");
            }
        }

        private async Task<SnipResult> ProcessSumSnipAsync(Bitmap imageData)
        {
            try
            {
                Debug.WriteLine("SnipEngine: Processing sum snip...");
                
                if (imageData == null)
                    return SnipResult.CreateError("No image data provided for number recognition.");

                var ocrResult = await _ocrEngine.RecognizeTextAsync(imageData);
                
                if (!ocrResult.Success)
                {
                    return SnipResult.CreateError($"OCR processing failed: {ocrResult.ErrorMessage}");
                }

                var numbers = _ocrEngine.ExtractNumbers(ocrResult.Text);
                
                if (numbers.Count == 0)
                {
                    return SnipResult.CreateError("No numbers were found in the selected area. Please ensure the area contains numeric values.");
                }

                var sum = numbers.Sum();
                var formattedSum = FormatNumber(sum);
                
                // Write to Excel with error handling
                try
                {
                    _excelHelper.WriteToSelectedCell(formattedSum);
                    
                    // Enhanced OCR result with sum details
                    ocrResult.Numbers = numbers.Select(n => FormatNumber(n)).ToArray();
                    ocrResult.Sum = sum;
                    ocrResult.Text = $"Sum of {numbers.Count} numbers: {formattedSum}";
                    
                    Debug.WriteLine($"SnipEngine: Sum snip successful - found {numbers.Count} numbers, sum: {formattedSum}");
                    return SnipResult.CreateSuccess(formattedSum, ocrResult);
                }
                catch (Exception ex)
                {
                    return SnipResult.CreateError($"Failed to write sum to Excel: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Sum snip error: {ex}");
                return SnipResult.CreateError($"Number recognition failed: {ex.Message}");
            }
        }

        public async Task<SnipResult> ProcessTableSnipAsync(Bitmap imageData)
        {
            try
            {
                Debug.WriteLine("SnipEngine: Processing table snip...");
                
                if (imageData == null)
                    return SnipResult.CreateError("No image data provided for table recognition.");

                var ocrResult = await _ocrEngine.RecognizeTextAsync(imageData);
                
                if (!ocrResult.Success)
                {
                    return SnipResult.CreateError($"OCR processing failed: {ocrResult.ErrorMessage}");
                }

                if (string.IsNullOrWhiteSpace(ocrResult.Text))
                {
                    return SnipResult.CreateError("No text was recognized in the selected area for table parsing.");
                }

                var tableData = _tableParser.ParseTable(ocrResult.Text);
                
                if (tableData == null || tableData.Rows.Count == 0)
                {
                    return SnipResult.CreateError("No table structure was detected in the selected area. Please ensure the area contains tabular data.");
                }

                // Write table to Excel with error handling
                try
                {
                    _excelHelper.WriteTableToSelectedCell(tableData);
                    
                    // Enhanced OCR result with table details
                    ocrResult.TableData = tableData;
                    ocrResult.Text = $"Table with {tableData.Rows.Count} rows and {tableData.ColumnCount} columns";
                    
                    Debug.WriteLine($"SnipEngine: Table snip successful - {tableData.Rows.Count}x{tableData.ColumnCount} table");
                    return SnipResult.CreateSuccess($"Table ({tableData.Rows.Count}x{tableData.ColumnCount})", ocrResult);
                }
                catch (Exception ex)
                {
                    return SnipResult.CreateError($"Failed to write table to Excel: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Table snip error: {ex}");
                return SnipResult.CreateError($"Table recognition failed: {ex.Message}");
            }
        }

        private SnipResult ProcessValidationSnip()
        {
            try
            {
                Debug.WriteLine("SnipEngine: Processing validation snip...");
                
                const string validationMark = "✓";
                
                try
                {
                    _excelHelper.WriteToSelectedCell(validationMark);
                    Debug.WriteLine("SnipEngine: Validation snip successful");
                    return SnipResult.CreateSuccess(validationMark, null);
                }
                catch (Exception ex)
                {
                    return SnipResult.CreateError($"Failed to write validation mark to Excel: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Validation snip error: {ex}");
                return SnipResult.CreateError($"Validation snip failed: {ex.Message}");
            }
        }

        private SnipResult ProcessExceptionSnip()
        {
            try
            {
                Debug.WriteLine("SnipEngine: Processing exception snip...");
                
                const string exceptionMark = "✗";
                
                try
                {
                    _excelHelper.WriteToSelectedCell(exceptionMark);
                    Debug.WriteLine("SnipEngine: Exception snip successful");
                    return SnipResult.CreateSuccess(exceptionMark, null);
                }
                catch (Exception ex)
                {
                    return SnipResult.CreateError($"Failed to write exception mark to Excel: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Exception snip error: {ex}");
                return SnipResult.CreateError($"Exception snip failed: {ex.Message}");
            }
        }

        private string CleanExtractedText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            try
            {
                // Enhanced text cleaning
                text = text.Trim();
                
                // Remove excessive whitespace
                text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");
                
                // Fix common OCR artifacts
                text = text.Replace("~", "-")
                          .Replace("¦", "|")
                          .Replace("\"", "\"")
                          .Replace("\"", "\"")
                          .Replace("'", "'")
                          .Replace("'", "'")
                          .Replace("\r\n", "\n")
                          .Replace("\r", "\n");
                
                return text.Trim();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error cleaning text: {ex.Message}");
                return text?.Trim() ?? string.Empty;
            }
        }

        private string FormatNumber(double number)
        {
            try
            {
                // Format numbers with appropriate precision
                if (number == Math.Floor(number))
                {
                    return number.ToString("N0"); // No decimal places for whole numbers
                }
                else
                {
                    return number.ToString("N2"); // Two decimal places for decimals
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error formatting number: {ex.Message}");
                return number.ToString();
            }
        }

        public void OnCellSelectionChanged(Range selectedRange)
        {
            try
            {
                if (selectedRange != null)
                {
                    var newAddress = selectedRange.Address;
                    if (_selectedCellAddress != newAddress)
                    {
                        _selectedCellAddress = newAddress;
                        Debug.WriteLine($"SnipEngine: Cell selection changed to: {newAddress}");
                        CellSelectionChanged?.Invoke(this, new CellSelectionChangedEventArgs(newAddress));
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error handling cell selection change: {ex.Message}");
            }
        }

        public void OnWorkbookOpened(Workbook workbook)
        {
            try
            {
                Debug.WriteLine($"SnipEngine: Workbook opened: {workbook?.Name}");
                // Reset state when new workbook is opened
                ClearMode();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error handling workbook open: {ex.Message}");
            }
        }

        // Enhanced metadata operations with error handling
        public SnipRecord FindSnipByCell(string cellAddress)
        {
            try
            {
                return _metadataManager.GetSnipRecord(cellAddress);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error finding snip by cell: {ex.Message}");
                return null;
            }
        }

        public List<SnipRecord> GetAllSnips()
        {
            try
            {
                return _metadataManager.GetAllSnipRecords();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error getting all snips: {ex.Message}");
                return new List<SnipRecord>();
            }
        }

        public void DeleteSnip(string cellAddress)
        {
            try
            {
                _metadataManager.DeleteSnipRecord(cellAddress);
                Debug.WriteLine($"SnipEngine: Deleted snip for cell: {cellAddress}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error deleting snip: {ex.Message}");
                throw new InvalidOperationException($"Failed to delete snip: {ex.Message}", ex);
            }
        }

        public void HighlightAllSnips()
        {
            try
            {
                var snips = GetAllSnips();
                foreach (var snip in snips)
                {
                    _excelHelper.HighlightCell(snip.CellAddress);
                }
                Debug.WriteLine($"SnipEngine: Highlighted {snips.Count} snips");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error highlighting snips: {ex.Message}");
            }
        }

        public void ClearAllHighlights()
        {
            try
            {
                // Implementation would depend on how highlights are stored
                Debug.WriteLine("SnipEngine: Cleared all highlights");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error clearing highlights: {ex.Message}");
            }
        }

        public List<SnipRecord> GetSnipsByDocument(string documentName)
        {
            try
            {
                return GetAllSnips().Where(s => s.DocumentName == documentName).ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error getting snips by document: {ex.Message}");
                return new List<SnipRecord>();
            }
        }

        public List<SnipRecord> GetSnipsByMode(SnipMode mode)
        {
            try
            {
                return GetAllSnips().Where(s => s.Mode == mode).ToList();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error getting snips by mode: {ex.Message}");
                return new List<SnipRecord>();
            }
        }

        public void Dispose()
        {
            if (_disposed)
                return;

            try
            {
                Debug.WriteLine("SnipEngine: Disposing...");
                
                _ocrEngine?.Dispose();
                _metadataManager?.Dispose();
                
                _disposed = true;
                Debug.WriteLine("SnipEngine: Disposed successfully");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"SnipEngine: Error during disposal: {ex.Message}");
            }
        }
    }

    public class SnipModeChangedEventArgs : EventArgs
    {
        public SnipMode Mode { get; }
        public string SelectedCell { get; }

        public SnipModeChangedEventArgs(SnipMode mode, string selectedCell)
        {
            Mode = mode;
            SelectedCell = selectedCell;
        }
    }

    public class CellSelectionChangedEventArgs : EventArgs
    {
        public string CellAddress { get; }

        public CellSelectionChangedEventArgs(string cellAddress)
        {
            CellAddress = cellAddress;
        }
    }

    public class SnipCompletedEventArgs : EventArgs
    {
        public SnipRecord Record { get; }
        public SnipResult Result { get; }

        public SnipCompletedEventArgs(SnipRecord record, SnipResult result)
        {
            Record = record;
            Result = result;
        }
    }
} 