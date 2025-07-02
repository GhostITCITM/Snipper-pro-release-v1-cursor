using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using SnipperCloneCleanFinal.Infrastructure;

namespace SnipperCloneCleanFinal.Core
{
    public class SnipEngine : IDisposable
    {
        private readonly ExcelHelper _excelHelper;
        private readonly OCREngine _ocrEngine;
        private SnipMode _currentMode;
        private string _selectedCellAddress;
        private bool _disposed;

        public event EventHandler<SnipModeChangedEventArgs> ModeChanged;
        public event EventHandler<CellSelectionChangedEventArgs> CellSelectionChanged;
        public event EventHandler<SnipCompletedEventArgs> SnipCompleted;

        public SnipMode CurrentMode => _currentMode;
        public string SelectedCellAddress => _selectedCellAddress;

        public SnipEngine(Excel.Application application)
        {
            _excelHelper = new ExcelHelper(application);
            _ocrEngine = new OCREngine();
            _currentMode = SnipMode.None;
            Logger.Info("SnipEngine initialized");
        }

        public void SetMode(SnipMode mode)
        {
            var oldMode = _currentMode;
            _currentMode = mode;
            ModeChanged?.Invoke(this, new SnipModeChangedEventArgs(oldMode, mode));
            Logger.Info($"SnipEngine mode changed from {oldMode} to {mode}");
        }

        public void OnCellSelectionChanged(string cellAddress)
        {
            var oldAddress = _selectedCellAddress;
            _selectedCellAddress = cellAddress;
            CellSelectionChanged?.Invoke(this, new CellSelectionChangedEventArgs(oldAddress, cellAddress));
        }

        public SnipResult ProcessSnip(Bitmap imageData, int pageNumber, System.Drawing.Rectangle rectangle)
        {
            try
            {
                Logger.Info($"Processing {_currentMode} snip...");

                switch (_currentMode)
                {
                    case SnipMode.Text:
                        return ProcessTextSnip(imageData);
                    case SnipMode.Sum:
                        return ProcessSumSnip(imageData);
                    case SnipMode.Exception:
                        return ProcessExceptionSnip();
                    case SnipMode.Validation:
                        return ProcessValidationSnip();
                    case SnipMode.Image:
                        return ProcessImageSnip(imageData);
                    default:
                        return SnipResult.CreateError("Unknown snip mode");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Error processing snip: {ex.Message}", ex);
                return SnipResult.CreateError($"Error processing snip: {ex.Message}");
            }
        }

        private SnipResult ProcessTextSnip(Bitmap imageData)
        {
            if (imageData == null)
                return SnipResult.CreateError("No image data provided");

            var ocrResult = _ocrEngine.RecognizeText(imageData);
            
            // With our new fallback system, we should always get some result
            // Only fail if we get a completely empty result
            if (!ocrResult.Success && string.IsNullOrWhiteSpace(ocrResult.Text))
            {
                Logger.Info($"OCR returned no text: {ocrResult.ErrorMessage}");
                // Even in this case, provide a meaningful result
                _excelHelper.WriteToSelectedCell("[No text detected]");
                return SnipResult.CreateSuccess("[No text detected]", ocrResult);
            }

            // Use whatever text we got, even if Success = false
            var textToUse = !string.IsNullOrWhiteSpace(ocrResult.Text) ? ocrResult.Text : "[Text detection attempted]";
            _excelHelper.WriteToSelectedCell(textToUse);
            return SnipResult.CreateSuccess(textToUse, ocrResult);
        }

        private SnipResult ProcessSumSnip(Bitmap imageData)
        {
            if (imageData == null)
                return SnipResult.CreateError("No image data provided");

            var ocrResult = _ocrEngine.RecognizeText(imageData);
            
            // Use the enhanced number extraction from OCR result
            double sum = 0;
            int numberCount = 0;
            var processedNumbers = new List<double>();
            
            // First try to use pre-extracted numbers from OCR
            if (ocrResult.Numbers != null && ocrResult.Numbers.Length > 0)
            {
                foreach (var numStr in ocrResult.Numbers)
                {
                    if (double.TryParse(numStr, out double number))
                    {
                        sum += number;
                        numberCount++;
                        processedNumbers.Add(number);
                    }
                }
            }
            
            // If no numbers from OCR extraction, try manual extraction
            if (numberCount == 0 && !string.IsNullOrWhiteSpace(ocrResult.Text))
            {
                // Enhanced pattern to catch more number formats
                var patterns = new[]
                {
                    @"\$?[\d,]+\.?\d*",                    // Currency with optional $
                    @"-?\d+\.?\d+",                        // Decimals with optional negative
                    @"-?\d{1,3}(?:,\d{3})*(?:\.\d+)?",    // Thousands with optional decimal
                    @"\(\d+\.?\d*\)",                      // Accounting format negative
                    @"-?\d+"                               // Plain integers
                };
                
                var foundNumbers = new HashSet<string>();
                foreach (var pattern in patterns)
                {
                    var matches = System.Text.RegularExpressions.Regex.Matches(ocrResult.Text, pattern);
                    foreach (System.Text.RegularExpressions.Match match in matches)
                    {
                        var value = match.Value.Trim();
                        if (!foundNumbers.Contains(value) && NumberHelper.TryParseFlexible(value, out double number))
                        {
                            foundNumbers.Add(value);
                            sum += number;
                            numberCount++;
                            processedNumbers.Add(number);
                        }
                    }
                }
            }

            if (numberCount > 0)
            {
                // Format the sum nicely
                var formattedSum = sum % 1 == 0 ? sum.ToString("N0") : sum.ToString("N2");
                _excelHelper.WriteToSelectedCell(formattedSum);
                
                // Log the numbers found for debugging
                Logger.Info($"Sum snip found {numberCount} numbers: {string.Join(", ", processedNumbers)}. Total: {formattedSum}");
                
                return SnipResult.CreateSuccess(formattedSum, ocrResult);
            }
            else
            {
                // No numbers found, but don't fail completely
                _excelHelper.WriteToSelectedCell("[No numbers detected]");
                Logger.Info($"Sum snip found no numbers in text: {ocrResult.Text}");
                return SnipResult.CreateSuccess("[No numbers detected]", ocrResult);
            }
        }

        private SnipResult ProcessExceptionSnip()
        {
            const string exceptionMark = "✗";
            _excelHelper.WriteToSelectedCell(exceptionMark);
            return SnipResult.CreateSuccess(exceptionMark);
        }

        private SnipResult ProcessValidationSnip()
        {
            const string validationMark = "✓";
            _excelHelper.WriteToSelectedCell(validationMark);
            return SnipResult.CreateSuccess(validationMark);
        }

        private SnipResult ProcessImageSnip(Bitmap imageData)
        {
            if (imageData == null)
                return SnipResult.CreateError("No image data provided");

            try
            {
                var preprocessor = new ImagePreprocessor();
                using var cleaned = preprocessor.Clean(imageData);
                _excelHelper.InsertPictureAtSelection(cleaned);
            }
            catch (Exception ex)
            {
                Logger.Error($"Error inserting image snip: {ex.Message}", ex);
                return SnipResult.CreateError($"Error inserting image: {ex.Message}");
            }

            return SnipResult.CreateSuccess("[Image inserted]");
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
                _excelHelper?.Dispose();
                _ocrEngine?.Dispose();
                Logger.Info("SnipEngine disposed");
            }
        }
    }
} 