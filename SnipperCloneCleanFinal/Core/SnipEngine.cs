using System;
using System.Drawing;
using System.Threading.Tasks;
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
            if (!ocrResult.Success)
                return SnipResult.CreateError($"OCR failed: {ocrResult.ErrorMessage}");

            _excelHelper.WriteToSelectedCell(ocrResult.Text);
            return SnipResult.CreateSuccess(ocrResult.Text, ocrResult);
        }

        private SnipResult ProcessSumSnip(Bitmap imageData)
        {
            if (imageData == null)
                return SnipResult.CreateError("No image data provided");

            var ocrResult = _ocrEngine.RecognizeText(imageData);
            if (!ocrResult.Success)
                return SnipResult.CreateError($"OCR failed: {ocrResult.ErrorMessage}");

            // Simple sum calculation - extract numbers from text
            var numbers = System.Text.RegularExpressions.Regex.Matches(ocrResult.Text, @"\d+\.?\d*");
            double sum = 0;
            foreach (System.Text.RegularExpressions.Match match in numbers)
            {
                if (double.TryParse(match.Value, out double number))
                    sum += number;
            }

            _excelHelper.WriteToSelectedCell(sum.ToString());
            return SnipResult.CreateSuccess(sum.ToString(), ocrResult);
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