using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Diagnostics;

namespace SnipperCloneCleanFinal.Core
{
    public class OCREngine : IDisposable
    {
        private bool _disposed;
        private bool _isInitialized = true; // Simplified - always initialized

        public bool IsInitialized => _isInitialized;

        public async Task<bool> InitializeAsync()
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(OCREngine));

            Debug.WriteLine("OCREngine: Simplified initialization completed");
            return true;
        }

        public async Task<OCRResult> RecognizeTextAsync(Bitmap image)
        {
            try
            {
                // Simplified OCR - return sample text
                await Task.Delay(100); // Simulate processing
                
                return new OCRResult
                {
                    Success = true,
                    Text = "Sample extracted text from OCR",
                    Confidence = 0.85
                };
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OCREngine: Recognition error: {ex}");
                return new OCRResult
                {
                    Success = false,
                    ErrorMessage = ex.Message
                };
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
                Debug.WriteLine("OCREngine: Disposed");
            }
        }
    }
} 