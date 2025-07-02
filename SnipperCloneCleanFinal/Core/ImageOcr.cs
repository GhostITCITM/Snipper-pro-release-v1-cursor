using System.Drawing;

namespace SnipperCloneCleanFinal.Core
{
    internal static class ImageOcr
    {
        private static OCREngine _ocrEngine = new OCREngine();

        public static string Read(Bitmap bmp)
        {
            if (!_ocrEngine.Initialize())
            {
                return "OCR not available";
            }

            var result = _ocrEngine.RecognizeText(bmp);
            return result?.Text ?? "No text detected";
        }
    }
} 