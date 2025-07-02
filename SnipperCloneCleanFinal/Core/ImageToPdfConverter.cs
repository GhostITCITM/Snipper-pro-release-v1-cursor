using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace SnipperCloneCleanFinal.Core
{
    /// <summary>
    /// Helper that converts common raster image formats (PNG, JPG, BMP, GIF, TIFF) into a single-page PDF.
    /// If the Tesseract executable is available on the PATH (or alongside the application), the PDF will
    /// contain an invisible text layer which makes the document searchable / selectable by Pdfium.  When
    /// Tesseract is unavailable, the image is merely embedded as-is, so text extraction will return empty.
    /// </summary>
    public static class ImageToPdfConverter
    {
        private const double TARGET_DPI = 300.0; // Optimal DPI for OCR without blur
        
        /// <summary>
        /// Converts the provided image file to a PDF and returns the generated PDF path.
        /// The output file will be placed in the system TEMP directory with a random GUID filename.
        /// </summary>
        /// <param name="imagePath">Absolute path to the source image.</param>
        /// <returns>Absolute path to the newly created PDF.</returns>
        public static string Convert(string imagePath)
        {
            if (string.IsNullOrWhiteSpace(imagePath) || !File.Exists(imagePath))
                throw new FileNotFoundException("Image file not found", imagePath);

            // Build a temp PDF path
            string tempPdf = Path.Combine(Path.GetTempPath(), $"Snipper_{Guid.NewGuid():N}.pdf");

            // 1) Try Tesseract CLI with "pdf" output mode  -----------------------------
            try
            {
                // Determine tesseract executable path – assume it is either on PATH or in same folder as the running EXE
                string exeName = Environment.OSVersion.Platform == PlatformID.Win32NT ? "tesseract.exe" : "tesseract";
                string exePath = exeName; // rely on PATH by default

                // If not on PATH, probe alongside application
                if (!IsExeAvailable(exePath))
                {
                    string local = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, exeName);
                    if (File.Exists(local)) exePath = local;
                }

                if (IsExeAvailable(exePath))
                {
                    // First, create a high-quality version of the image
                    string highQualityImage = CreateHighQualityImage(imagePath);
                    
                    try
                    {
                        // Tesseract wants output path without extension
                        string outBase = Path.Combine(Path.GetTempPath(), $"Snipper_{Guid.NewGuid():N}");
                        var psi = new ProcessStartInfo
                        {
                            FileName = exePath,
                            Arguments = $"\"{highQualityImage}\" \"{outBase}\" pdf --dpi {TARGET_DPI}",
                            CreateNoWindow = true,
                            UseShellExecute = false,
                            RedirectStandardError = true,
                            RedirectStandardOutput = true
                        };
                        using var proc = Process.Start(psi);
                        proc.WaitForExit(30000); // 30 second timeout
                        if (proc.ExitCode == 0)
                        {
                            string createdPdf = outBase + ".pdf";
                            if (File.Exists(createdPdf))
                            {
                                File.Move(createdPdf, tempPdf);
                                return tempPdf;
                            }
                        }
                    }
                    finally
                    {
                        // Clean up temporary high-quality image
                        if (File.Exists(highQualityImage))
                            File.Delete(highQualityImage);
                    }
                }
            }
            catch
            {
                // swallow and fall back
            }

            // 2) Fallback: embed the bitmap into a PDF page with PDFsharp  --------------
            using var img = XImage.FromFile(imagePath);
            var doc = new PdfDocument();
            var page = doc.AddPage();

            // Set high DPI for better quality
            double dpiX = img.HorizontalResolution > 1 ? img.HorizontalResolution : TARGET_DPI;
            double dpiY = img.VerticalResolution > 1 ? img.VerticalResolution : TARGET_DPI;

            // Ensure we use at least TARGET_DPI for high quality
            if (dpiX < TARGET_DPI) dpiX = TARGET_DPI;
            if (dpiY < TARGET_DPI) dpiY = TARGET_DPI;

            double widthPts = (img.PixelWidth / dpiX) * 72.0;
            double heightPts = (img.PixelHeight / dpiY) * 72.0;
            page.Width = XUnit.FromPoint(widthPts);
            page.Height = XUnit.FromPoint(heightPts);

            using (var gfx = XGraphics.FromPdfPage(page))
            {
                gfx.DrawImage(img, 0, 0, page.Width, page.Height);
            }
            doc.Save(tempPdf);
            return tempPdf;
        }

        private static string CreateHighQualityImage(string imagePath)
        {
            string tempPath = Path.Combine(Path.GetTempPath(), $"Snipper_HQ_{Guid.NewGuid():N}.png");
            
            using (var original = Image.FromFile(imagePath))
            {
                // Create a high-quality bitmap with proper DPI settings
                using (var bitmap = new Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb))
                {
                    bitmap.SetResolution((float)TARGET_DPI, (float)TARGET_DPI);
                    
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        // Use highest quality settings
                        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        graphics.CompositingQuality = CompositingQuality.HighQuality;
                        graphics.SmoothingMode = SmoothingMode.HighQuality;
                        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        
                        graphics.DrawImage(original, 0, 0, original.Width, original.Height);
                    }
                    
                    // Save as PNG to preserve quality
                    bitmap.Save(tempPath, ImageFormat.Png);
                }
            }
            
            return tempPath;
        }

        private static bool IsExeAvailable(string exe)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = exe,
                    Arguments = "--version",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                };
                using var proc = Process.Start(psi);
                proc.WaitForExit(3000);
                return proc.ExitCode == 0;
            }
            catch
            {
                return false;
            }
        }
    }
} 