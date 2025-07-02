using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace SnipperCloneCleanFinal.Core
{
    /// <summary>
    /// Advanced image pre-processing pipeline to enhance images for better OCR accuracy.
    /// Applies sharpening, contrast enhancement, and high-quality scaling.
    /// </summary>
    public class ImagePreprocessor
    {
        private const float TARGET_DPI = 300f;
        
        public Bitmap Clean(Bitmap input)
        {
            if (input == null) throw new ArgumentNullException(nameof(input));
            
            // Create a high-quality copy with proper DPI
            var output = new Bitmap(input.Width, input.Height, PixelFormat.Format32bppArgb);
            output.SetResolution(TARGET_DPI, TARGET_DPI);
            
            using (var graphics = Graphics.FromImage(output))
            {
                // Use highest quality settings
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.CompositingQuality = CompositingQuality.HighQuality;
                graphics.SmoothingMode = SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                
                // Draw the original image
                graphics.DrawImage(input, 0, 0, input.Width, input.Height);
            }
            
            // Apply image enhancements
            output = ApplySharpening(output);
            output = EnhanceContrast(output);
            
            return output;
        }
        
        private Bitmap ApplySharpening(Bitmap image)
        {
            // Create gentler sharpening kernel
            float[,] kernel = {
                { 0, -0.5f, 0 },
                { -0.5f,  3, -0.5f },
                { 0, -0.5f, 0 }
            };
            
            return ApplyConvolution(image, kernel);
        }
        
        private Bitmap EnhanceContrast(Bitmap image)
        {
            var result = new Bitmap(image.Width, image.Height, image.PixelFormat);
            result.SetResolution(image.HorizontalResolution, image.VerticalResolution);
            
            // Create ImageAttributes for contrast adjustment
            var imageAttributes = new ImageAttributes();
            
            // Contrast enhancement matrix
            float contrast = 1.2f; // Gentle contrast increase by 20%
            float translation = 0.001f * (1.0f - contrast);
            
            ColorMatrix colorMatrix = new ColorMatrix(new float[][]
            {
                new float[] {contrast, 0, 0, 0, 0},
                new float[] {0, contrast, 0, 0, 0},
                new float[] {0, 0, contrast, 0, 0},
                new float[] {0, 0, 0, 1, 0},
                new float[] {translation, translation, translation, 0, 1}
            });
            
            imageAttributes.SetColorMatrix(colorMatrix, ColorMatrixFlag.Default, ColorAdjustType.Bitmap);
            
            using (var graphics = Graphics.FromImage(result))
            {
                graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphics.DrawImage(image, new System.Drawing.Rectangle(0, 0, image.Width, image.Height),
                    0, 0, image.Width, image.Height, GraphicsUnit.Pixel, imageAttributes);
            }
            
            image.Dispose();
            return result;
        }
        
        private Bitmap ApplyConvolution(Bitmap image, float[,] kernel)
        {
            var result = new Bitmap(image.Width, image.Height, image.PixelFormat);
            result.SetResolution(image.HorizontalResolution, image.VerticalResolution);
            
            var sourceData = image.LockBits(new System.Drawing.Rectangle(0, 0, image.Width, image.Height),
                ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            var resultData = result.LockBits(new System.Drawing.Rectangle(0, 0, result.Width, result.Height),
                ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
            
            int bytes = Math.Abs(sourceData.Stride) * image.Height;
            byte[] sourceBuffer = new byte[bytes];
            byte[] resultBuffer = new byte[bytes];
            
            System.Runtime.InteropServices.Marshal.Copy(sourceData.Scan0, sourceBuffer, 0, bytes);
            
            int kernelSize = kernel.GetLength(0);
            int kernelOffset = kernelSize / 2;
            
            for (int y = kernelOffset; y < image.Height - kernelOffset; y++)
            {
                for (int x = kernelOffset; x < image.Width - kernelOffset; x++)
                {
                    float rSum = 0, gSum = 0, bSum = 0;
                    
                    for (int ky = 0; ky < kernelSize; ky++)
                    {
                        for (int kx = 0; kx < kernelSize; kx++)
                        {
                            int pixelX = x + kx - kernelOffset;
                            int pixelY = y + ky - kernelOffset;
                            int pixelIndex = (pixelY * sourceData.Stride) + (pixelX * 4);
                            
                            float kernelValue = kernel[ky, kx];
                            bSum += sourceBuffer[pixelIndex] * kernelValue;
                            gSum += sourceBuffer[pixelIndex + 1] * kernelValue;
                            rSum += sourceBuffer[pixelIndex + 2] * kernelValue;
                        }
                    }
                    
                    int index = (y * sourceData.Stride) + (x * 4);
                    resultBuffer[index] = (byte)Math.Max(0, Math.Min(255, bSum));
                    resultBuffer[index + 1] = (byte)Math.Max(0, Math.Min(255, gSum));
                    resultBuffer[index + 2] = (byte)Math.Max(0, Math.Min(255, rSum));
                    resultBuffer[index + 3] = sourceBuffer[index + 3]; // Alpha
                }
            }
            
            System.Runtime.InteropServices.Marshal.Copy(resultBuffer, 0, resultData.Scan0, bytes);
            
            image.UnlockBits(sourceData);
            result.UnlockBits(resultData);
            
            image.Dispose();
            return result;
        }
    }
} 