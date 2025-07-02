using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Office = Microsoft.Office.Core;

namespace SnipperCloneCleanFinal.Core
{
    public class ExcelHelper : IDisposable
    {
        private readonly Excel.Application _application;
        private bool _disposed;

        public ExcelHelper(Excel.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public void WriteToSelectedCell(string value)
        {
            try
            {
                var selection = _application.Selection as Excel.Range;
                if (selection != null)
                {
                    selection.Value2 = value;
                    Debug.WriteLine($"ExcelHelper: Written '{value}' to cell {selection.Address}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error writing to cell: {ex.Message}");
                throw;
            }
        }

        public void WriteTableToRange(TableData tableData, Excel.Range startCell)
        {
            if (tableData?.Rows == null || tableData.Rows.Count == 0)
                return;

            try
            {
                var worksheet = startCell.Worksheet;
                int headerRows = tableData.HasHeader && tableData.Headers != null ? 1 : 0;
                int rows = tableData.Rows.Count + headerRows;
                int cols = tableData.ColumnCount;

                // Build data array (headers + rows) for efficient assignment
                object[,] values = new object[rows, cols];

                int currentRow = 0;
                if (headerRows == 1)
                {
                    for (int c = 0; c < cols; c++)
                        values[0, c] = c < tableData.Headers.Count ? tableData.Headers[c] : null;
                    currentRow = 1;
                }

                for (int r = 0; r < tableData.Rows.Count; r++, currentRow++)
                {
                    var rowData = tableData.Rows[r];
                    for (int c = 0; c < cols; c++)
                        values[currentRow, c] = c < rowData.Length ? rowData[c] : null;
                }

                var endCell = startCell.Offset[rows - 1, cols - 1];
                var writeRange = worksheet.Range[startCell, endCell];
                writeRange.Value2 = values;

                Debug.WriteLine($"ExcelHelper: Written table data to range starting at {startCell.Address}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error writing table: {ex.Message}");
                throw;
            }
        }

        public string GetSelectedCellAddress()
        {
            try
            {
                var selection = _application.Selection as Excel.Range;
                return selection?.Address ?? string.Empty;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error getting selected cell address: {ex.Message}");
                return string.Empty;
            }
        }

        public void InsertPictureAtSelection(Bitmap image)
        {
            if (image == null) return;
            try
            {
                var selection = _application.Selection as Excel.Range;
                if (selection == null) return;

                // Save image to temporary PNG
                string tempPath = Path.Combine(Path.GetTempPath(), $"snip_{Guid.NewGuid()}.png");
                image.Save(tempPath, ImageFormat.Png);

                Excel.Worksheet sheet = selection.Worksheet;
                // Calculate placement – align with top-left of the selection
                float left = (float)selection.Left;
                float top = (float)selection.Top;

                // Scale image so it fits within a reasonable width (max cell width * 5)
                float maxWidth = (float)(selection.Width * 5);
                float scale = 1f;
                if (image.Width > maxWidth)
                {
                    scale = maxWidth / image.Width;
                }

                float width = image.Width * scale;
                float height = image.Height * scale;

                Office.MsoTriState linkToFile = Office.MsoTriState.msoFalse;
                Office.MsoTriState saveWithDoc = Office.MsoTriState.msoTrue;

                var pic = sheet.Shapes.AddPicture(tempPath, linkToFile, saveWithDoc, left, top, width, height);
                pic.LockAspectRatio = Office.MsoTriState.msoTrue;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error inserting picture: {ex.Message}");
                throw;
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
            }
        }
    }
}
