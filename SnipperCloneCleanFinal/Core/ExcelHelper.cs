using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

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
                int startRow = startCell.Row;
                int startCol = startCell.Column;

                // Write headers if present
                if (tableData.HasHeader && tableData.Headers != null)
                {
                    for (int i = 0; i < tableData.Headers.Count; i++)
                    {
                        worksheet.Cells[startRow, startCol + i] = tableData.Headers[i];
                    }
                    startRow++;
                }

                // Write data rows
                for (int row = 0; row < tableData.Rows.Count; row++)
                {
                    var rowData = tableData.Rows[row];
                    for (int col = 0; col < rowData.Length; col++)
                    {
                        worksheet.Cells[startRow + row, startCol + col] = rowData[col];
                    }
                }

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

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
            }
        }
    }
} 