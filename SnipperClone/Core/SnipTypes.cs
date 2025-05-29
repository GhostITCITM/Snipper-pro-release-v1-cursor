using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace SnipperClone.Core
{
    public enum SnipMode
    {
        None,
        Text,
        Sum,
        Table,
        Validation,
        Exception
    }

    public struct Rectangle
    {
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }

        public static Rectangle Empty => new Rectangle(0, 0, 0, 0);

        public Rectangle(int x, int y, int width, int height)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        public System.Drawing.Rectangle ToDrawingRectangle()
        {
            return new System.Drawing.Rectangle(X, Y, Width, Height);
        }

        public bool IsEmpty => Width <= 0 || Height <= 0;

        public override string ToString()
        {
            return $"{{X={X}, Y={Y}, Width={Width}, Height={Height}}}";
        }
    }

    public class SnipRecord
    {
        public string CellAddress { get; set; }
        public int PageNumber { get; set; }
        public Rectangle Rectangle { get; set; }
        public SnipMode Mode { get; set; }
        public string ExtractedText { get; set; }
        public DateTime CreatedAt { get; set; }
        public string DocumentName { get; set; }
        public double Confidence { get; set; }
        public long ProcessingTimeMs { get; set; }

        public SnipRecord()
        {
            CreatedAt = DateTime.Now;
            CellAddress = string.Empty;
            DocumentName = string.Empty;
            ExtractedText = string.Empty;
            PageNumber = 1;
            Rectangle = Rectangle.Empty;
            Confidence = 0.0;
            ProcessingTimeMs = 0;
        }
    }

    public class OCRResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string Text { get; set; }
        public double Confidence { get; set; }
        public string[] Numbers { get; set; }
        public double Sum { get; set; }
        public TableData TableData { get; set; }

        public OCRResult()
        {
            Success = false;
            ErrorMessage = string.Empty;
            Text = string.Empty;
            Numbers = new string[0];
            Confidence = 0.0;
            Sum = 0.0;
        }
    }

    public class TableData
    {
        public List<string[]> Rows { get; set; }
        public List<string> Headers { get; set; }
        public int ColumnCount { get; set; }
        public bool HasHeader { get; set; }
        public bool HasHeaders
        {
            get => HasHeader;
            set => HasHeader = value;
        }

        // Number of data rows in the table (convenience property used by ExcelHelper)
        public int RowCount => Rows?.Count ?? 0;

        public TableData()
        {
            Rows = new List<string[]>();
            Headers = null;
            ColumnCount = 0;
            HasHeader = false;
        }

        public TableData(List<string[]> rows, string[] headers = null, bool hasHeader = false)
        {
            Rows = rows ?? new List<string[]>();
            Headers = headers?.ToList();
            HasHeader = hasHeader;
            
            // Calculate column count
            if (Rows.Count > 0)
            {
                ColumnCount = Rows.Max(row => row?.Length ?? 0);
            }
            else if (Headers != null)
            {
                ColumnCount = Headers.Count;
            }
            else
            {
                ColumnCount = 0;
            }
        }

        public void AddRow(string[] row)
        {
            if (row != null)
            {
                Rows.Add(row);
                ColumnCount = Math.Max(ColumnCount, row.Length);
            }
        }

        public void SetHeaders(string[] headers)
        {
            Headers = headers?.ToList();
            HasHeader = headers != null && headers.Length > 0;
            if (HasHeader)
            {
                ColumnCount = Math.Max(ColumnCount, headers.Length);
            }
        }

        public string GetCell(int row, int column)
        {
            if (row >= 0 && row < Rows.Count && 
                column >= 0 && column < (Rows[row]?.Length ?? 0))
            {
                return Rows[row][column] ?? string.Empty;
            }
            return string.Empty;
        }

        public void SetCell(int row, int column, string value)
        {
            if (row >= 0 && row < Rows.Count)
            {
                var currentRow = Rows[row];
                if (currentRow != null && column >= 0 && column < currentRow.Length)
                {
                    currentRow[column] = value ?? string.Empty;
                }
            }
        }

        public bool IsEmpty => Rows.Count == 0 || ColumnCount == 0;

        public override string ToString()
        {
            return $"Table: {Rows.Count} rows × {ColumnCount} columns{(HasHeader ? " (with header)" : "")}";
        }
    }

    public class SnipResult
    {
        public bool Success { get; set; }
        public string Value { get; set; }
        public string ErrorMessage { get; set; }
        public OCRResult OCRData { get; set; }

        public SnipResult()
        {
            Success = false;
            Value = string.Empty;
            ErrorMessage = string.Empty;
        }

        public static SnipResult CreateSuccess(string value, OCRResult ocrData = null)
        {
            return new SnipResult
            {
                Success = true,
                Value = value,
                OCRData = ocrData
            };
        }

        public static SnipResult CreateError(string errorMessage)
        {
            return new SnipResult
            {
                Success = false,
                ErrorMessage = errorMessage
            };
        }
    }
} 