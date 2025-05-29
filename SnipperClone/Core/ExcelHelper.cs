using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.Drawing;
using System.Diagnostics;

namespace SnipperClone.Core
{
    public class ExcelHelper
    {
        private readonly Microsoft.Office.Interop.Excel.Application _application;
        private const string SNIPS_WORKSHEET_NAME = "_SnipperClone_Metadata";

        public ExcelHelper(Microsoft.Office.Interop.Excel.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
            Debug.WriteLine("ExcelHelper: Initialized successfully");
        }

        public void WriteToSelectedCell(string value)
        {
            try
            {
                if (string.IsNullOrEmpty(value))
                {
                    Debug.WriteLine("ExcelHelper: Warning - Attempting to write empty value to cell");
                    return;
                }

                var selection = _application.Selection as Range;
                if (selection != null)
                {
                    // Validate Excel state
                    if (!ValidateExcelState())
                    {
                        throw new InvalidOperationException("Excel is not in a valid state for writing");
                    }

                    // Store original value for potential undo
                    var originalValue = selection.Value;
                    
                    // Disable screen updating for better performance
                    _application.ScreenUpdating = false;
                    
                    try
                    {
                        selection.Value = value;
                        selection.Columns.AutoFit();
                        
                        Debug.WriteLine($"ExcelHelper: Successfully wrote '{value}' to cell {selection.Address}");
                    }
                    finally
                    {
                        _application.ScreenUpdating = true;
                    }
                }
                else
                {
                    throw new InvalidOperationException("No cell is currently selected");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Failed to write to selected cell: {ex.Message}");
                throw new InvalidOperationException($"Failed to write to selected cell: {ex.Message}", ex);
            }
        }

        public void WriteTableToSelectedCell(TableData tableData)
        {
            try
            {
                if (tableData == null || tableData.RowCount == 0)
                {
                    Debug.WriteLine("ExcelHelper: Warning - No table data to write");
                    return;
                }

                var selection = _application.Selection as Range;
                if (selection == null)
                {
                    throw new InvalidOperationException("No cell is currently selected");
                }

                if (!ValidateExcelState())
                {
                    throw new InvalidOperationException("Excel is not in a valid state for writing");
                }

                Debug.WriteLine($"ExcelHelper: Writing table with {tableData.RowCount} rows and {tableData.ColumnCount} columns");

                // Disable screen updating for better performance
                _application.ScreenUpdating = false;
                _application.Calculation = XlCalculation.xlCalculationManual;

                try
                {
                    // Convert TableData to object array for Excel
                    object[,] values = new object[tableData.RowCount, tableData.ColumnCount];
                    for (int row = 0; row < tableData.RowCount; row++)
                    {
                        for (int col = 0; col < tableData.ColumnCount; col++)
                        {
                            var cellValue = tableData.Rows[row][col] ?? "";
                            
                            // Try to convert to number if it looks like one
                            if (!string.IsNullOrWhiteSpace(cellValue) && 
                                double.TryParse(cellValue.Replace(",", "").Replace("$", ""), out var numValue))
                            {
                                values[row, col] = numValue;
                            }
                            else
                            {
                                values[row, col] = cellValue;
                            }
                        }
                    }

                    // Resize range to fit table
                    var targetRange = selection.Resize[tableData.RowCount, tableData.ColumnCount];
                    targetRange.Value = values;

                    // Apply professional table formatting
                    ApplyProfessionalTableFormatting(targetRange, tableData);
                    
                    Debug.WriteLine($"ExcelHelper: Successfully wrote table ({tableData.RowCount}x{tableData.ColumnCount}) to {selection.Address}");
                }
                finally
                {
                    _application.ScreenUpdating = true;
                    _application.Calculation = XlCalculation.xlCalculationAutomatic;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Failed to write table to cell: {ex.Message}");
                throw new InvalidOperationException($"Failed to write table to cell: {ex.Message}", ex);
            }
        }

        private void ApplyProfessionalTableFormatting(Range targetRange, TableData tableData)
        {
            try
            {
                // Apply borders
                targetRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                targetRange.Borders.Weight = XlBorderWeight.xlThin;
                targetRange.Borders.Color = ColorTranslator.ToOle(Color.Gray);

                // Format header row if table has headers
                if (tableData.HasHeader && tableData.RowCount > 1)
                {
                    var headerRange = (Range)targetRange.Rows[1];
                    headerRange.Font.Bold = true;
                    headerRange.Font.Color = ColorTranslator.ToOle(Color.White);
                    headerRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(68, 114, 196)); // Professional blue
                    headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    
                    // Thicker border for header
                    headerRange.Borders.Weight = XlBorderWeight.xlMedium;
                    headerRange.Borders.Color = ColorTranslator.ToOle(Color.White);
                }

                // Apply alternating row colors for better readability
                if (tableData.RowCount > 2)
                {
                    var startRow = tableData.HasHeader ? 2 : 1;
                    for (int i = startRow; i <= tableData.RowCount; i += 2)
                    {
                        var rowRange = (Range)targetRange.Rows[i];
                        rowRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242)); // Light gray
                    }
                }

                // Auto-fit columns
                targetRange.Columns.AutoFit();

                // Apply number formatting to numeric columns
                for (int col = 1; col <= tableData.ColumnCount; col++)
                {
                    if (IsNumericColumn(tableData, col - 1))
                    {
                        var columnRange = (Range)targetRange.Columns[col];
                        columnRange.NumberFormat = "#,##0.00";
                        columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                    }
                }

                Debug.WriteLine("ExcelHelper: Applied professional table formatting");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Table formatting failed: {ex.Message}");
                // Continue even if formatting fails
            }
        }

        public void WriteTableToExcel(TableData tableData, string startCellAddress)
        {
            if (tableData == null || tableData.Rows.Count == 0)
            {
                Debug.WriteLine("ExcelHelper: No table data to write");
                return;
            }

            try
            {
                var workbook = _application.ActiveWorkbook;
                if (workbook == null)
                {
                    throw new InvalidOperationException("No active workbook");
                }

                var worksheet = workbook.ActiveSheet as Worksheet;
                if (worksheet == null)
                {
                    throw new InvalidOperationException("No active worksheet");
                }

                Debug.WriteLine($"ExcelHelper: Writing table with {tableData.Rows.Count} rows and {tableData.ColumnCount} columns starting at {startCellAddress}");

                // Disable screen updating for better performance
                _application.ScreenUpdating = false;
                _application.Calculation = XlCalculation.xlCalculationManual;

                try
                {
                    // Parse start cell address
                    var startRange = worksheet.Range[startCellAddress];
                    var startRow = startRange.Row;
                    var startCol = startRange.Column;
                    var currentRow = startRow;

                    // Write headers if present
                    if (tableData.HasHeader && tableData.Headers != null && tableData.Headers.Count > 0)
                    {
                        var headerEndCol = startCol + Math.Min(tableData.Headers.Count, tableData.ColumnCount) - 1;
                        var headerRange = worksheet.Range[
                            worksheet.Cells[currentRow, startCol],
                            worksheet.Cells[currentRow, headerEndCol]
                        ];

                        // Prepare header data
                        var headerData = new object[1, tableData.ColumnCount];
                        for (int i = 0; i < tableData.ColumnCount; i++)
                        {
                            headerData[0, i] = i < tableData.Headers.Count ? tableData.Headers[i] : $"Column {i + 1}";
                        }

                        headerRange.Value = headerData;

                        // Format headers professionally
                        headerRange.Font.Bold = true;
                        headerRange.Font.Color = ColorTranslator.ToOle(Color.White);
                        headerRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(68, 114, 196));
                        headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        headerRange.VerticalAlignment = XlVAlign.xlVAlignCenter;

                        // Add borders to headers
                        headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                        headerRange.Borders.Weight = XlBorderWeight.xlMedium;
                        headerRange.Borders.Color = ColorTranslator.ToOle(Color.White);

                        currentRow++;
                    }

                    // Write data rows
                    if (tableData.Rows.Count > 0)
                    {
                        var dataEndRow = currentRow + tableData.Rows.Count - 1;
                        var dataEndCol = startCol + tableData.ColumnCount - 1;
                        var dataRange = worksheet.Range[
                            worksheet.Cells[currentRow, startCol],
                            worksheet.Cells[dataEndRow, dataEndCol]
                        ];

                        // Prepare data array
                        var dataArray = new object[tableData.Rows.Count, tableData.ColumnCount];
                        for (int i = 0; i < tableData.Rows.Count; i++)
                        {
                            var row = tableData.Rows[i];
                            for (int j = 0; j < tableData.ColumnCount; j++)
                            {
                                var cellValue = j < row.Length ? row[j] : "";
                                
                                // Try to convert to number if it looks like one
                                if (!string.IsNullOrWhiteSpace(cellValue) && 
                                    double.TryParse(cellValue.Replace(",", "").Replace("$", ""), out var numValue))
                                {
                                    dataArray[i, j] = numValue;
                                }
                                else
                                {
                                    dataArray[i, j] = cellValue;
                                }
                            }
                        }

                        dataRange.Value = dataArray;

                        // Apply data formatting
                        dataRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                        dataRange.Borders.Weight = XlBorderWeight.xlThin;
                        dataRange.Borders.Color = ColorTranslator.ToOle(Color.Gray);

                        // Apply alternating row colors
                        for (int i = 0; i < tableData.Rows.Count; i += 2)
                        {
                            var rowRange = worksheet.Range[
                                worksheet.Cells[currentRow + i, startCol],
                                worksheet.Cells[currentRow + i, dataEndCol]
                            ];
                            rowRange.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242));
                        }

                        // Format numeric columns
                        for (int col = 0; col < tableData.ColumnCount; col++)
                        {
                            if (IsNumericColumn(tableData, col))
                            {
                                var columnRange = worksheet.Range[
                                    worksheet.Cells[currentRow, startCol + col],
                                    worksheet.Cells[dataEndRow, startCol + col]
                                ];
                                columnRange.NumberFormat = "#,##0.00";
                                columnRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
                            }
                        }

                        // Auto-fit columns
                        var fullRange = worksheet.Range[
                            worksheet.Cells[startRow, startCol],
                            worksheet.Cells[dataEndRow, dataEndCol]
                        ];
                        fullRange.Columns.AutoFit();
                    }

                    Debug.WriteLine($"ExcelHelper: Successfully wrote table to Excel at {startCellAddress}");
                }
                finally
                {
                    _application.ScreenUpdating = true;
                    _application.Calculation = XlCalculation.xlCalculationAutomatic;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Failed to write table to Excel: {ex.Message}");
                throw new InvalidOperationException($"Failed to write table to Excel: {ex.Message}", ex);
            }
        }

        private bool IsNumericColumn(TableData tableData, int columnIndex)
        {
            try
            {
                if (columnIndex >= tableData.ColumnCount)
                    return false;

                var numericCount = 0;
                var totalCount = 0;

                foreach (var row in tableData.Rows)
                {
                    if (columnIndex < row.Length)
                    {
                        var cellValue = row[columnIndex];
                        if (!string.IsNullOrWhiteSpace(cellValue))
                        {
                            totalCount++;
                            if (double.TryParse(cellValue.Replace(",", "").Replace("$", ""), out _))
                            {
                                numericCount++;
                            }
                        }
                    }
                }

                // Consider column numeric if more than 70% of values are numeric
                return totalCount > 0 && (double)numericCount / totalCount > 0.7;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error checking if column is numeric: {ex.Message}");
                return false;
            }
        }

        public void WriteTickMark()
        {
            try
            {
                var selection = _application.Selection as Range;
                if (selection != null)
                {
                    selection.Value = "✓";
                    selection.Font.Color = ColorTranslator.ToOle(Color.Green);
                    selection.Font.Size = 14;
                    selection.Font.Bold = true;
                    selection.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    selection.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    
                    Debug.WriteLine($"ExcelHelper: Applied tick mark to {selection.Address}");
                }
                else
                {
                    throw new InvalidOperationException("No cell is currently selected");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Failed to write tick mark: {ex.Message}");
                throw new InvalidOperationException($"Failed to write tick mark: {ex.Message}", ex);
            }
        }

        public void WriteCrossMark()
        {
            try
            {
                var selection = _application.Selection as Range;
                if (selection != null)
                {
                    selection.Value = "✗";
                    selection.Font.Color = ColorTranslator.ToOle(Color.Red);
                    selection.Font.Size = 14;
                    selection.Font.Bold = true;
                    selection.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    selection.VerticalAlignment = XlVAlign.xlVAlignCenter;
                    
                    Debug.WriteLine($"ExcelHelper: Applied cross mark to {selection.Address}");
                }
                else
                {
                    throw new InvalidOperationException("No cell is currently selected");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Failed to write cross mark: {ex.Message}");
                throw new InvalidOperationException($"Failed to write cross mark: {ex.Message}", ex);
            }
        }

        public string GetSelectedCellAddress()
        {
            try
            {
                var selection = _application.Selection as Range;
                if (selection != null)
                {
                    var address = selection.Address[false, false];
                    Debug.WriteLine($"ExcelHelper: Current selection: {address}");
                    return address;
                }
                
                Debug.WriteLine("ExcelHelper: No cell currently selected");
                return null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error getting selected cell address: {ex.Message}");
                return null;
            }
        }

        public void LogSnip(SnipRecord record)
        {
            try
            {
                if (record == null)
                {
                    Debug.WriteLine("ExcelHelper: Warning - Null snip record provided");
                    return;
                }

                var workbook = _application.ActiveWorkbook;
                if (workbook == null)
                {
                    Debug.WriteLine("ExcelHelper: No active workbook for logging snip");
                    return;
                }

                var snipsWorksheet = GetOrCreateSnipsWorksheet(workbook);
                if (snipsWorksheet == null)
                {
                    Debug.WriteLine("ExcelHelper: Failed to get or create snips worksheet");
                    return;
                }

                // Find the next empty row
                var lastRow = GetLastRow(snipsWorksheet);
                var nextRow = lastRow + 1;

                // If this is the first entry, add headers
                if (lastRow == 1 && string.IsNullOrEmpty(GetCellValue(snipsWorksheet, 1, 1)?.ToString()))
                {
                    snipsWorksheet.Cells[1, 1] = "Cell Address";
                    snipsWorksheet.Cells[1, 2] = "Mode";
                    snipsWorksheet.Cells[1, 3] = "Extracted Text";
                    snipsWorksheet.Cells[1, 4] = "Document Name";
                    snipsWorksheet.Cells[1, 5] = "Page Number";
                    snipsWorksheet.Cells[1, 6] = "Rectangle";
                    snipsWorksheet.Cells[1, 7] = "Created At";
                    snipsWorksheet.Cells[1, 8] = "Processing Time (ms)";
                    snipsWorksheet.Cells[1, 9] = "Confidence";

                    // Format headers
                    var headerRange = snipsWorksheet.Range["A1:I1"];
                    headerRange.Font.Bold = true;
                    headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                    headerRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                    nextRow = 2;
                }

                // Add the snip record
                ((Range)snipsWorksheet.Cells[nextRow, 1]).Value2 = record.CellAddress;
                ((Range)snipsWorksheet.Cells[nextRow, 2]).Value2 = record.Mode.ToString();
                ((Range)snipsWorksheet.Cells[nextRow, 3]).Value2 = record.ExtractedText;
                ((Range)snipsWorksheet.Cells[nextRow, 4]).Value2 = record.DocumentName;
                ((Range)snipsWorksheet.Cells[nextRow, 5]).Value2 = record.PageNumber;
                ((Range)snipsWorksheet.Cells[nextRow, 6]).Value2 = $"{record.Rectangle.X},{record.Rectangle.Y},{record.Rectangle.Width},{record.Rectangle.Height}";
                ((Range)snipsWorksheet.Cells[nextRow, 7]).Value2 = record.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss");
                ((Range)snipsWorksheet.Cells[nextRow, 8]).Value2 = record.ProcessingTimeMs;
                ((Range)snipsWorksheet.Cells[nextRow, 9]).Value2 = record.Confidence;

                // Auto-fit columns
                snipsWorksheet.Columns.AutoFit();

                Debug.WriteLine($"ExcelHelper: Logged snip record for {record.CellAddress}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Failed to log snip: {ex.Message}");
                // Don't throw here as logging failure shouldn't break the snip operation
            }
        }

        public SnipRecord FindSnipByCell(string cellAddress)
        {
            try
            {
                if (string.IsNullOrEmpty(cellAddress))
                    return null;

                var workbook = _application.ActiveWorkbook;
                if (workbook == null)
                    return null;

                var snipsWorksheet = GetSnipsWorksheet(workbook);
                if (snipsWorksheet == null)
                    return null;

                // Find the record with matching cell address
                var lastRow = GetLastRow(snipsWorksheet);
                
                for (int row = 2; row <= lastRow; row++) // Start from row 2 to skip headers
                {
                    var recordCellAddress = GetCellValue(snipsWorksheet, row, 1)?.ToString();
                    if (string.Equals(recordCellAddress, cellAddress, StringComparison.OrdinalIgnoreCase))
                    {
                        // Found the record, create SnipRecord object
                        var record = new SnipRecord
                        {
                            CellAddress = recordCellAddress,
                            Mode = Enum.TryParse<SnipMode>(GetCellValue(snipsWorksheet, row, 2)?.ToString(), out var mode) ? mode : SnipMode.None,
                            ExtractedText = GetCellValue(snipsWorksheet, row, 3)?.ToString() ?? "",
                            DocumentName = GetCellValue(snipsWorksheet, row, 4)?.ToString() ?? "",
                            PageNumber = int.TryParse(GetCellValue(snipsWorksheet, row, 5)?.ToString(), out var pageNum) ? pageNum : 1
                        };

                        // Parse rectangle
                        var rectangleStr = GetCellValue(snipsWorksheet, row, 6)?.ToString();
                        if (!string.IsNullOrEmpty(rectangleStr))
                        {
                            var parts = rectangleStr.Split(',');
                            if (parts.Length == 4 &&
                                int.TryParse(parts[0].ToString(), out var x) &&
                                int.TryParse(parts[1].ToString(), out var y) &&
                                int.TryParse(parts[2].ToString(), out var width) &&
                                int.TryParse(parts[3].ToString(), out var height))
                            {
                                record.Rectangle = new Rectangle(x, y, width, height);
                            }
                        }

                        // Parse created date
                        if (DateTime.TryParse(GetCellValue(snipsWorksheet, row, 7)?.ToString(), out var createdAt))
                        {
                            record.CreatedAt = createdAt;
                        }

                        // Parse processing time
                        if (int.TryParse(GetCellValue(snipsWorksheet, row, 8)?.ToString(), out var processingTime))
                        {
                            record.ProcessingTimeMs = processingTime;
                        }

                        // Parse confidence
                        if (double.TryParse(GetCellValue(snipsWorksheet, row, 9)?.ToString(), out var confidence))
                        {
                            record.Confidence = confidence;
                        }

                        Debug.WriteLine($"ExcelHelper: Found snip record for {cellAddress}");
                        return record;
                    }
                }

                Debug.WriteLine($"ExcelHelper: No snip record found for {cellAddress}");
                return null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error finding snip by cell: {ex.Message}");
                return null;
            }
        }

        public List<SnipRecord> GetAllSnips()
        {
            var records = new List<SnipRecord>();

            try
            {
                var workbook = _application.ActiveWorkbook;
                if (workbook == null)
                    return records;

                var snipsWorksheet = GetSnipsWorksheet(workbook);
                if (snipsWorksheet == null)
                    return records;

                var lastRow = GetLastRow(snipsWorksheet);
                
                for (int row = 2; row <= lastRow; row++) // Start from row 2 to skip headers
                {
                    var cellAddress = GetCellValue(snipsWorksheet, row, 1)?.ToString();
                    if (!string.IsNullOrEmpty(cellAddress))
                    {
                        var record = FindSnipByCell(cellAddress);
                        if (record != null)
                        {
                            records.Add(record);
                        }
                    }
                }

                Debug.WriteLine($"ExcelHelper: Retrieved {records.Count} snip records");
                return records;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error getting all snips: {ex.Message}");
                return records;
            }
        }

        public void DeleteSnip(string cellAddress)
        {
            try
            {
                if (string.IsNullOrEmpty(cellAddress))
                    return;

                var workbook = _application.ActiveWorkbook;
                if (workbook == null)
                    return;

                var snipsWorksheet = GetSnipsWorksheet(workbook);
                if (snipsWorksheet == null)
                    return;

                var lastRow = GetLastRow(snipsWorksheet);
                
                for (int row = 2; row <= lastRow; row++) // Start from row 2 to skip headers
                {
                    var recordCellAddress = GetCellValue(snipsWorksheet, row, 1)?.ToString();
                    if (string.Equals(recordCellAddress, cellAddress, StringComparison.OrdinalIgnoreCase))
                    {
                        // Delete the row
                        ((Range)snipsWorksheet.Rows[row]).Delete();
                        Debug.WriteLine($"ExcelHelper: Deleted snip record for {cellAddress}");
                        return;
                    }
                }

                Debug.WriteLine($"ExcelHelper: No snip record found to delete for {cellAddress}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error deleting snip: {ex.Message}");
                throw new InvalidOperationException($"Failed to delete snip: {ex.Message}", ex);
            }
        }

        private Worksheet GetOrCreateSnipsWorksheet(Workbook workbook)
        {
            Worksheet snipsSheet = null;
            try
            {
                snipsSheet = GetSnipsWorksheet(workbook);
                if (snipsSheet != null)
                {
                    return snipsSheet;
                }

                // If not found, create it
                snipsSheet = (Worksheet)workbook.Sheets.Add();
                snipsSheet.Name = SNIPS_WORKSHEET_NAME;
                snipsSheet.Visible = XlSheetVisibility.xlSheetVeryHidden; 

                // Optional: Add headers for JSON data storage
                // ((Range)snipsSheet.Cells[1, 1]).Value2 = "SnipRecordJson";
                Debug.WriteLine($"ExcelHelper: Created snips worksheet '{SNIPS_WORKSHEET_NAME}'.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error getting or creating snips worksheet: {ex.Message}");
            }
            return snipsSheet;
        }

        private Worksheet GetSnipsWorksheet(Workbook workbook)
        {
            try
            {
                foreach (object sheetObj in workbook.Sheets)
                {
                    Worksheet sheet = sheetObj as Worksheet;
                    if (sheet != null && sheet.Name == SNIPS_WORKSHEET_NAME)
                    {
                        return sheet;
                    }
                }
                Debug.WriteLine($"ExcelHelper: Snips worksheet '{SNIPS_WORKSHEET_NAME}' not found.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error getting snips worksheet: {ex.Message}");
            }
            return null;
        }

        public void HighlightCell(string cellAddress)
        {
            try
            {
                if (string.IsNullOrEmpty(cellAddress))
                    return;

                var workbook = _application.ActiveWorkbook;
                if (workbook == null)
                    return;

                var worksheet = workbook.ActiveSheet as Worksheet;
                if (worksheet == null)
                    return;

                var range = worksheet.Range[cellAddress];
                if (range != null)
                {
                    // Apply highlight formatting
                    range.Interior.Color = ColorTranslator.ToOle(Color.Yellow);
                    range.Borders.LineStyle = XlLineStyle.xlContinuous;
                    range.Borders.Weight = XlBorderWeight.xlMedium;
                    range.Borders.Color = ColorTranslator.ToOle(Color.Orange);

                    Debug.WriteLine($"ExcelHelper: Highlighted cell {cellAddress}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error highlighting cell: {ex.Message}");
            }
        }

        public void NavigateToSnip(SnipRecord snip)
        {
            try
            {
                if (snip == null || string.IsNullOrEmpty(snip.CellAddress))
                    return;

                var workbook = _application.ActiveWorkbook;
                if (workbook == null)
                    return;

                var worksheet = workbook.ActiveSheet as Worksheet;
                if (worksheet == null)
                    return;

                // Navigate to the cell
                var range = worksheet.Range[snip.CellAddress];
                if (range != null)
                {
                    range.Select();
                    
                    // Temporarily highlight the cell
                    var originalColor = range.Interior.Color;
                    range.Interior.Color = ColorTranslator.ToOle(Color.LightBlue);
                    
                    // Create a timer to remove highlight after 2 seconds
                    var timer = new System.Windows.Forms.Timer();
                    timer.Interval = 2000;
                    timer.Tick += (sender, e) =>
                    {
                        try
                        {
                            range.Interior.Color = originalColor;
                            timer.Stop();
                            timer.Dispose();
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"ExcelHelper: Error removing highlight: {ex.Message}");
                        }
                    };
                    timer.Start();

                    Debug.WriteLine($"ExcelHelper: Navigated to snip at {snip.CellAddress}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error navigating to snip: {ex.Message}");
            }
        }

        public bool ValidateExcelState()
        {
            try
            {
                if (_application == null)
                {
                    Debug.WriteLine("ExcelHelper: Excel application is null");
                    return false;
                }

                var workbook = _application.ActiveWorkbook;
                if (workbook == null)
                {
                    Debug.WriteLine("ExcelHelper: No active workbook");
                    return false;
                }

                var worksheet = workbook.ActiveSheet as Worksheet;
                if (worksheet == null)
                {
                    Debug.WriteLine("ExcelHelper: No active worksheet");
                    return false;
                }

                // Check if Excel is in edit mode
                try
                {
                    var testValue = _application.Selection;
                    if (testValue == null)
                    {
                        Debug.WriteLine("ExcelHelper: Excel appears to be in edit mode");
                        return false;
                    }
                }
                catch
                {
                    Debug.WriteLine("ExcelHelper: Excel is not accessible (possibly in edit mode)");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ExcelHelper: Error validating Excel state: {ex.Message}");
                return false;
            }
        }

        private int GetLastRow(Worksheet sheet)
        {
            if (sheet == null) return 1;
            var lastCell = (Range)sheet.Cells[sheet.Rows.Count, 1];
            var last = (Range)lastCell.End[XlDirection.xlUp];
            return last.Row;
        }

        private static object GetCellValue(Worksheet sheet, int row, int col)
        {
            return ((Range)sheet.Cells[row, col]).Value2;
        }
    }
} 