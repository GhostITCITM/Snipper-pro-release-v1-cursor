using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using System.Drawing;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace SnipperClone.Core
{
    public class MetadataManager : IDisposable
    {
        private readonly Microsoft.Office.Interop.Excel.Application _application;
        private const string METADATA_SHEET_NAME = "_SnipperClone_Metadata";
        private const string METADATA_RANGE_NAME = "SnipperClone_Data";
        private bool _disposed;

        public MetadataManager(Microsoft.Office.Interop.Excel.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        public void SaveSnipRecord(SnipRecord record)
        {
            try
            {
                var workbook = _application.ActiveWorkbook;
                if (workbook == null) return;

                var metadataSheet = GetOrCreateMetadataSheet(workbook);
                var existingRecords = LoadAllRecords(metadataSheet);

                // Remove any existing record for the same cell
                existingRecords.RemoveAll(r => r.CellAddress.Equals(record.CellAddress, StringComparison.OrdinalIgnoreCase));
                
                // Add the new record
                existingRecords.Add(record);

                // Save back to sheet
                SaveAllRecords(metadataSheet, existingRecords);

                System.Diagnostics.Debug.WriteLine($"Saved snip record for cell {record.CellAddress}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving snip record: {ex.Message}");
            }
        }

        public SnipRecord GetSnipRecord(string cellAddress)
        {
            try
            {
                var workbook = _application.ActiveWorkbook;
                if (workbook == null) return null;

                var metadataSheet = GetMetadataSheet(workbook);
                if (metadataSheet == null) return null;

                var records = LoadAllRecords(metadataSheet);
                return records.FirstOrDefault(r => r.CellAddress.Equals(cellAddress, StringComparison.OrdinalIgnoreCase));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting snip record: {ex.Message}");
                return null;
            }
        }

        public List<SnipRecord> GetAllSnipRecords()
        {
            try
            {
                var workbook = _application.ActiveWorkbook;
                if (workbook == null) return new List<SnipRecord>();

                var metadataSheet = GetMetadataSheet(workbook);
                if (metadataSheet == null) return new List<SnipRecord>();

                return LoadAllRecords(metadataSheet);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting all snip records: {ex.Message}");
                return new List<SnipRecord>();
            }
        }

        public void DeleteSnipRecord(string cellAddress)
        {
            try
            {
                var workbook = _application.ActiveWorkbook;
                if (workbook == null) return;

                var metadataSheet = GetMetadataSheet(workbook);
                if (metadataSheet == null) return;

                var records = LoadAllRecords(metadataSheet);
                var recordsToKeep = records.Where(r => !r.CellAddress.Equals(cellAddress, StringComparison.OrdinalIgnoreCase)).ToList();

                if (recordsToKeep.Count != records.Count)
                {
                    SaveAllRecords(metadataSheet, recordsToKeep);
                    System.Diagnostics.Debug.WriteLine($"Deleted snip record for cell {cellAddress}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error deleting snip record: {ex.Message}");
            }
        }

        public void ClearAllRecords()
        {
            try
            {
                var workbook = _application.ActiveWorkbook;
                if (workbook == null) return;

                var metadataSheet = GetMetadataSheet(workbook);
                if (metadataSheet == null) return;

                SaveAllRecords(metadataSheet, new List<SnipRecord>());
                System.Diagnostics.Debug.WriteLine("Cleared all snip records");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error clearing snip records: {ex.Message}");
            }
        }

        public List<SnipRecord> GetRecordsByDocument(string documentName)
        {
            try
            {
                var allRecords = GetAllSnipRecords();
                return allRecords.Where(r => r.DocumentName?.Equals(documentName, StringComparison.OrdinalIgnoreCase) == true).ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting records by document: {ex.Message}");
                return new List<SnipRecord>();
            }
        }

        public List<SnipRecord> GetRecordsByMode(SnipMode mode)
        {
            try
            {
                var allRecords = GetAllSnipRecords();
                return allRecords.Where(r => r.Mode == mode).ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting records by mode: {ex.Message}");
                return new List<SnipRecord>();
            }
        }

        public void HighlightSnipCells()
        {
            try
            {
                var workbook = _application.ActiveWorkbook;
                if (workbook == null) return;

                var records = GetAllSnipRecords();
                
                foreach (var record in records)
                {
                    try
                    {
                        var range = (Range)workbook.Application.Range[record.CellAddress];
                        if (range != null)
                        {
                            // Set different colors based on snip mode
                            switch (record.Mode)
                            {
                                case SnipMode.Text:
                                    range.Interior.Color = ColorTranslator.ToOle(Color.LightBlue);
                                    break;
                                case SnipMode.Sum:
                                    range.Interior.Color = ColorTranslator.ToOle(Color.LightGreen);
                                    break;
                                case SnipMode.Table:
                                    range.Interior.Color = ColorTranslator.ToOle(Color.LightYellow);
                                    break;
                                case SnipMode.Validation:
                                    range.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                                    break;
                                case SnipMode.Exception:
                                    range.Interior.Color = ColorTranslator.ToOle(Color.LightPink);
                                    break;
                            }

                            // Add a comment with snip information
                            if (range.Comment == null)
                            {
                                var comment = range.AddComment($"Snip: {record.Mode}\nDocument: {record.DocumentName}\nPage: {record.PageNumber}\nCreated: {record.CreatedAt:yyyy-MM-dd HH:mm}");
                                comment.Visible = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error highlighting cell {record.CellAddress}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error highlighting snip cells: {ex.Message}");
            }
        }

        public void ClearHighlights()
        {
            try
            {
                var workbook = _application.ActiveWorkbook;
                if (workbook == null) return;

                var records = GetAllSnipRecords();
                
                foreach (var record in records)
                {
                    try
                    {
                        var range = (Range)workbook.Application.Range[record.CellAddress];
                        if (range != null)
                        {
                            range.Interior.ColorIndex = -4142; // No fill
                            
                            // Remove comment if it exists
                            if (range.Comment != null)
                            {
                                range.Comment.Delete();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error clearing highlight for cell {record.CellAddress}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error clearing highlights: {ex.Message}");
            }
        }

        private Worksheet GetOrCreateMetadataSheet(Workbook workbook)
        {
            Worksheet metadataSheet = null;
            try
            {
                // Try to get existing sheet
                metadataSheet = GetMetadataSheet(workbook);
                if (metadataSheet != null)
                {
                    return metadataSheet;
                }

                // If not found, create it
                metadataSheet = (Worksheet)workbook.Sheets.Add();
                metadataSheet.Name = METADATA_SHEET_NAME;
                metadataSheet.Visible = XlSheetVisibility.xlSheetVeryHidden; // Make it very hidden

                // Optional: Add headers or initial structure if needed
                // ((Range)metadataSheet.Cells[1, 1]).Value2 = "SnipDataJson";

                System.Diagnostics.Debug.WriteLine("Created metadata sheet.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting or creating metadata sheet: {ex.Message}");
                // Consider re-throwing or specific error handling
            }
            return metadataSheet;
        }

        private Worksheet GetMetadataSheet(Workbook workbook)
        {
            try
            {
                foreach (object sheetObj in workbook.Sheets)
                {
                    Worksheet sheet = sheetObj as Worksheet;
                    if (sheet != null && sheet.Name == METADATA_SHEET_NAME)
                    {
                        return sheet;
                    }
                }
                System.Diagnostics.Debug.WriteLine("Metadata sheet not found.");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting metadata sheet: {ex.Message}");
            }
            return null;
        }

        private List<SnipRecord> LoadAllRecords(Worksheet metadataSheet)
        {
            var records = new List<SnipRecord>();

            try
            {
                var lastRow = metadataSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                
                for (int row = 2; row <= lastRow; row++) // Start from row 2 (skip header)
                {
                    try
                    {
                        var cellAddress = GetCellValue(metadataSheet,row,1)?.ToString();
                        if (string.IsNullOrEmpty(cellAddress)) continue;

                        var record = new SnipRecord
                        {
                            CellAddress = cellAddress,
                            DocumentName = GetCellValue(metadataSheet,row,2)?.ToString() ?? "",
                            PageNumber = int.TryParse(GetCellValue(metadataSheet,row,3)?.ToString(), out var pageNum) ? pageNum : 1,
                            Mode = Enum.TryParse<SnipMode>(GetCellValue(metadataSheet,row,4)?.ToString(), out var mode) ? mode : SnipMode.Text,
                            ExtractedText = GetCellValue(metadataSheet,row,6)?.ToString() ?? "",
                            CreatedAt = DateTime.TryParse(GetCellValue(metadataSheet,row,7)?.ToString(), out var createdAt) ? createdAt : DateTime.Now
                        };

                        // Parse rectangle
                        var rectangleStr = GetCellValue(metadataSheet,row,5)?.ToString();
                        if (!string.IsNullOrEmpty(rectangleStr))
                        {
                            try
                            {
                                var parts = rectangleStr.Split(',');
                                if (parts.Length == 4)
                                {
                                    record.Rectangle = new Rectangle(
                                        int.Parse(parts[0]),
                                        int.Parse(parts[1]),
                                        int.Parse(parts[2]),
                                        int.Parse(parts[3])
                                    );
                                }
                            }
                            catch
                            {
                                record.Rectangle = Rectangle.Empty;
                            }
                        }

                        records.Add(record);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error loading record from row {row}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading records: {ex.Message}");
            }

            return records;
        }

        private void SaveAllRecords(Worksheet metadataSheet, List<SnipRecord> records)
        {
            try
            {
                // Clear existing data (except headers)
                var lastRow = metadataSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                if (lastRow > 1)
                {
                    var range = metadataSheet.Range[$"A2:H{lastRow}"];
                    range.Clear();
                }

                // Write new data
                for (int i = 0; i < records.Count; i++)
                {
                    var row = i + 2; // Start from row 2
                    var record = records[i];

                    metadataSheet.Cells[row, 1] = record.CellAddress;
                    metadataSheet.Cells[row, 2] = record.DocumentName;
                    metadataSheet.Cells[row, 3] = record.PageNumber;
                    metadataSheet.Cells[row, 4] = record.Mode.ToString();
                    metadataSheet.Cells[row, 5] = $"{record.Rectangle.X},{record.Rectangle.Y},{record.Rectangle.Width},{record.Rectangle.Height}";
                    metadataSheet.Cells[row, 6] = record.ExtractedText;
                    metadataSheet.Cells[row, 7] = record.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss");
                    metadataSheet.Cells[row, 8] = JsonConvert.SerializeObject(record, Formatting.None);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving records: {ex.Message}");
                throw;
            }
        }

        private static object GetCellValue(Worksheet sheet,int row,int col) => ((Range)sheet.Cells[row,col]).Value2;

        public void Dispose()
        {
            if (!_disposed)
            {
                _disposed = true;
            }
        }
    }
} 