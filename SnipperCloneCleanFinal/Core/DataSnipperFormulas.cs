using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using SnipperCloneCleanFinal.Infrastructure;

namespace SnipperCloneCleanFinal.Core
{
    public static class DataSnipperFormulas
    {
        // Dictionary storing all snips for the current workbook. It needs to be
        // reinitialised when loading a workbook, so it cannot be readonly.
        private static Dictionary<string, SnipData> _snipDatabase = new Dictionary<string, SnipData>();
        private static readonly JsonSerializerSettings JsonSettings = new JsonSerializerSettings
        {
            Formatting = Formatting.None,
            NullValueHandling = NullValueHandling.Ignore
        };
        
        public static string CreateTextFormula(string documentPath, int pageNumber, string extractedText, Rectangle bounds, string cellAddress, out string snipId)
        {
            snipId = Guid.NewGuid().ToString();
            var snipData = new SnipData
            {
                Id = snipId,
                Type = SnipMode.Text,
                DocumentPath = documentPath,
                PageNumber = pageNumber,
                ExtractedValue = extractedText,
                Bounds = bounds,
                Created = DateTime.Now,
                CellAddress = cellAddress
            };
            
            _snipDatabase[snipId] = snipData;
            return $"=SnipperPro.Connect.TEXTS(\"{snipId}\")";
        }
        
        public static string CreateSumFormula(string documentPath, int pageNumber, double sumValue, Rectangle bounds, List<double> numbers, string cellAddress, out string snipId)
        {
            snipId = Guid.NewGuid().ToString();
            var snipData = new SnipData
            {
                Id = snipId,
                Type = SnipMode.Sum,
                DocumentPath = documentPath,
                PageNumber = pageNumber,
                ExtractedValue = sumValue.ToString(),
                Bounds = bounds,
                Created = DateTime.Now,
                Numbers = numbers,
                CellAddress = cellAddress
            };

            _snipDatabase[snipId] = snipData;
            return $"=SnipperPro.Connect.SUMS(\"{snipId}\")";
        }

        public static string CreateTableFormula(string documentPath, int pageNumber, TableData table, Rectangle bounds, string cellAddress, out string snipId)
        {
            snipId = Guid.NewGuid().ToString();
            var snipData = new SnipData
            {
                Id = snipId,
                Type = SnipMode.Table,
                DocumentPath = documentPath,
                PageNumber = pageNumber,
                Bounds = bounds,
                Created = DateTime.Now,
                Table = table,
                CellAddress = cellAddress
            };

            _snipDatabase[snipId] = snipData;
            return $"=SnipperPro.Connect.TABLE(\"{snipId}\")";
        }
        
        public static string CreateValidationFormula(string documentPath, int pageNumber, Rectangle bounds, string cellAddress, out string snipId)
        {
            snipId = Guid.NewGuid().ToString();
            var snipData = new SnipData
            {
                Id = snipId,
                Type = SnipMode.Validation,
                DocumentPath = documentPath,
                PageNumber = pageNumber,
                ExtractedValue = "✓",
                Bounds = bounds,
                Created = DateTime.Now,
                CellAddress = cellAddress
            };
            
            _snipDatabase[snipId] = snipData;
            System.Diagnostics.Debug.WriteLine($"Created validation snip: ID={snipId}, Doc={documentPath}, Page={pageNumber}");
            return $"=SnipperPro.Connect.VALIDATION(\"{snipId}\")";
        }
        
        public static string CreateExceptionFormula(string documentPath, int pageNumber, Rectangle bounds, string cellAddress, out string snipId, string reason = "")
        {
            snipId = Guid.NewGuid().ToString();
            var snipData = new SnipData
            {
                Id = snipId,
                Type = SnipMode.Exception,
                DocumentPath = documentPath,
                PageNumber = pageNumber,
                ExtractedValue = "✗",
                Bounds = bounds,
                Created = DateTime.Now,
                ExceptionReason = reason,
                CellAddress = cellAddress
            };
            
            _snipDatabase[snipId] = snipData;
            System.Diagnostics.Debug.WriteLine($"Created exception snip: ID={snipId}, Doc={documentPath}, Page={pageNumber}");
            return $"=SnipperPro.Connect.EXCEPTION(\"{snipId}\")";
        }
        
        public static SnipData GetSnipData(string snipId)
        {
            return _snipDatabase.TryGetValue(snipId, out var data) ? data : null;
        }
        
        public static bool NavigateToSnip(string snipId)
        {
            System.Diagnostics.Debug.WriteLine($"NavigateToSnip called with ID: {snipId}");
            
            var snipData = GetSnipData(snipId);
            if (snipData == null) 
            {
                System.Diagnostics.Debug.WriteLine($"No snip data found for ID: {snipId}");
                return false;
            }
            
            System.Diagnostics.Debug.WriteLine($"Found snip data: Type={snipData.Type}, Document={snipData.DocumentPath}, Page={snipData.PageNumber}");
            
            try
            {
                // Get the main document viewer from the add-in
                System.Diagnostics.Debug.WriteLine("Getting add-in instance...");
                var addIn = SnipperCloneCleanFinal.ThisAddIn.Instance;
                
                if (addIn?.DocumentViewer != null && !addIn.DocumentViewer.IsDisposed)
                {
                    System.Diagnostics.Debug.WriteLine("Add-in and document viewer found");
                    
                    // Use the main document viewer for navigation
                    var viewer = addIn.DocumentViewer;
                    
                    // First check if document is already loaded to avoid reopening
                    System.Diagnostics.Debug.WriteLine($"Checking if document is already loaded: {snipData.DocumentPath}");
                    viewer.LoadDocumentIfNeeded(snipData.DocumentPath);
                    
                    // Navigate to the page and highlight the region
                    System.Diagnostics.Debug.WriteLine($"Navigating to page {snipData.PageNumber}");
                    viewer.NavigateToPage(snipData.PageNumber);
                    
                    System.Diagnostics.Debug.WriteLine($"Highlighting region: {snipData.Bounds}");
                    viewer.HighlightRegion(snipData.Bounds.ToDrawingRectangle(), GetSnipColor(snipData.Type));
                    
                    // Show and bring to front - but don't create multiple instances
                    System.Diagnostics.Debug.WriteLine("Ensuring viewer is visible");
                    if (viewer.WindowState == System.Windows.Forms.FormWindowState.Minimized)
                    {
                        viewer.WindowState = System.Windows.Forms.FormWindowState.Normal;
                    }
                    viewer.Show();
                    viewer.BringToFront();
                    viewer.Focus();
                    
                    Logger.Info($"Successfully navigated to snip {snipId} on page {snipData.PageNumber}");
                    System.Diagnostics.Debug.WriteLine("Navigation completed successfully");
                    return true;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"Document viewer not available - addIn: {addIn}, viewer: {addIn?.DocumentViewer}, disposed: {addIn?.DocumentViewer?.IsDisposed}");
                    Logger.Error("Document viewer not available for navigation");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to navigate to snip {snipId}: {ex.Message}", ex);
                System.Diagnostics.Debug.WriteLine($"Exception in NavigateToSnip: {ex.Message}");
                return false;
            }
        }
        
        private static System.Drawing.Color GetSnipColor(SnipMode type)
        {
            return type switch
            {
                SnipMode.Text => System.Drawing.Color.Blue,
                SnipMode.Sum => System.Drawing.Color.Purple,
                SnipMode.Table => System.Drawing.Color.Purple,
                SnipMode.Validation => System.Drawing.Color.Green,
                SnipMode.Exception => System.Drawing.Color.Red,
                _ => System.Drawing.Color.Gray
            };
        }

        public static void SaveSnips(Excel.Workbook workbook)
        {
            if (workbook == null) return;
            try
            {
                Office.DocumentProperties props = (Office.DocumentProperties)workbook.CustomDocumentProperties;
                string json = _snipDatabase.Count > 0 ? JsonConvert.SerializeObject(_snipDatabase, JsonSettings) : null;
                try
                {
                    var prop = props["SnipperProSnips"];
                    prop.Delete();
                }
                catch { }
                if (!string.IsNullOrEmpty(json))
                {
                    props.Add("SnipperProSnips", false, Office.MsoDocProperties.msoPropertyTypeString, json);
                }
            }
            catch { }
        }


        public static void LoadSnips(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                _snipDatabase.Clear();
                return;
            }
            try
            {
                Office.DocumentProperties props = (Office.DocumentProperties)workbook.CustomDocumentProperties;
                var prop = props["SnipperProSnips"];
                var json = prop.Value as string;
                _snipDatabase = string.IsNullOrEmpty(json)
                    ? new Dictionary<string, SnipData>()
                    : JsonConvert.DeserializeObject<Dictionary<string, SnipData>>(json) ?? new Dictionary<string, SnipData>();
            }
            catch
            {
                _snipDatabase = new Dictionary<string, SnipData>();
            }
        }


        public static void UpdateSnip(string snipId, SnipData updated)
        {
            if (_snipDatabase.ContainsKey(snipId))
            {
            _snipDatabase[snipId] = updated;
            }
        }
        
        public static bool DeleteSnip(string snipId)
        {
            try
            {
                // Get snip data before deleting
                var snipData = GetSnipData(snipId);
                if (snipData == null)
                {
                    System.Diagnostics.Debug.WriteLine($"No snip data found for deletion: {snipId}");
                    return false;
                }
                
                // Remove from database
                var removed = _snipDatabase.Remove(snipId);
                System.Diagnostics.Debug.WriteLine($"Snip {snipId} removed from database: {removed}");
                
                // Clear the Excel cell
                try
                {
                    var addIn = SnipperCloneCleanFinal.ThisAddIn.Instance;
                    if (addIn?.Application?.ActiveWorkbook != null && !string.IsNullOrEmpty(snipData.CellAddress))
                {
                        var workbook = addIn.Application.ActiveWorkbook;
                        
                        // Parse cell address to get sheet and cell reference
                        // Format could be "Sheet1!A1" or just "A1"
                        string sheetName = null;
                        string cellRef = snipData.CellAddress;
                        
                        if (snipData.CellAddress.Contains("!"))
                        {
                            var parts = snipData.CellAddress.Split('!');
                            sheetName = parts[0];
                            cellRef = parts[1];
                        }
                        
                        Excel.Worksheet worksheet;
                        if (!string.IsNullOrEmpty(sheetName))
                        {
                            worksheet = workbook.Worksheets[sheetName];
                        }
                        else
                        {
                            worksheet = workbook.ActiveSheet;
                        }
                        
                                                 var range = worksheet.Range[cellRef];
                         
                         // Complete cleanup - clear everything
                         range.ClearContents();  // Clear values and formulas
                         range.ClearComments();  // Clear any comments
                         range.ClearFormats();   // Clear formatting
                         range.ClearNotes();     // Clear notes
                         
                         // Also clear any hyperlinks
                         if (range.Hyperlinks.Count > 0)
                         {
                             range.Hyperlinks.Delete();
                         }
                         
                         System.Diagnostics.Debug.WriteLine($"Completely cleared Excel cell: {snipData.CellAddress}");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error($"Failed to clear Excel cell for snip {snipId}: {ex.Message}", ex);
                    System.Diagnostics.Debug.WriteLine($"Failed to clear Excel cell: {ex.Message}");
                }
                
                return removed;
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to delete snip {snipId}: {ex.Message}", ex);
                System.Diagnostics.Debug.WriteLine($"Exception in DeleteSnip: {ex.Message}");
                return false;
                }
            }
        
        public static Dictionary<string, SnipData> GetAllSnips()
        {
            return new Dictionary<string, SnipData>(_snipDatabase);
        }
    }
    
    public class SnipData
    {
        public string Id { get; set; }
        public SnipMode Type { get; set; }
        public string DocumentPath { get; set; }
        public int PageNumber { get; set; }
        public string ExtractedValue { get; set; }
        public Rectangle Bounds { get; set; }
        public DateTime Created { get; set; }
        public List<double> Numbers { get; set; } = new List<double>();
        public string ExceptionReason { get; set; }
        public TableData Table { get; set; }
        public string CellAddress { get; set; }
    }
}
